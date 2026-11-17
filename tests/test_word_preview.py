"""Контракт рендерера Word и асинхронное отображение страниц."""
import base64
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch
import zipfile

from sz import docx, word_preview

PNG = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAABAAAAAYCAIAAAB8wupbAAAAHUlEQVR4nGP8//8/AymAiSTVoxpGNYxqGNVAsgYAJ+gDLaR+vkEAAAAASUVORK5CYII=')


class PreviewContractTest(unittest.TestCase):
    def test_digest_ignores_zip_timestamps_but_detects_document_changes(self):
        def package(text, year):
            result = io.BytesIO()
            with zipfile.ZipFile(result, 'w') as archive:
                info = zipfile.ZipInfo('word/document.xml', (year, 1, 1, 0, 0, 0))
                archive.writestr(info, text)
            return result.getvalue()
        self.assertEqual(word_preview.document_digest(package('Текст', 2020)),
                         word_preview.document_digest(package('Текст', 2026)))
        self.assertNotEqual(word_preview.document_digest(package('Текст', 2020)),
                            word_preview.document_digest(package('Другая СЗ', 2020)))

    def test_manifest_pages_errors_and_paths(self):
        with tempfile.TemporaryDirectory() as root:
            root = Path(root)
            (root / 'page-0001.png').write_bytes(PNG)
            (root / 'result.json').write_text(json.dumps({'pages': ['page-0001.png']}))
            self.assertEqual(word_preview.read_pages(root), [root / 'page-0001.png'])
            for result in ({'pages': []}, {'error': 'Word unavailable'}, {'pages': ['../page.png']},
                           {'pages': ['missing.png']}):
                (root / 'result.json').write_text(json.dumps(result))
                with self.assertRaises(RuntimeError):
                    word_preview.read_pages(root)

    def test_non_windows_error_is_explicit(self):
        with patch.object(word_preview.sys, 'platform', 'linux'):
            with self.assertRaisesRegex(RuntimeError, 'Microsoft Word'):
                word_preview.powershell_command()


try:
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    from PyQt5.QtWidgets import QApplication
except ImportError:
    QApplication = None


@unittest.skipIf(QApplication is None, 'PyQt5 не установлен')
class PreviewUiTest(unittest.TestCase):
    def setUp(self):
        from sz.ui.preview import WordPreview
        self.app = QApplication.instance() or QApplication([])
        self.view = WordPreview()
        self.view.resize(650, 700)
        self.view.show()
        self.app.processEvents()

    def tearDown(self):
        self.view.close()
        self.view.service.shutdown()
        if hasattr(self.app, '_sz_word_renderer'):
            del self.app._sz_word_renderer
        self.app.processEvents()

    def wait_for(self, condition):
        end = time.monotonic() + 5
        while not condition() and time.monotonic() < end:
            self.app.processEvents()
            time.sleep(0.01)
        self.assertTrue(condition())

    def fake_command(self):
        script = ('import os,json,base64;from pathlib import Path;'
                  'p=Path(os.environ["SZ_PREVIEW_DIRECTORY"]);'
                  '(p/"page-0001.png").write_bytes(base64.b64decode(%r));'
                  '(p/"result.json").write_text(json.dumps({"pages":["page-0001.png"]}))'
                  % base64.b64encode(PNG).decode('ascii'))
        return [sys.executable, '-c', script]

    def test_async_success_zoom_and_saved_document_source(self):
        with tempfile.TemporaryDirectory() as root, patch.object(word_preview, 'powershell_command', self.fake_command):
            saved = docx.build_docx({'body': 'Сохранённый Word'}, [])
            source = Path(root) / 'saved.docx'
            source.write_bytes(saved)
            self.view.show_memo({'reg_number': 'НТЦ-0001-26', 'body': 'Другой текст БД'}, [], document_path=str(source))
            self.view.request()
            self.wait_for(lambda: bool(self.view.images))
            self.assertEqual(self.view.data, saved)
            self.assertIn('Страниц: 1', self.view.status.text())
            self.view.zoom.setCurrentIndex(2)
            self.view.scale_pages()
            self.assertTrue(self.view.images)

    def test_stale_result_does_not_replace_current_preview(self):
        self.view.digest = 'new'
        before = self.view.status.text()
        self.view.completed('old', [], 'старый результат')
        self.assertEqual(self.view.status.text(), before)

    def test_repeated_requests_use_one_render_and_then_cache(self):
        with patch.object(word_preview, 'powershell_command', side_effect=self.fake_command) as command:
            self.view.show_memo({'body': 'Один документ'}, [])
            self.view.request()
            self.view.request()
            self.wait_for(lambda: bool(self.view.images))
            self.view.request()
            self.assertEqual(command.call_count, 1)

    def test_two_pages_do_not_overlap_and_fit_without_horizontal_scroll(self):
        with tempfile.TemporaryDirectory() as root:
            page = Path(root) / 'page.png'
            page.write_bytes(PNG)
            self.view.digest = 'two-pages'
            self.view.completed('two-pages', [page, page], None)
            for _ in range(4):
                self.app.processEvents()
            first = self.view.images[0][0].geometry()
            second = self.view.images[1][0].geometry()
            self.assertLess(first.bottom(), second.top())
            self.assertGreater(self.view.scroll.verticalScrollBar().maximum(), 0)
            self.assertEqual(self.view.scroll.horizontalScrollBar().maximum(), 0)

    def test_unavailable_word_and_clear(self):
        with patch.object(word_preview, 'powershell_command', side_effect=RuntimeError('Нет Microsoft Word')):
            self.view.show_memo({'body': 'Текст'}, [])
            self.view.request()
            self.wait_for(lambda: 'Нет Microsoft Word' in self.view.status.text())
        self.assertFalse(self.view.images)
        self.view.clear()
        self.assertIsNone(self.view.digest)
        self.assertFalse(self.view.timer.isActive())


@unittest.skipUnless(sys.platform == 'win32' and os.environ.get('SZ_TEST_WORD_PREVIEW') == '1',
                     'Нужны Windows, Microsoft Word и SZ_TEST_WORD_PREVIEW=1')
class NativeWordTest(unittest.TestCase):
    def test_native_word_returns_pages(self):
        with tempfile.TemporaryDirectory() as root:
            root = Path(root)
            memo = {'body': '\n'.join(['Проверка разбиения на страницы.'] * 100),
                    'executor_name': 'Иванов Иван Иванович', 'executor_phone': '12-34'}
            (root / 'input.docx').write_bytes(docx.build_docx(memo, []))
            env = os.environ.copy()
            env['SZ_PREVIEW_DIRECTORY'] = str(root)
            try:
                subprocess.run(word_preview.powershell_command(), env=env, timeout=100, check=True)
                self.assertGreater(len(word_preview.read_pages(root)), 1)
            finally:
                word_preview.stop_owned_word(root)


if __name__ == '__main__':
    unittest.main()
