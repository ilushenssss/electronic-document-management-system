"""Оформление текста записки: модель, документ Word, сохранение, панель кнопок."""
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from xml.dom import minidom

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from sz import actions, docx  # noqa: E402
from sz.store import Store  # noqa: E402

from test_core import memo_data, seed  # noqa: E402

try:
    from PyQt5.QtCore import QMimeData, Qt
    from PyQt5.QtGui import QTextCursor
    from PyQt5.QtWidgets import QApplication
except ImportError:  # pragma: no cover
    QApplication = None

RICH = [
    {'align': 'both', 'indent': True, 'runs': [
        {'text': 'Прошу '}, {'text': 'срочно', 'bold': True, 'underline': True},
        {'text': ' согласовать', 'italic': True, 'size': 12}]},
    {'align': 'center', 'indent': False, 'runs': [{'text': 'Первая строка\nвторая строка'}]},
]


def body_xml(memo):
    data = docx.build_docx(memo, [])
    xml = zipfile.ZipFile(io.BytesIO(data)).read('word/document.xml').decode('utf-8')
    minidom.parseString(xml)
    return xml


class RichModelTest(unittest.TestCase):
    def test_normalize(self):
        paragraphs = docx.normalize_rich([
            {'align': 'weird', 'runs': [{'text': '  а'}, {'text': 'б'}, {'text': 'в ', 'bold': 1},
                                        {'text': 'x', 'size': 99, 'color': 'red'}, {'text': 5}]},
            {'runs': [{'text': '   '}]},  # пустой абзац
            'мусор',
        ])
        self.assertEqual(paragraphs, [{'align': 'both', 'indent': False, 'runs': [
            {'text': 'аб'}, {'text': 'в ', 'bold': True}, {'text': 'x'}]}])

    def test_plain_needs_no_rich(self):
        plain = [{'align': 'both', 'indent': True, 'runs': [{'text': 'Абзац'}]}]
        self.assertIsNone(docx.dump_rich(plain))
        self.assertIsNone(docx.dump_rich([]))
        self.assertEqual(docx.parse_rich(docx.dump_rich(RICH)), RICH)
        self.assertIsNone(docx.parse_rich('{битый json'))
        self.assertEqual(docx.rich_plain_text(RICH), 'Прошу срочно согласовать\nПервая строка\nвторая строка')

    def test_word_xml(self):
        xml = body_xml({'reg_number': 'X', 'memo_date': '2026-01-01', 'body': 'игнорируется',
                        'body_rich': docx.dump_rich(RICH)})
        self.assertNotIn('игнорируется', xml)
        # Порядок элементов оформления - как требует схема Word
        self.assertIn('<w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/><w:u w:val="single"/></w:rPr>'
                      '<w:t xml:space="preserve">срочно</w:t>', xml)
        self.assertIn('<w:rPr><w:i/><w:iCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>', xml)
        self.assertIn('Первая строка</w:t><w:br/><w:t xml:space="preserve">вторая строка', xml)
        self.assertIn('<w:jc w:val="center"/>', xml)

    def test_left_align_has_no_jc(self):
        rich = [{'align': 'left', 'indent': False, 'runs': [{'text': 'Слева', 'bold': True}]}]
        xml = body_xml({'reg_number': 'X', 'body_rich': docx.dump_rich(rich)})
        p = xml[xml.rindex('<w:p>', 0, xml.index('Слева')):xml.index('Слева')]
        self.assertNotIn('<w:jc', p)
        self.assertNotIn('<w:ind', p)


class RichStoreTest(unittest.TestCase):
    def test_create_with_formatting(self):
        root = tempfile.mkdtemp()
        store = Store(root)
        store.open()
        seed(store)
        data = memo_data()
        data['body_rich'] = RICH
        memo_id, err = actions.create_memo(store, data, [1])
        self.assertIsNone(err)
        memo = store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        self.assertEqual(memo['body'], docx.rich_plain_text(RICH))
        self.assertEqual(json.loads(memo['body_rich']), RICH)
        self.assertEqual(len(actions.search_memos(store, 'срочно')), 1)
        with zipfile.ZipFile(store.abs_path(memo['file_path'])) as z:
            self.assertIn('<w:u w:val="single"/>', z.read('word/document.xml').decode('utf-8'))
        # Текст без оформления хранится как раньше
        data['body_rich'] = [{'align': 'both', 'indent': True, 'runs': [{'text': 'Просто'}]}]
        plain_id, _ = actions.create_memo(store, data, [1])
        plain = store.one('SELECT * FROM memos WHERE id = ?', (plain_id,))
        self.assertEqual((plain['body'], plain['body_rich']), ('Просто', None))
        store.close()
        shutil.rmtree(root, ignore_errors=True)


@unittest.skipIf(QApplication is None, 'PyQt5 не установлен')
class EditorTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def select(self, editor, start, length):
        cursor = editor.edit.textCursor()
        cursor.setPosition(start)
        cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, length)
        editor.edit.setTextCursor(cursor)

    def test_toolbar(self):
        from sz.ui.editor import FormattedTextEdit
        ed = FormattedTextEdit()
        ed.set_text('Прошу согласовать\nВторой абзац')
        self.assertIsNone(docx.dump_rich(ed.paragraphs()))  # без оформления

        self.select(ed, 6, 11)  # «согласовать»
        ed.bold.trigger()
        ed.italic.trigger()
        ed.size.setCurrentIndex(ed.size.findData(16))
        ed.set_size(ed.size.currentIndex())
        self.assertTrue(ed.bold.isChecked())

        cursor = ed.edit.textCursor()
        cursor.setPosition(len('Прошу согласовать\nВторой'))
        ed.edit.setTextCursor(cursor)
        ed.align_actions['center'].trigger()
        ed.indent.trigger()  # снять отступ
        ed.insert_break()

        self.assertEqual(ed.paragraphs(), [
            {'align': 'both', 'indent': True, 'runs': [
                {'text': 'Прошу '}, {'text': 'согласовать', 'bold': True, 'italic': True, 'size': 16}]},
            {'align': 'center', 'indent': False, 'runs': [{'text': 'Второй\n абзац'}]},
        ])

        # Кнопки показывают оформление под курсором
        self.select(ed, 8, 0)
        ed.sync()
        self.assertTrue(ed.bold.isChecked() and ed.italic.isChecked())
        self.assertEqual(ed.size.currentData(), 16)

        # «Сбросить» убирает оформление, но не перенос строки - это часть текста
        ed.edit.selectAll()
        ed.clear_format()
        self.assertEqual(ed.paragraphs(), [
            {'align': 'both', 'indent': True, 'runs': [{'text': 'Прошу согласовать'}]},
            {'align': 'both', 'indent': True, 'runs': [{'text': 'Второй\n абзац'}]},
        ])

    def test_paste_keeps_only_allowed_formatting(self):
        from sz.ui.editor import FormattedTextEdit
        ed = FormattedTextEdit()
        mime = QMimeData()
        mime.setHtml('<p style="font-family:Arial; font-size:20pt; color:red; text-align:center">'
                     'Из <b>Word</b> <i>с</i> <u>оформлением</u></p><p>второй</p>')
        ed.edit.insertFromMimeData(mime)
        self.assertEqual(ed.paragraphs(), [
            {'align': 'both', 'indent': True, 'runs': [
                {'text': 'Из '}, {'text': 'Word', 'bold': True}, {'text': ' '}, {'text': 'с', 'italic': True},
                {'text': ' '}, {'text': 'оформлением', 'underline': True}]},
            {'align': 'both', 'indent': True, 'runs': [{'text': 'второй'}]},
        ])

    def test_create_from_form(self):
        from sz.ui.window import MainWindow
        root = tempfile.mkdtemp()
        store = Store(root)
        store.open()
        seed(store)
        store.close()
        win = MainWindow(lambda: root)
        self.assertTrue(win.connect())
        win.navigate('new')
        page = win.pages['new']
        page.topic.setCurrentIndex(1)
        page.executor.setCurrentIndex(1)
        page.addr_list.item(0).setCheckState(Qt.Checked)
        page.body.set_text('Важно')
        page.body.edit.selectAll()
        page.body.bold.trigger()
        self.assertIn('<b>Важно</b>', _sheet_source(page))
        page.create()
        memo = win.store.one('SELECT * FROM memos')
        self.assertEqual(json.loads(memo['body_rich'])[0]['runs'], [{'text': 'Важно', 'bold': True}])
        self.assertEqual(page.body.plain_text(), '')  # форма очищена
        win.close()
        shutil.rmtree(root, ignore_errors=True)


def _sheet_source(page):
    """HTML, из которого построен лист предпросмотра."""
    from sz.ui.common import sheet_html
    d = page.form_data()
    return sheet_html(docx.memo_blocks({'reg_number': 'X', 'body_rich': docx.dump_rich(d['body_rich'])}, []))


if __name__ == '__main__':
    unittest.main()
