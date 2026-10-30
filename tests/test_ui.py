"""Проверка окна без экрана: все страницы открываются, записка создаётся из формы."""
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

try:
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QApplication
except ImportError:  # pragma: no cover
    QApplication = None

from sz.store import Store  # noqa: E402

from test_core import seed  # noqa: E402


@unittest.skipIf(QApplication is None, 'PyQt5 не установлен')
class WindowTest(unittest.TestCase):
    def test_pages_and_create(self):
        from sz.ui.common import STYLE
        from sz.ui.window import MainWindow

        root = tempfile.mkdtemp()
        store = Store(root)
        store.open()
        seed(store)
        store.close()

        app = QApplication.instance() or QApplication([])
        app.setStyleSheet(STYLE)
        win = MainWindow(lambda: root)
        win.show()
        self.assertTrue(win.connect())
        for page, arg in [('list', None), ('new', None), ('stats', None), ('dirs', 'addressees'),
                          ('dirs', 'settings')]:
            win.navigate(page, arg)
            app.processEvents()
            self.assertFalse(win.flash_label.isVisible(), win.flash_label.text())

        new = win.pages['new']
        # Адресаты используют то же многострочное поле, что и сотрудники.
        from PyQt5.QtWidgets import QPlainTextEdit
        win.navigate('dirs', 'addressees')
        tab = win.pages['dirs'].tab_widgets['addressees']
        tab.set_editing(tab.rows[0])
        position = tab.inputs['head_position']
        self.assertIsInstance(position, QPlainTextEdit)
        position.setPlainText('Начальнику\nотдела кадров')
        tab.save()
        tab.set_editing(tab.rows[0])
        self.assertEqual(position.toPlainText(), 'Начальнику\nотдела кадров')
        self.assertEqual(len(tab.rows), 1)
        win.navigate('new')
        new.topic.setCurrentIndex(1)
        new.executor.setCurrentIndex(1)
        new.addr_list.item(0).setCheckState(Qt.Checked)
        new.body.set_text('Прошу согласовать.')
        new.create()
        self.assertEqual(win.flash_label.objectName(), 'flashSuccess', win.flash_label.text())
        self.assertIs(win.stack.currentWidget(), win.pages['memo'])
        self.assertEqual(win.pages['memo'].reg.text(), 'НТЦ-0001-26')

        # Ошибка показывается сообщением, а не исключением
        win.navigate('new')
        new.create()
        self.assertEqual(win.flash_label.objectName(), 'flashError')

        # Изменение другим пользователем подтягивается в реестр
        win.navigate('list')
        other = Store(root)
        other.open()
        other.write(lambda s: s.run("UPDATE memos SET topic_name = 'Изменено'"))
        other.close()
        win.check_changes()
        self.assertEqual(win.pages['list'].table.item(0, 2).text(), 'Изменено')

        win.close()
        shutil.rmtree(root, ignore_errors=True)


if __name__ == '__main__':
    unittest.main()

