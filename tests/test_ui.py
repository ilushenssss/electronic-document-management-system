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
    def test_registry_page_navigation_and_filter_reset(self):
        from sz.ui.window import MainWindow
        from test_paging import fill_registry

        root = tempfile.TemporaryDirectory()
        self.addCleanup(root.cleanup)
        store = Store(root.name)
        store.open()
        fill_registry(store, 123)
        store.close()
        app = QApplication.instance() or QApplication([])
        win = MainWindow(lambda: root.name)
        self.addCleanup(win.close)
        self.assertTrue(win.connect())
        win.navigate('list')
        registry = win.pages['list']
        self.assertEqual(registry.table.rowCount(), 25)
        self.assertFalse(registry.prev_btn.isEnabled())
        registry.table.selectRow(0)
        registry.next_btn.click()
        self.assertEqual(registry.page, 1)
        self.assertIsNone(registry.selected_id())
        self.assertEqual(registry.table.item(0, 0).text(), 'НТЦ-0098-26')
        while registry.next_btn.isEnabled():
            registry.next_btn.click()
        self.assertEqual(registry.table.rowCount(), 23)
        self.assertFalse(registry.next_btn.isEnabled())
        registry.q.setText('оборудование')
        registry.reload()
        self.assertEqual(registry.page, 0)
        self.assertEqual(registry.table.rowCount(), 2)
        self.assertFalse(registry.next_btn.isEnabled())
        registry.reset_filters()
        self.assertEqual(registry.page, 0)
        registry.next_btn.click()
        registry.table.selectRow(0)
        registry.open_btn.click()
        self.assertEqual(win.pages['memo'].reg.text(), 'НТЦ-0098-26')
        win.navigate('list')
        self.assertEqual(registry.page, 1)
        while registry.next_btn.isEnabled():
            registry.next_btn.click()
        win.store.write(lambda s: s.run('DELETE FROM memos WHERE number <= 23'))
        registry.reload()
        self.assertEqual(registry.page, 3)
        self.assertEqual(registry.table.rowCount(), 25)
        app.processEvents()

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

        # Поиск исполнителя не сбрасывает уже сделанный выбор.
        win.navigate('new')
        new.topic.setCurrentIndex(1)
        new.executor.setCurrentIndex(1)
        new.executor_filter.setText('Нет такого сотрудника')
        self.assertTrue(new.executor.view().isRowHidden(1))
        self.assertEqual(new.executor.currentData(), 1)
        new.executor_filter.clear()
        self.assertFalse(new.executor.view().isRowHidden(1))
        new.executor_filter.setText('12-34')
        self.assertFalse(new.executor.view().isRowHidden(1))
        new.addr_list.item(0).setCheckState(Qt.Checked)
        new.addr_filter.setText('поиск')
        new.body.set_text('Несохранённый текст')
        new.attachment.setPlainText('Приложение')
        from PyQt5.QtCore import QDate
        new.date.setDate(QDate.currentDate().addDays(-5))
        new.clear_btn.click()
        self.assertIsNone(new.topic.currentData())
        self.assertIsNone(new.executor.currentData())
        self.assertEqual(new.chosen, [])
        self.assertEqual(new.addr_list.item(0).checkState(), Qt.Unchecked)
        self.assertEqual(new.addr_filter.text(), '')
        self.assertEqual(new.executor_filter.text(), '')
        self.assertEqual(new.body.plain_text(), '')
        self.assertEqual(new.attachment.toPlainText(), '')
        self.assertEqual(new.date.date(), QDate.currentDate())
        self.assertIsNotNone(new.signer)
        self.assertEqual(win.store.one('SELECT COUNT(*) AS n FROM memos')['n'], 1)

        # Ошибка показывается сообщением, а не исключением
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

        # Полный текст показывается над любой ячейкой, а в колонке — краткий.
        registry = win.pages['list']
        self.assertTrue(registry.table.showGrid())
        self.assertEqual(registry.table.horizontalHeaderItem(3).text(), 'Текст СЗ')
        self.assertEqual(registry.table.item(0, 3).text(), 'Прошу согласовать.')
        registry.executor_filter.setCurrentIndex(1)
        registry.addressee_filter.setCurrentIndex(1)
        self.assertEqual(registry.table.rowCount(), 1)
        self.assertEqual(registry.executor_filter.currentIndex(), 1)
        self.assertEqual(registry.addressee_filter.currentIndex(), 1)
        registry.q.setText('Ничего не найдётся')
        registry.reload()
        self.assertEqual(registry.table.rowCount(), 0)
        registry.reset_filters()
        self.assertEqual(registry.table.rowCount(), 1)
        self.assertIsNone(registry.executor_filter.currentData())
        self.assertIsNone(registry.addressee_filter.currentData())
        body = '<важно> & проверка\n' + 'Длинный текст записки. ' * 20
        win.store.write(lambda s: s.run('UPDATE memos SET body = ?', (body,)))
        registry.reload()
        compact = ' '.join(body.split())
        self.assertEqual(registry.table.item(0, 3).text(), compact[:200].rstrip() + '…')
        for col in range(registry.table.columnCount()):
            tooltip = registry.table.item(0, col).toolTip()
            self.assertIn('&lt;важно&gt; &amp; проверка<br>', tooltip)
            self.assertIn(body.split('\n')[1], tooltip)
        registry.table.selectRow(0)
        registry.table.cellDoubleClicked.emit(0, 3)
        self.assertIs(win.stack.currentWidget(), win.pages['memo'])
        win.navigate('list')
        registry.table.selectRow(0)
        registry.open_btn.click()
        self.assertIs(win.stack.currentWidget(), win.pages['memo'])

        win.close()
        shutil.rmtree(root, ignore_errors=True)


if __name__ == '__main__':
    unittest.main()

