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
    def test_pdf_drop_number_dialog_and_independent_number_edit(self):
        from pathlib import Path
        from unittest.mock import patch
        from PyQt5.QtCore import QMimeData, QUrl, QPointF, QPoint
        from PyQt5.QtGui import QDropEvent, QDragEnterEvent, QDragMoveEvent
        from PyQt5.QtWidgets import QDialog
        from sz import actions, scans
        from sz.ui.pages import PdfDropArea
        from sz.ui.window import MainWindow
        from test_core import memo_data

        root = tempfile.TemporaryDirectory()
        self.addCleanup(root.cleanup)
        store = Store(root.name)
        store.open()
        seed(store)
        memo_id, error = actions.create_memo(store, memo_data(), [1])
        self.assertIsNone(error)
        store.close()
        app = QApplication.instance() or QApplication([])
        win = MainWindow(lambda: root.name)
        self.addCleanup(win.close)
        self.assertTrue(win.connect())
        win.navigate('memo', memo_id)
        page = win.pages['memo']
        source = Path(root.name) / 'scan.pdf'
        source.write_bytes(b'%PDF-1.4\n%%EOF\n')
        with patch('sz.ui.pages.SedNumberDialog') as dialog:
            dialog.return_value.exec_.return_value = QDialog.Rejected
            page.upload_scan(str(source))
        self.assertEqual(win.store.one('SELECT sed_number FROM memos')['sed_number'], '')
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(source))])
        self.assertEqual(PdfDropArea.local_pdf(mime), str(source))
        for target in (page.drop_area, page):
            enter = QDragEnterEvent(QPoint(10, 10), Qt.CopyAction | Qt.MoveAction,
                                    mime, Qt.LeftButton, Qt.NoModifier)
            target.dragEnterEvent(enter)
            self.assertTrue(enter.isAccepted())
            move = QDragMoveEvent(QPoint(10, 10), Qt.CopyAction | Qt.MoveAction,
                                  mime, Qt.LeftButton, Qt.NoModifier)
            target.dragMoveEvent(move)
            self.assertTrue(move.isAccepted())
            self.assertEqual(move.dropAction(), Qt.CopyAction)
        self.assertFalse(page.sed_number.acceptDrops())
        self.assertFalse(page.sheet.acceptDrops())
        self.assertFalse(page.sheet.viewport().acceptDrops())
        with patch('sz.ui.pages.SedNumberDialog') as dialog:
            dialog.return_value.exec_.return_value = QDialog.Accepted
            dialog.return_value.value = '1/441062/1100/СЗ'
            dialog.return_value.in_paper = False
            event = QDropEvent(QPointF(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
            page.dropEvent(event)
        self.assertTrue(page.open_scan_btn.isEnabled())
        self.assertEqual(page.sed_number.text(), '1/441062/1100/СЗ')
        self.assertTrue(source.exists())
        page.sed_number.setText('1/441062/1101/СЗ')
        page.save_sed_btn.click()
        memo = win.store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        self.assertTrue(scans.document_path(win.store, memo).endswith('1_441062_1101_СЗ.pdf'))
        page.sed_number.setText('1/441062/1102/СЗ')
        page.sed_number.textEdited.emit(page.sed_number.text())
        win.confirm = lambda message: False
        win.navigate('list')
        self.assertIs(win.stack.currentWidget(), page)
        page.on_external_change()
        self.assertEqual(page.sed_number.text(), '1/441062/1102/СЗ')
        page.sed_number.setText(memo['sed_number'])
        win.confirm = lambda message: True
        win.navigate('list')
        registry = win.pages['list']
        self.assertEqual(registry.table.horizontalHeaderItem(1).text(), 'Номер СЭД')
        self.assertEqual(registry.table.item(0, 1).text(), '1/441062/1101/СЗ')
        registry.q.setText('441062/1101/сз')
        registry.reload()
        self.assertEqual(registry.table.rowCount(), 1)
        mime.setUrls([QUrl.fromLocalFile(str(source)), QUrl.fromLocalFile(str(source))])
        self.assertIsNone(PdfDropArea.local_pdf(mime))
        mime.setUrls([QUrl('https://example.com/scan.pdf')])
        self.assertIsNone(PdfDropArea.local_pdf(mime))
        # Outlook передаёт FileContents без пути на диске. Временный файл
        # должен дожить до сохранения, затем удалиться.
        from test_scans import FakeOutlookMime
        win.navigate('memo', memo_id)
        virtual = QMimeData()
        for fmt, content in FakeOutlookMime().formats.items():
            virtual.setData(fmt, content)
        with patch('sz.ui.pages.SedNumberDialog') as dialog, \
                patch('sz.ui.pages.scans.check_pdf', wraps=scans.check_pdf) as checked:
            dialog.return_value.exec_.return_value = QDialog.Accepted
            dialog.return_value.value = '1/441062/1103/СЗ'
            dialog.return_value.in_paper = False
            event = QDropEvent(QPointF(10, 10), Qt.CopyAction, virtual, Qt.LeftButton, Qt.NoModifier)
            page.dropEvent(event)
            temporary = [Path(call.args[0]) for call in checked.call_args_list
                         if Path(call.args[0]).name == 'attachment.pdf']
        self.assertTrue(temporary)
        self.assertTrue(all(not path.exists() for path in temporary))
        self.assertEqual(page.sed_number.text(), '1/441062/1103/СЗ')
        self.assertTrue(page.open_scan_btn.isEnabled())
        from sz.ui.pages import SedNumberDialog
        number_dialog = SedNumberDialog(page, 'не номер', page.memo)
        number_dialog.paper.setChecked(True)
        self.assertFalse(number_dialog.number.isEnabled())
        self.assertIn('НТЦ-0001-26.pdf', number_dialog.filename.text())
        number_dialog.validate()
        self.assertEqual(number_dialog.result(), QDialog.Accepted)
        self.assertTrue(number_dialog.in_paper)
        self.assertEqual(number_dialog.value, '')
        page.paper.setChecked(True)
        page.save_sed_btn.click()
        self.assertFalse(page.sed_number.isEnabled())
        self.assertTrue(scans.document_path(win.store, page.memo).endswith('НТЦ-0001-26.pdf'))
        win.navigate('list')
        registry.reset_filters()
        self.assertEqual(registry.table.item(0, 1).text(), 'В бумаге')
        registry.missing_sed.setChecked(True)
        self.assertEqual(registry.page, 0)
        self.assertEqual(registry.table.rowCount(), 0)
        registry.reset_filters()
        self.assertFalse(registry.missing_sed.isChecked())
        self.assertEqual(registry.table.rowCount(), 1)
        app.processEvents()

    def test_edit_memo_locked_word_retry_and_conflict(self):
        from pathlib import Path
        from sz import actions
        from sz.ui.window import MainWindow
        from test_core import memo_data

        root = tempfile.TemporaryDirectory()
        self.addCleanup(root.cleanup)
        store = Store(root.name)
        store.open()
        seed(store)
        memo_id, error = actions.create_memo(store, memo_data(body='Исходный текст'), [1])
        self.assertIsNone(error)
        store.close()
        app = QApplication.instance() or QApplication([])
        win = MainWindow(lambda: root.name)
        win.confirm = lambda message: True
        self.addCleanup(win.close)
        self.assertTrue(win.connect())
        win.navigate('memo', memo_id)
        win.pages['memo'].edit_btn.click()
        edit = win.pages['edit']
        self.assertIs(win.stack.currentWidget(), edit)
        self.assertEqual(edit.body.plain_text(), 'Исходный текст')
        self.assertEqual(edit.topic.currentData(), -1)
        self.assertEqual(edit.executor.currentData(), -1)
        self.assertEqual(edit.chosen, [-1])
        self.assertFalse(edit.has_unsaved_changes())
        edit.body.set_text('Исправленный текст')
        path = Path(win.store.document_path(edit.edit_snapshot['memo']['file_path']))
        owner = path.with_name('~$' + path.name)
        owner.touch()
        edit.submit.click()
        self.assertIs(win.stack.currentWidget(), edit)
        self.assertIn('Word не обновлён', win.flash_label.text())
        self.assertEqual(edit.body.plain_text(), 'Исправленный текст')
        self.assertTrue(edit.pending_document)
        owner.unlink()
        edit.submit.click()
        self.assertIs(win.stack.currentWidget(), win.pages['memo'])
        self.assertEqual(win.pages['memo'].reg.text(), 'НТЦ-0001-26')
        win.pages['memo'].edit_btn.click()
        edit.body.set_text('Мой несохранённый текст')
        # Обновления других пользователей не затирают открытую форму.
        win.store.write(lambda s: s.run('UPDATE memos SET body = ? WHERE id = ?', ('Правка коллеги', memo_id)))
        edit.on_external_change()
        self.assertEqual(edit.body.plain_text(), 'Мой несохранённый текст')
        edit.submit.click()
        self.assertFalse(edit.submit.isEnabled())
        self.assertFalse(edit.reload_saved_btn.isHidden())
        edit.reload_saved_btn.click()
        self.assertEqual(edit.body.plain_text(), 'Правка коллеги')
        self.assertIn('Мой несохранённый текст', edit.previous_draft)
        edit.body.set_text('Не терять при уходе')
        win.confirm = lambda message: False
        win.navigate('list')
        self.assertIs(win.stack.currentWidget(), edit)
        win.confirm = lambda message: True
        app.processEvents()

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
        self.assertEqual(win.pages['list'].table.item(0, 3).text(), 'Изменено')

        # Полный текст показывается над любой ячейкой, а в колонке — краткий.
        registry = win.pages['list']
        self.assertTrue(registry.table.showGrid())
        self.assertEqual(registry.table.horizontalHeaderItem(4).text(), 'Текст СЗ')
        self.assertEqual(registry.table.item(0, 4).text(), 'Прошу согласовать.')
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
        self.assertEqual(registry.table.item(0, 4).text(), compact[:200].rstrip() + '…')
        for col in range(registry.table.columnCount()):
            tooltip = registry.table.item(0, col).toolTip()
            self.assertIn('&lt;важно&gt; &amp; проверка<br>', tooltip)
            self.assertIn(body.split('\n')[1], tooltip)
        registry.table.selectRow(0)
        registry.table.cellDoubleClicked.emit(0, 4)
        self.assertIs(win.stack.currentWidget(), win.pages['memo'])
        win.navigate('list')
        registry.table.selectRow(0)
        registry.open_btn.click()
        self.assertIs(win.stack.currentWidget(), win.pages['memo'])

        win.close()
        shutil.rmtree(root, ignore_errors=True)


if __name__ == '__main__':
    unittest.main()

