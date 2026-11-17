"""Шаблоны: копирование текста, конфликты общей БД и реальные Qt-формы."""
import os
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from sz import actions, docx, templates
from sz.store import Store, StoreError
from test_core import seed, memo_data


def content(text):
    return [{'runs': [{'text': line}]} for line in text.splitlines()]


def data(name='Закупка', text='Прошу {наименование}.', kind='memo', **extra):
    return dict(name=name, kind=kind, paragraphs=content(text), **extra)


class TemplateTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.store = Store(self.tmp.name)
        self.store.open()
        self.addCleanup(self.store.close)
        seed(self.store)

    def test_crud_filters_and_topic_detachment(self):
        s = self.store
        actions.save_directory_row(s, 'topics', {'name': 'Другая'})
        first = templates.save_template(s, data(topic_id=1))
        second = templates.save_template(s, data('Общий', kind='fragment'))
        third = templates.save_template(s, data('Неактивный', active=False))
        templates.save_template(s, data('Другая тема', topic_id=2))
        self.assertEqual({r['id'] for r in templates.list_templates(s, True, 1)}, {first, second})
        self.assertEqual(len(templates.list_templates(s, True)), 3)
        row = templates.get_template(s, first)
        templates.save_template(s, data('Исправленный', text='Новый'), first, row['version'])
        with self.assertRaises(StoreError):
            templates.delete_template(s, first, row['version'])
        current = templates.get_template(s, first)
        templates.delete_template(s, first, current['version'])
        with self.assertRaises(StoreError):
            templates.get_template(s, first)
        self.assertFalse(templates.get_template(s, third)['active'])
        linked = templates.save_template(s, data('Связанный', topic_id=2))
        actions.delete_directory_row(s, 'topics', 2)
        self.assertIsNone(templates.get_template(s, linked)['topic_id'])
        self.assertEqual(templates.get_template(s, linked)['version'], 2)

    def test_validation_and_conflict_with_second_connection(self):
        s = self.store
        for invalid in (data(name='  '), data(text='  '), data(kind='other'), data(topic_id=999)):
            with self.assertRaises(StoreError):
                templates.save_template(s, invalid)
        row_id = templates.save_template(s, data())
        with self.assertRaises(StoreError):
            templates.save_template(s, data('ЗАКУПКА'))
        other = Store(self.tmp.name)
        other.open()
        self.addCleanup(other.close)
        templates.save_template(other, data(text='Чужое изменение'), row_id, 1)
        with self.assertRaises(StoreError):
            templates.save_template(s, data(text='Мой текст'), row_id, 1)
        self.assertEqual(templates.get_template(s, row_id)['body'], 'Чужое изменение')
        templates.delete_template(other, row_id, 2)
        with self.assertRaises(StoreError):
            templates.save_template(s, data(), row_id, 2)

    def test_rich_placeholder_across_runs_and_literal_replacement(self):
        source = [{'align': 'center', 'indent': False, 'runs': [
            {'text': 'До {На', 'bold': True}, {'text': 'именование}', 'italic': True},
            {'text': ' и {наименование}.', 'underline': True}]}]
        self.assertEqual(templates.fields(source), {'наименование': 'Наименование'})
        result = templates.render(source, {'Наименование': '{дата}'})
        self.assertEqual(docx.rich_plain_text(result), 'До {дата} и {дата}.')
        self.assertEqual(result[0]['align'], 'center')
        self.assertFalse(result[0]['indent'])
        self.assertTrue(result[0]['runs'][0]['bold'])
        self.assertTrue(result[0]['runs'][-1]['underline'])
        self.assertEqual(source[0]['runs'][1]['text'], 'именование}')
        with self.assertRaises(StoreError):
            templates.render(source, {'наименование': ' '})
        literal = content('{{образец}} {x.y} {__import__(x)}')
        self.assertEqual(templates.fields(literal), {})
        self.assertEqual(docx.rich_plain_text(templates.render(literal, {})),
                         '{{образец}} {x.y} {__import__(x)}')

    def test_examples_are_opt_in_and_memos_are_independent(self):
        s = self.store
        self.assertEqual(templates.list_templates(s), [])
        self.assertEqual(templates.add_examples(s), 3)
        self.assertEqual(templates.add_examples(s), 0)
        row = templates.list_templates(s)[0]
        rich = templates.render(templates.paragraphs(row),
                                {key: 'Значение' for key in templates.fields(templates.paragraphs(row))})
        raw = memo_data(body=docx.rich_plain_text(rich))
        raw['body_rich'] = rich
        memo_id, error = actions.create_memo(s, raw, [1])
        self.assertIsNone(error)
        before = s.one('SELECT * FROM memos WHERE id=?', (memo_id,))
        templates.save_template(s, data('Изменённый', 'Совсем другой текст'), row['id'], row['version'])
        templates.delete_template(s, row['id'], row['version'] + 1)
        self.assertEqual(s.one('SELECT * FROM memos WHERE id=?', (memo_id,)), before)
        with zipfile.ZipFile(s.document_path(before['file_path'])) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertIn('Значение', xml)
        self.assertNotIn('Совсем другой текст', xml)


try:
    from PyQt5.QtGui import QTextCursor
    from PyQt5.QtWidgets import QApplication, QDialog
except ImportError:
    QApplication = None


@unittest.skipIf(QApplication is None, 'PyQt5 не установлен')
class TemplatesUiTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        from sz.ui.window import MainWindow
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        store = Store(self.tmp.name)
        store.open()
        seed(store)
        store.close()
        self.win = MainWindow(lambda: self.tmp.name)
        self.assertTrue(self.win.connect())
        self.addCleanup(self.finish)

    def finish(self):
        self.win.pages['dirs'].discard_changes()
        with patch.object(self.win, 'confirm', return_value=True):
            self.win.close()

    def add(self, **kwargs):
        return templates.save_template(self.win.store, data(**kwargs))

    def picker(self, row_id, page=None):
        page = page or self.win.pages['new']
        page.reload()
        page.template_picker.setCurrentIndex(page.template_picker.findData(row_id))
        return page

    def test_fragment_cursor_selection_rich_and_single_undo(self):
        from sz.ui.editor import FormattedTextEdit
        editor = FormattedTextEdit()
        editor.set_text('До  после')
        cursor = editor.edit.textCursor()
        cursor.setPosition(3)
        editor.edit.setTextCursor(cursor)
        rich = [{'runs': [{'text': 'ВСТАВКА ', 'bold': True}]}]
        editor.insert_paragraphs(rich)
        self.assertEqual(editor.plain_text(), 'До ВСТАВКА после')
        self.assertTrue(any(run.get('bold') for run in editor.paragraphs()[0]['runs']))
        editor.edit.undo()
        self.assertEqual(editor.plain_text(), 'До  после')
        cursor = editor.edit.textCursor()
        cursor.setPosition(3)
        cursor.movePosition(QTextCursor.End, QTextCursor.KeepAnchor)
        editor.edit.setTextCursor(cursor)
        editor.insert_paragraphs(content('Первая\nВторая'))
        self.assertEqual(editor.plain_text(), 'До Первая\nВторая')
        editor.edit.undo()
        self.assertEqual(editor.plain_text(), 'До  после')
        editor.insert_paragraphs(content('Полная замена'), replace=True)
        self.assertEqual(editor.plain_text(), 'Полная замена')
        editor.edit.undo()
        self.assertEqual(editor.plain_text(), 'До  после')

    def test_cancel_and_insert_with_known_defaults(self):
        from sz.ui.templates import TemplateFieldsDialog
        row_id = self.add(text='Тема {тематика}, дата {дата}, автор {исполнитель}, {цель}.')
        self.win.navigate('new')
        page = self.picker(row_id)
        page.topic.setCurrentIndex(page.topic.findData(1))
        page.executor.setCurrentIndex(page.executor.findData(1))
        page.body.set_text('Мой текст')
        with patch.object(self.win, 'confirm', return_value=False):
            page.insert_template()
        self.assertEqual(page.body.plain_text(), 'Мой текст')
        with patch.object(self.win, 'confirm', return_value=True), \
                patch.object(TemplateFieldsDialog, 'exec_', return_value=QDialog.Rejected):
            page.insert_template()
        self.assertEqual(page.body.plain_text(), 'Мой текст')
        captured = {}

        def accept(dialog):
            captured.update(dialog.values())
            dialog.inputs['цель'].setText('проверка')
            return QDialog.Accepted

        with patch.object(self.win, 'confirm', return_value=True), \
                patch.object(TemplateFieldsDialog, 'exec_', accept):
            page.insert_template()
        self.assertEqual(captured['тематика'], 'Закупка')
        self.assertEqual(captured['дата'], page.date.date().toString('dd.MM.yyyy'))
        self.assertEqual(captured['исполнитель'], 'Сидоров Сидор Сидорович')
        self.assertIn('проверка', page.body.plain_text())
        self.assertNotIn('{', page.body.plain_text())
        page.body.edit.undo()
        self.assertEqual(page.body.plain_text(), 'Мой текст')

    def test_directory_save_unsaved_guard_conflict_and_delete(self):
        self.win.navigate('dirs', 'templates')
        dirs = self.win.pages['dirs']
        tab = dirs.current()
        self.assertFalse(tab.has_unsaved_changes())
        self.assertFalse(tab.edit_btn.isEnabled())
        tab.name.setText('Мой шаблон')
        tab.body.set_text('Текст {цель}')
        with patch.object(self.win, 'confirm', return_value=False):
            dirs.tabs.setCurrentIndex(0)
            self.assertIs(dirs.current(), tab)
            self.win.navigate('list')
            self.assertIs(self.win.stack.currentWidget(), dirs)
        tab.save()
        self.assertFalse(tab.has_unsaved_changes())
        self.assertEqual(len(tab.rows), 1)
        tab.table.selectRow(0)
        tab.edit_selected()
        tab.body.set_text('Мой новый текст')
        row = templates.get_template(self.win.store, tab.edit_id)
        templates.save_template(self.win.store, data('Мой шаблон', 'Изменено извне'), row['id'], row['version'])
        tab.reload()
        self.assertEqual(tab.body.plain_text(), 'Мой новый текст')
        with patch.object(self.win, 'report') as report:
            tab.save()
            self.assertTrue(report.called)
        self.assertEqual(tab.body.plain_text(), 'Мой новый текст')
        self.assertEqual(templates.get_template(self.win.store, row['id'])['body'], 'Изменено извне')
        with patch.object(self.win, 'confirm', return_value=True):
            tab.edit_selected()
            self.assertEqual(tab.body.plain_text(), 'Изменено извне')
            tab.delete_selected()
        self.assertEqual(tab.rows, [])

    def test_inactive_and_deleted_template_and_edit_form(self):
        row_id = self.add(text='Фрагмент', kind='fragment')
        self.win.navigate('new')
        page = self.picker(row_id)
        row = templates.get_template(self.win.store, row_id)
        templates.save_template(self.win.store, data(text='Отключённый', active=False), row_id, row['version'])
        page.insert_template()
        self.assertEqual(page.body.plain_text(), '')
        self.assertEqual(page.template_picker.findData(row_id), -1)
        other = self.add(name='Другой', text='Без переменных', kind='fragment')
        self.picker(other, page)
        templates.delete_template(self.win.store, other, 1)
        with patch.object(self.win, 'report'):
            page.insert_template()
        self.assertEqual(page.body.plain_text(), '')
        memo_id, error = actions.create_memo(self.win.store, memo_data(body='Было'), [1])
        self.assertIsNone(error)
        fragment = self.add(name='Редактирование', text=' Стало', kind='fragment')
        self.win.navigate('edit', memo_id)
        edit = self.win.pages['edit']
        edit.template_picker.setCurrentIndex(edit.template_picker.findData(fragment))
        edit.insert_template()
        self.assertIn('Стало', edit.body.plain_text())
        self.assertTrue(edit.has_unsaved_changes())
        self.assertEqual(self.win.store.one('SELECT body FROM memos WHERE id=?', (memo_id,))['body'], 'Было')

    def test_template_to_memo_word_and_clear(self):
        from sz.ui.templates import TemplateFieldsDialog
        row_id = templates.save_template(self.win.store, dict(
            name='Оформленный', kind='memo', topic_id=1,
            paragraphs=[{'align': 'left', 'indent': True, 'runs': [
                {'text': 'Прошу {цель}.', 'bold': True, 'size': 12}]}]))
        self.win.navigate('new')
        page = self.picker(row_id)
        page.executor.setCurrentIndex(page.executor.findData(1))
        page.chosen = [1]

        def fill(dialog):
            dialog.inputs['цель'].setText('проверить документ')
            return QDialog.Accepted

        with patch.object(TemplateFieldsDialog, 'exec_', fill):
            page.insert_template()
        self.assertEqual(page.topic.currentData(), 1)
        self.assertEqual(page.body.plain_text(), 'Прошу проверить документ.')
        page.create()
        memo = self.win.store.one('SELECT * FROM memos')
        self.assertEqual(memo['body'], 'Прошу проверить документ.')
        rich = docx.parse_rich(memo['body_rich'])
        self.assertTrue(rich[0]['runs'][0]['bold'])
        self.assertEqual(rich[0]['runs'][0]['size'], 12)
        with zipfile.ZipFile(self.win.store.document_path(memo['file_path'])) as archive:
            xml = archive.read('word/document.xml').decode()
        self.assertIn('Прошу проверить документ.', xml)
        self.assertNotIn('{цель}', xml)
        self.win.navigate('new')
        self.assertIsNone(page.template_picker.currentData())
        page.template_picker.setCurrentIndex(page.template_picker.findData(row_id))
        page.clear_form()
        self.assertIsNone(page.template_picker.currentData())
        self.assertEqual(page.body.plain_text(), '')

    def test_keyboard_popup_and_unsaved_close_and_discard(self):
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QCloseEvent
        from PyQt5.QtTest import QTest
        self.add(text='Текст')
        self.win.show()
        self.win.navigate('new')
        page = self.win.pages['new']
        page.template_picker.setFocus()
        page.template_picker.showPopup()
        self.app.processEvents()
        QTest.keyClick(page.template_picker, Qt.Key_Down)
        QTest.keyClick(page.template_picker, Qt.Key_Return)
        self.assertIsNotNone(page.template_picker.currentData())
        self.win.navigate('dirs', 'templates')
        dirs = self.win.pages['dirs']
        tab = dirs.current()
        tab.name.setText('Черновик')
        self.win.resize(980, 620)
        self.app.processEvents()
        self.assertTrue(tab.save_btn.isVisible())
        self.assertTrue(tab.save_btn.visibleRegion().boundingRect().height() > 0)
        with patch.object(self.win, 'confirm', return_value=False):
            event = QCloseEvent()
            self.win.closeEvent(event)
            self.assertFalse(event.isAccepted())
            self.assertEqual(tab.name.text(), 'Черновик')
        with patch.object(self.win, 'confirm', return_value=True):
            dirs.tabs.setCurrentIndex(0)
        self.assertEqual(tab.name.text(), '')
        self.assertFalse(tab.has_unsaved_changes())
        self.assertEqual(dirs.tabs.currentIndex(), 0)

    def test_field_dialog_validation_and_escape(self):
        from sz.ui.templates import TemplateFieldsDialog
        dialog = TemplateFieldsDialog({'цель': 'Цель', 'дата': 'Дата'}, {'дата': '01.01.2026'})
        dialog.accept()
        self.assertNotEqual(dialog.result(), QDialog.Accepted)
        self.assertFalse(dialog.error.isHidden())
        dialog.inputs['цель'].setText('Тест')
        dialog.accept()
        self.assertEqual(dialog.result(), QDialog.Accepted)
        self.assertEqual(dialog.values()['дата'], '01.01.2026')


if __name__ == '__main__':
    unittest.main()
