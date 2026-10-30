"""Редактирование СЗ, конфликты, резервные версии и недоступный Word."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from sz import actions
from sz.store import Store, StoreError
from test_core import seed, memo_data


class EditTest(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.TemporaryDirectory()
        self.store = Store(self.root.name)
        self.store.open()
        seed(self.store)
        self.mid, error = actions.create_memo(self.store, memo_data(body='Исходный текст'), [1])
        self.assertIsNone(error)
        self.snapshot = actions.load_memo_for_edit(self.store, self.mid)
        self.path = Path(self.store.document_path(self.snapshot['memo']['file_path']))
        self.original_file = self.path.read_bytes()

    def tearDown(self):
        self.store.close()
        self.root.cleanup()

    def data(self, **changes):
        data = {'memo_date': '2026-04-01', 'body': 'Новый текст', 'topic_id': -1,
                'executor_id': -1, 'attachment': 'Новая смета'}
        data.update(changes)
        return data

    def save(self, version=None, **changes):
        return actions.update_memo(self.store, self.snapshot, version or self.snapshot['version'],
                                   self.data(**changes), [-1])

    def test_update_preserves_identity_and_backs_up_word(self):
        result = self.save(body_rich=[{'align': 'right', 'indent': False,
                                    'runs': [{'text': 'Оформленный текст', 'bold': True}]}])
        self.assertIsNone(result['document_error'])
        current = actions.load_memo_for_edit(self.store, self.mid)
        self.assertEqual(self.store.one('SELECT COUNT(*) AS n FROM memos')['n'], 1)
        for key in ('id', 'reg_number', 'number', 'year', 'file_path', 'created_at',
                    'signer_name', 'signer_position', 'org_name', 'city'):
            self.assertEqual(current['memo'][key], self.snapshot['memo'][key])
        self.assertEqual(current['memo']['body'], 'Оформленный текст')
        self.assertTrue(json.loads(current['memo']['body_rich'])[0]['runs'][0]['bold'])
        self.assertNotEqual(self.path.read_bytes(), self.original_file)
        backups = list((self.path.parent / 'Предыдущие версии').glob('*.docx'))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), self.original_file)
        self.assertEqual(result['version'], current['version'])

    def test_owner_file_warns_and_retry_keeps_input_and_old_document(self):
        owner = self.path.with_name('~$' + self.path.name[2:])
        owner.touch()
        result = self.save()
        self.assertIn('Закройте документ', result['document_error'])
        self.assertEqual(self.path.read_bytes(), self.original_file)
        self.assertEqual(self.store.one('SELECT body FROM memos WHERE id = ?', (self.mid,))['body'], 'Новый текст')
        owner.unlink()
        retried = self.save(version=result['version'])
        self.assertIsNone(retried['document_error'])
        self.assertNotEqual(self.path.read_bytes(), self.original_file)

    def test_replace_failure_does_not_truncate_existing_file(self):
        with patch('sz.store.os.replace', side_effect=PermissionError('занят')):
            result = self.save()
        self.assertIsNotNone(result['document_error'])
        self.assertEqual(self.path.read_bytes(), self.original_file)
        self.assertEqual(list(self.path.parent.glob('*.tmp')), [])

    def test_other_editor_conflicts_before_database_or_word_changes(self):
        other = Store(self.root.name)
        other.open()
        try:
            snapshot = actions.load_memo_for_edit(other, self.mid)
            actions.update_memo(other, snapshot, snapshot['version'], self.data(body='Правка коллеги'), [-1])
        finally:
            other.close()
        file_after_other = self.path.read_bytes()
        with self.assertRaises(actions.MemoConflict):
            self.save()
        self.assertEqual(self.path.read_bytes(), file_after_other)
        self.assertEqual(self.store.one('SELECT body FROM memos WHERE id = ?', (self.mid,))['body'], 'Правка коллеги')

    def test_deleted_directories_do_not_change_saved_values(self):
        actions.delete_directory_row(self.store, 'topics', 1)
        actions.delete_directory_row(self.store, 'addressees', 1)
        actions.delete_directory_row(self.store, 'employees', 1)
        # Удаление сотрудников снимает ссылки, поэтому форму открываем заново,
        # как это сделает пользователь после изменения справочников.
        self.snapshot = actions.load_memo_for_edit(self.store, self.mid)
        result = self.save()
        self.assertIsNone(result['document_error'])
        current = actions.load_memo_for_edit(self.store, self.mid)
        for field in ('executor_name', 'executor_phone', 'topic_name', 'signer_name', 'signer_position'):
            self.assertEqual(current['memo'][field], self.snapshot['memo'][field])
        self.assertEqual(current['addressees'][0]['head_position'], 'Начальнику отдела кадров')

    def test_year_validation_and_deleted_memo(self):
        for changes in ({'memo_date': '2027-01-01'}, {'memo_date': '2026-02-30'}, {'body': ''}):
            with self.assertRaises(StoreError):
                self.save(**changes)
        self.assertEqual(self.path.read_bytes(), self.original_file)
        actions.delete_memo(self.store, self.mid)
        with self.assertRaises(actions.MemoConflict):
            self.save()

    def test_can_choose_current_directories_and_preserve_recipient_order(self):
        actions.save_directory_row(self.store, 'employees', {'full_name': 'Иванов Иван Иванович', 'phone': '56-78'})
        actions.save_directory_row(self.store, 'topics', {'name': 'Ремонт'})
        actions.save_directory_row(self.store, 'addressees', {'department': 'Бухгалтерия',
            'head_position': 'Главному бухгалтеру', 'head_name': 'Смирнова Анна Ивановна',
            'head_name_dative': 'Смирновой А.И.'})
        result = actions.update_memo(self.store, self.snapshot, self.snapshot['version'],
                                     self.data(topic_id=2, executor_id=2), [2, -1])
        self.assertIsNone(result['document_error'])
        current = actions.load_memo_for_edit(self.store, self.mid)
        self.assertEqual(current['memo']['topic_name'], 'Ремонт')
        self.assertEqual(current['memo']['executor_phone'], '56-78')
        self.assertEqual([a['department'] for a in current['addressees']], ['Бухгалтерия', 'Отдел кадров'])


if __name__ == '__main__':
    unittest.main()
