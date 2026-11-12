"""Сохранение отметки конфигурационного контроля и перенос в Word."""
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

from sz import actions, docx
from sz.store import Store
from test_core import seed, memo_data

NOTE = 'Для конфигурационного\nконтроля'
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


class ConfigurationControlTest(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.TemporaryDirectory()
        self.store = Store(self.root.name)
        self.store.open()
        seed(self.store)

    def tearDown(self):
        self.store.close()
        self.root.cleanup()

    def test_note_follows_topic_without_blank_and_has_word_line_break(self):
        for preview in (False, True):
            memo = {'topic_name': 'Закупка', 'configuration_control': 1}
            blocks = docx.memo_blocks(memo, [], preview=preview)
            topic_index = next(i for i, b in enumerate(blocks) if b.get('text') == 'Закупка')
            self.assertEqual(blocks[topic_index + 1].get('text'), NOTE)
        document = ET.fromstring(docx.document_xml(memo, []))
        paragraphs = document.findall('w:body/w:p', NS)
        index = next(i for i, p in enumerate(paragraphs) if ''.join(p.itertext()) == 'Закупка')
        self.assertEqual([t.text for t in paragraphs[index + 1].findall('.//w:t', NS)],
                         ['Для конфигурационного', 'контроля'])
        self.assertEqual(len(paragraphs[index + 1].findall('.//w:br', NS)), 1)
        self.assertFalse(any(b.get('text') == NOTE for b in docx.memo_blocks({'topic_name': 'Закупка'}, [])))

    def test_creation_persistence_editing_and_updated_word(self):
        data = dict(memo_data(), configuration_control=True)
        memo_id, error = actions.create_memo(self.store, data, [1])
        self.assertIsNone(error)
        self.store.close()
        self.store.open()
        snapshot = actions.load_memo_for_edit(self.store, memo_id)
        self.assertEqual(snapshot['memo']['configuration_control'], 1)
        with zipfile.ZipFile(self.store.document_path(snapshot['memo']['file_path'])) as word:
            self.assertIn('Для конфигурационного', word.read('word/document.xml').decode('utf-8'))
        edited = dict(memo_data(), topic_id=-1, executor_id=-1, configuration_control=False)
        result = actions.update_memo(self.store, snapshot, snapshot['version'], edited, [-1])
        self.assertIsNone(result['document_error'])
        memo = self.store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        self.assertEqual(memo['configuration_control'], 0)
        self.assertEqual(memo['reg_number'], snapshot['memo']['reg_number'])
        with zipfile.ZipFile(self.store.document_path(memo['file_path'])) as word:
            self.assertNotIn('Для конфигурационного', word.read('word/document.xml').decode('utf-8'))

    def test_legacy_call_without_field_preserves_saved_flag(self):
        memo_id, error = actions.create_memo(self.store, dict(memo_data(), configuration_control=True), [1])
        self.assertIsNone(error)
        snapshot = actions.load_memo_for_edit(self.store, memo_id)
        result = actions.update_memo(self.store, snapshot, snapshot['version'],
                                     dict(memo_data(), topic_id=-1, executor_id=-1), [-1])
        self.assertIsNone(result['document_error'])
        self.assertEqual(self.store.one('SELECT configuration_control FROM memos WHERE id = ?',
                                        (memo_id,))['configuration_control'], 1)

    def test_unchecked_default_does_not_add_note(self):
        memo_id, error = actions.create_memo(self.store, memo_data(), [1])
        self.assertIsNone(error)
        memo = self.store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        self.assertEqual(memo['configuration_control'], 0)
        with zipfile.ZipFile(self.store.document_path(memo['file_path'])) as word:
            self.assertNotIn('Для конфигурационного', word.read('word/document.xml').decode('utf-8'))


if __name__ == '__main__':
    unittest.main()
