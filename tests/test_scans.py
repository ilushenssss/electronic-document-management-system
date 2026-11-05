"""Загрузка PDF, номер СЭД, переименование и восстановление при ошибках."""
from pathlib import Path
import sqlite3
import re
import struct
import tempfile
import unittest
from unittest.mock import patch

from sz import actions, scans
from sz.store import Store, StoreError
from test_core import seed, memo_data

NUMBER = '1/441062/1100/СЗ'
NEXT = '1/441062/1101/СЗ'
PDF = b'%PDF-1.4\n% scanned document\n%%EOF\n'


DESCRIPTOR = 'application/x-qt-windows-mime;value="FileGroupDescriptorW"'
CONTENTS = 'application/x-qt-windows-mime;value="FileContents";index=0'


class FakeOutlookMime:
    def __init__(self, name='Скан.pdf', content=PDF):
        descriptor = bytearray(596)
        struct.pack_into('<I', descriptor, 0, 1)
        encoded = name.encode('utf-16-le')
        descriptor[76:76 + len(encoded)] = encoded
        self.formats = {DESCRIPTOR: bytes(descriptor), CONTENTS: content}
        self.reads = []

    def hasFormat(self, fmt):
        return fmt in self.formats

    def data(self, fmt):
        self.reads.append(fmt)
        return self.formats.get(fmt, b'')


class OutlookMimeTest(unittest.TestCase):
    def test_descriptor_does_not_load_attachment_until_drop(self):
        mime = FakeOutlookMime()
        self.assertEqual(scans.outlook_pdf_name(mime), 'Скан.pdf')
        self.assertNotIn(CONTENTS, mime.reads)
        self.assertEqual(scans.outlook_pdf_bytes(mime), PDF)
        self.assertIn(CONTENTS, mime.reads)

    def test_rejects_multiple_non_pdf_and_missing_or_invalid_contents(self):
        for descriptor in (b'', struct.pack('<I', 2) + bytes(592)):
            mime = FakeOutlookMime()
            mime.formats[DESCRIPTOR] = descriptor
            self.assertIsNone(scans.outlook_pdf_name(mime))
        self.assertIsNone(scans.outlook_pdf_name(FakeOutlookMime('Письмо.msg')))
        for contents in (b'', b'not a PDF'):
            with self.subTest(contents=contents), self.assertRaises(StoreError):
                scans.outlook_pdf_bytes(FakeOutlookMime(content=contents))


class ScanTest(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.TemporaryDirectory()
        self.store = Store(self.root.name)
        self.store.open()
        seed(self.store)
        self.mid, error = actions.create_memo(self.store, memo_data(), [1])
        self.assertIsNone(error)
        self.source = Path(self.root.name) / 'original.PDF'
        self.source.write_bytes(PDF)

    def tearDown(self):
        self.store.close()
        self.root.cleanup()

    def memo(self):
        return self.store.one('SELECT * FROM memos WHERE id = ?', (self.mid,))

    def upload(self, **kwargs):
        return scans.save(self.store, self.mid, NUMBER, str(self.source), **kwargs)

    def test_upload_name_folder_original_and_registry_search(self):
        old = self.memo()
        word = Path(self.store.document_path(old['file_path'])).read_bytes()
        result = self.upload(expected=('', ''))
        path = Path(self.store.document_path(result['scan_path']))
        self.assertEqual(path.name, '1_441062_1100_СЗ.pdf')
        self.assertEqual(path.parent, Path(self.store.abs_path(old['file_path'])).parent)
        self.assertEqual(path.read_bytes(), PDF)
        self.assertEqual(self.source.read_bytes(), PDF)
        self.assertEqual(self.memo()['sed_number'], NUMBER)
        self.assertEqual(actions.search_memos_page(self.store, q='441062/1100/сз')['matched'], 1)
        self.assertEqual(Path(self.store.document_path(old['file_path'])).read_bytes(), word)
        self.store.close()
        self.store.open()
        self.assertEqual(self.memo()['sed_number'], NUMBER)
        self.assertEqual(scans.document_path(self.store, self.memo()), str(path))

    def test_number_without_pdf_and_manual_file(self):
        result = scans.save(self.store, self.mid, NUMBER)
        self.assertEqual(result['scan_path'], '')
        target = Path(self.store.abs_path(scans.target_path(self.store, self.memo(), NUMBER)))
        target.write_bytes(PDF)
        self.assertEqual(scans.document_path(self.store, self.memo()), str(target))
        result = scans.save(self.store, self.mid, NEXT)
        new = Path(self.store.abs_path(result['scan_path']))
        self.assertEqual(new.name, '1_441062_1101_СЗ.pdf')
        self.assertEqual(new.read_bytes(), PDF)
        self.assertFalse(target.exists())

    def test_paper_scan_uses_internal_number_and_can_change_back(self):
        result = scans.save(self.store, self.mid, '', str(self.source), in_paper=True)
        path = Path(self.store.abs_path(result['scan_path']))
        self.assertEqual(path.name, 'НТЦ-0001-26.pdf')
        self.assertEqual(path.read_bytes(), PDF)
        self.assertEqual((self.memo()['sed_number'], self.memo()['in_paper']), ('', 1))
        self.assertEqual(scans.sed_label(self.memo()), 'В бумаге')
        self.assertEqual(actions.search_memos_page(self.store, q='В БУМАГЕ')['matched'], 1)
        normal = scans.save(self.store, self.mid, NUMBER, expected=('', result['scan_path'], 1))
        self.assertEqual(Path(normal['scan_path']).name, '1_441062_1100_СЗ.pdf')
        self.assertFalse(path.exists())
        self.assertEqual(self.memo()['in_paper'], 0)
        paper = scans.save(self.store, self.mid, NUMBER, in_paper=True)
        self.assertEqual(Path(paper['scan_path']).name, 'НТЦ-0001-26.pdf')
        self.assertEqual(self.memo()['sed_number'], '')

    def test_missing_sed_filter_excludes_paper_and_registered_memos(self):
        registered, _ = actions.create_memo(self.store, memo_data(), [1])
        paper, _ = actions.create_memo(self.store, memo_data(), [1])
        scans.save(self.store, registered, NUMBER)
        scans.save(self.store, paper, '', in_paper=True)
        self.store.write(lambda s: s.run('UPDATE memos SET sed_number = ? WHERE id = ?', ('  ', self.mid)))
        result = actions.search_memos_page(self.store, missing_sed=True, page=99, page_size=25,
                                           year=2026, executor_name='Сидоров Сидор Сидорович',
                                           addressee_department='Отдел кадров')
        self.assertEqual((result['total'], result['matched'], result['page']), (3, 1, 0))
        self.assertEqual([m['id'] for m in result['memos']], [self.mid])
        self.assertEqual([m['id'] for m in actions.search_memos(self.store, missing_sed=True)], [self.mid])
        self.assertEqual(actions.search_memos_page(self.store, missing_sed=True, q='В бумаге')['matched'], 0)

    def test_paper_marker_without_upload_discovers_manual_internal_pdf(self):
        result = scans.save(self.store, self.mid, '', in_paper=True)
        self.assertEqual(result['scan_path'], '')
        path = Path(self.store.abs_path(scans.target_path(self.store, self.memo(), '', True)))
        path.write_bytes(PDF)
        self.assertEqual(scans.document_path(self.store, self.memo()), str(path))
        with self.assertRaises(StoreError):
            scans.save(self.store, self.mid, NUMBER, expected=('', ''))

    def test_paper_flag_rolls_back_if_rename_is_blocked(self):
        old = self.upload()
        with patch('sz.scans.os.replace', side_effect=PermissionError('открыт')):
            with self.assertRaises(StoreError):
                scans.save(self.store, self.mid, '', in_paper=True)
        self.assertEqual((self.memo()['sed_number'], self.memo()['in_paper']), (NUMBER, 0))
        self.assertEqual(Path(self.store.abs_path(old['scan_path'])).read_bytes(), PDF)

    def test_manually_named_file_is_bound_without_upload(self):
        target = Path(self.store.abs_path(scans.target_path(self.store, self.memo(), NUMBER)))
        target.write_bytes(PDF)
        result = scans.save(self.store, self.mid, NUMBER)
        self.assertEqual(self.store.document_path(result['scan_path']), str(target))

    def test_format_and_invalid_pdf_leave_database_unchanged(self):
        for number in ('', '../1/2/СЗ', '1_441062_1100_СЗ', '1/2/3/4/СЗ', '1/2/3/SZ'):
            with self.subTest(number=number), self.assertRaises(StoreError):
                scans.save(self.store, self.mid, number)
        self.assertEqual(scans.clean_number(' 01/00441062/01100/сз '), '01/00441062/01100/СЗ')
        self.source.write_bytes(b'not PDF')
        with self.assertRaises(StoreError):
            self.upload()
        self.assertEqual(self.memo()['sed_number'], '')

    def test_replacement_requires_confirmation_and_keeps_previous_version(self):
        first = self.upload()
        target = Path(self.store.abs_path(first['scan_path']))
        self.source.write_bytes(PDF + b'new')
        with self.assertRaises(scans.ScanExists):
            self.upload()
        self.assertEqual(target.read_bytes(), PDF)
        self.upload(overwrite=True)
        backups = list((target.parent / 'Предыдущие версии').glob('*.pdf'))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), PDF)
        self.assertEqual(target.read_bytes(), self.source.read_bytes())

    def test_locked_pdf_does_not_change_number_or_source(self):
        first = self.upload()
        old = Path(self.store.abs_path(first['scan_path']))
        with patch('sz.scans.os.replace', side_effect=PermissionError('открыт')):
            with self.assertRaises(StoreError):
                scans.save(self.store, self.mid, NEXT)
        self.assertEqual(self.memo()['sed_number'], NUMBER)
        self.assertEqual(old.read_bytes(), PDF)
        self.assertFalse(Path(self.store.abs_path(scans.target_path(self.store, self.memo(), NEXT))).exists())
        self.assertEqual(list(old.parent.glob('*.tmp')), [])

    def test_database_error_restores_replaced_pdf(self):
        first = self.upload()
        target = Path(self.store.abs_path(first['scan_path']))
        self.source.write_bytes(PDF + b'new')
        real_db = self.store.db

        class FailCommit:
            def execute(self, sql, params=()):
                if sql == 'COMMIT':
                    raise sqlite3.OperationalError('test commit failure')
                return real_db.execute(sql, params)

        with patch.object(self.store, 'db', FailCommit()):
            with self.assertRaises(sqlite3.OperationalError):
                self.upload(overwrite=True)
        self.assertEqual(target.read_bytes(), PDF)
        self.assertEqual(self.memo()['sed_number'], NUMBER)
        self.assertEqual(list(target.parent.glob('*.tmp')), [])

    def test_conflicting_sed_change_and_stale_memo_edit_are_rejected(self):
        snapshot = actions.load_memo_for_edit(self.store, self.mid)
        self.upload(expected=('', ''))
        with self.assertRaises(StoreError):
            scans.save(self.store, self.mid, NEXT, expected=('', ''))
        with self.assertRaises(actions.MemoConflict):
            actions.update_memo(self.store, snapshot, snapshot['version'],
                                 dict(memo_data(), topic_id=-1, executor_id=-1), [-1])

    def test_missing_word_path_still_uses_memo_folder(self):
        self.store.write(lambda s: s.run('UPDATE memos SET file_path = NULL WHERE id = ?', (self.mid,)))
        result = self.upload()
        self.assertIn('2026/НТЦ-0001-26/1_441062_1100_СЗ.pdf', result['scan_path'])

    def test_old_database_is_migrated_without_losing_records(self):
        self.store.close()
        db = sqlite3.connect(Path(self.root.name) / 'data' / 'memos.db')
        # Старые рабочие базы не имеют новых колонок.
        schema = db.execute("SELECT sql FROM sqlite_master WHERE name = 'memos'").fetchone()[0]
        schema = re.sub(r",\s*(?:sed_number|scan_path) TEXT NOT NULL DEFAULT ''", '', schema)
        schema = re.sub(r",\s*in_paper INTEGER NOT NULL DEFAULT 0", '', schema)
        columns = [r[1] for r in db.execute('PRAGMA table_info(memos)') if r[1] not in ('sed_number', 'scan_path', 'in_paper')]
        db.execute('ALTER TABLE memos RENAME TO old_memos')
        db.execute(schema)
        names = ','.join(columns)
        db.execute('INSERT INTO memos (%s) SELECT %s FROM old_memos' % (names, names))
        db.execute('DROP TABLE old_memos')
        db.commit()
        db.close()
        self.store.open()
        self.assertEqual(self.memo()['reg_number'], 'НТЦ-0001-26')
        self.assertEqual((self.memo()['sed_number'], self.memo()['scan_path']), ('', ''))
        self.assertEqual(self.memo()['in_paper'], 0)
        self.upload()


if __name__ == '__main__':
    unittest.main()
