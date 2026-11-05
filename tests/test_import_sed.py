"""Перенос только номера СЭД/бумажного статуса из старого журнала."""
import csv
from pathlib import Path
import sqlite3
import tempfile
import unittest

from sz import actions
from sz.store import Store, StoreError
from test_core import seed, memo_data
from tools import import_sed_numbers as imp

try:
    import openpyxl
except ImportError:
    openpyxl = None


class NormalizeSedTest(unittest.TestCase):
    def test_supported_separators_and_only_whole_numbers(self):
        for sep in ('/', '\\', '_', '|', '–', '／', ':', '.', ' ', '?', '*'):
            self.assertEqual(imp.normalize_sed(sep.join(['01', '441062', '01100', 'сз']))[:2],
                             ('01/441062/01100/СЗ', 0))
        for value in ('в бумаге', ' В БУМАГЕ ', 'в отд.40', 'В отд. 40'):
            self.assertEqual(imp.normalize_sed(value)[:2], ('', 1))
        for value in (None, '', 12345, '1С-1', 'см. письмо', 'x1/441062/1100/СЗ',
                      '1/441062/1100/СЗ и ещё текст', '1/441062/1100/SZ'):
            self.assertEqual(imp.normalize_sed(value)[:2], ('', 0))

    def test_missing_database_is_not_created(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / 'missing.db'
            with self.assertRaises(StoreError):
                imp.connect(path)
            self.assertFalse(path.exists())


@unittest.skipIf(openpyxl is None, 'openpyxl не установлен')
class ImportSedTest(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.TemporaryDirectory()
        self.addCleanup(self.root.cleanup)
        self.store = Store(self.root.name)
        self.store.open()
        self.addCleanup(self.store.close)
        seed(self.store)
        for _ in range(5):
            actions.create_memo(self.store, memo_data(body='Не менять текст'), [1])
        self.store.write(lambda s: s.run("UPDATE memos SET sed_number = '1/111111/1111/СЗ'"))
        self.store.write(lambda s: s.run("UPDATE memos SET sed_number = '', in_paper = 1 WHERE number = 4"))
        self.db_path = Path(self.store.db_path)
        self.xlsx = Path(self.root.name) / 'журнал.xlsx'

    def journal(self, rows):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(['Старый журнал'])
        sheet.append(['Номер СЗ', 'Номер СЗ 1C'])  # латинская C в заголовке
        for row in rows:
            sheet.append(row)
        workbook.save(self.xlsx)
        workbook.close()

    def plan(self):
        _, entries = imp.read_excel(self.xlsx)
        return imp.make_plan(self.store.db, entries)

    def test_import_updates_only_two_fields_and_makes_verified_backup(self):
        self.journal([['нтц–1–26', r'1\441062\1100\СЗ'], ['НТЦ-0002-26', 'в отд.40'],
                      ['НТЦ-3-26', 'какой-то текст'], ['НТЦ-4-26', None],
                      ['НТЦ-9999-26', '1/2/3/СЗ']])
        before = self.store.all('SELECT * FROM memos ORDER BY number')
        word_files = {path: path.read_bytes() for path in Path(self.root.name).rglob('*.docx')}
        plan, report = self.plan()
        self.assertEqual(len(plan), 4)
        self.assertTrue(any(item['status'] == 'Пропуск: СЗ не найдена' for item in report))
        backup = imp.apply_plan(self.db_path, plan)
        after = self.store.all('SELECT * FROM memos ORDER BY number')
        self.assertEqual([(m['sed_number'], m['in_paper']) for m in after],
                         [('1/441062/1100/СЗ', 0), ('', 1), ('', 0), ('', 0), ('1/111111/1111/СЗ', 0)])
        for old, new in zip(before, after):
            self.assertEqual({k: v for k, v in old.items() if k not in ('sed_number', 'in_paper')},
                             {k: v for k, v in new.items() if k not in ('sed_number', 'in_paper')})
        self.assertEqual(word_files, {path: path.read_bytes() for path in word_files})
        db = sqlite3.connect(backup)
        try:
            self.assertEqual(db.execute('PRAGMA integrity_check').fetchone()[0], 'ok')
            self.assertEqual(db.execute('SELECT sed_number FROM memos WHERE number=1').fetchone()[0],
                             '1/111111/1111/СЗ')
        finally:
            db.close()
        self.assertEqual(self.plan()[0], [])  # повторный импорт ничего не меняет

    def test_duplicates_prefix_and_uncached_formula_are_not_silently_imported(self):
        self.journal([['НТЦ-1-26', '1/2/3/СЗ'], ['НТЦ-0001-26', 'в бумаге'],
                      ['НТЦ-2-26', '1_2_3_СЗ'], ['НТЦ-2-26', '1/2/3/СЗ'],
                      ['ДРУГОЙ-3-26', '1/2/3/СЗ'], ['НТЦ-4-26', '=1+1']])
        plan, report = self.plan()
        self.assertEqual(len(plan), 1)
        self.assertEqual(plan[0]['id'], 2)
        self.assertEqual(sum('противоречивые' in r['status'] for r in report), 2)
        self.assertTrue(any('Формула без' in r['status'] for r in report))
        self.assertTrue(any('не найдена' in r['status'] for r in report))

    def test_changed_database_aborts_whole_import(self):
        self.journal([['НТЦ-1-26', None], ['НТЦ-2-26', None]])
        plan, _ = self.plan()
        self.store.write(lambda s: s.run("UPDATE memos SET sed_number = '1/2/3/СЗ' WHERE number=2"))
        with self.assertRaises(StoreError):
            imp.apply_plan(self.db_path, plan)
        self.assertEqual(self.store.one('SELECT sed_number FROM memos WHERE number=1')['sed_number'],
                         '1/111111/1111/СЗ')

    def test_sql_error_rolls_back_all_updates(self):
        self.journal([['НТЦ-1-26', None], ['НТЦ-2-26', None]])
        plan, _ = self.plan()
        self.store.db.execute("CREATE TRIGGER reject_second BEFORE UPDATE OF sed_number ON memos "
                              "WHEN NEW.number=2 BEGIN SELECT RAISE(ABORT, 'test'); END")
        with self.assertRaises(sqlite3.IntegrityError):
            imp.apply_plan(self.db_path, plan)
        self.assertEqual(self.store.one('SELECT sed_number FROM memos WHERE number=1')['sed_number'],
                         '1/111111/1111/СЗ')

    def test_cli_dry_run_then_confirmed_import(self):
        self.journal([['НТЦ-1-26', '1\\441062\\1100\\СЗ']])
        report = Path(self.root.name) / 'report.csv'
        self.assertEqual(imp.main([str(self.xlsx), self.root.name, '--dry-run', '--report', str(report)]), 0)
        self.assertEqual(self.store.one('SELECT sed_number FROM memos WHERE number=1')['sed_number'],
                         '1/111111/1111/СЗ')
        self.assertEqual(list(self.db_path.parent.glob('memos-before-sed-import-*.db')), [])
        with report.open(encoding='utf-8-sig', newline='') as file:
            rows = list(csv.reader(file, delimiter=';'))
        self.assertEqual(rows[1][4], '1/441062/1100/СЗ')
        self.assertEqual(imp.main([str(self.xlsx), str(self.db_path), '--yes', '--report', str(report)]), 0)
        self.assertEqual(self.store.one('SELECT sed_number FROM memos WHERE number=1')['sed_number'],
                         '1/441062/1100/СЗ')
        self.assertEqual(len(list(self.db_path.parent.glob('memos-before-sed-import-*.db'))), 1)


if __name__ == '__main__':
    unittest.main()
