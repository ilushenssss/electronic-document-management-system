"""Тесты формирования документа, хранилища и одновременной записи.

Запуск: python -m unittest discover -s tests
"""
import io
import json
import multiprocessing
import os
import shutil
import sqlite3
import sys
import tempfile
import time
import unittest
import zipfile
from xml.dom import minidom

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sz import actions, docx  # noqa: E402
from sz.store import FileLock, Store, StoreError  # noqa: E402


def seed(store):
    actions.save_settings(store, 'АО «НТЦ»\nОтдел', 'Москва', 'НТЦ', 'Служебные записки')
    actions.save_directory_row(store, 'employees', {'full_name': 'Сидоров Сидор Сидорович',
                                                    'position': 'Главный инженер', 'phone': '12-34'})
    actions.save_directory_row(store, 'addressees', {
        'department': 'Отдел кадров', 'head_position': 'Начальнику отдела кадров',
        'head_name': 'Петров Пётр Петрович', 'head_name_dative': 'Петрову П.П.'})
    actions.save_directory_row(store, 'topics', {'name': 'Закупка'})
    actions.set_signer(store, 1, True)


def memo_data(date='2026-03-05', body='Текст'):
    return {'topic_id': 1, 'executor_id': 1, 'memo_date': date, 'body': body, 'attachment': ''}


class DocxTest(unittest.TestCase):
    def test_formatting(self):
        self.assertEqual(docx.format_reg_number('НТЦ', 7, 2026), 'НТЦ-0007-26')
        self.assertEqual(docx.format_date('2026-01-14'), '14.01.2026')
        self.assertEqual(docx.format_long_date('2026-01-14'), '"14" января 2026 г.')
        self.assertEqual(docx.short_name('Иванов Иван Иванович'), 'Иванов И.И.')
        self.assertEqual(docx.short_name('Иванов И. И.'), 'Иванов И.И.')
        self.assertEqual(docx.greeting([{'head_name': 'Петров Пётр Петрович'}]), 'Уважаемый Пётр Петрович!')
        self.assertEqual(docx.greeting([{'head_name': 'Смирнова Ольга Ивановна'}]), 'Уважаемая Ольга Ивановна!')
        self.assertEqual(docx.greeting([{}, {}]), 'Уважаемые Коллеги!')
        self.assertEqual(docx.safe_name('a/b:c?'), 'a b c')

    def test_build_docx(self):
        memo = {'reg_number': 'НТЦ-0001-26', 'memo_date': '2026-01-14', 'topic_name': 'Тема <&>',
                'body': 'Абзац 1\nАбзац 2', 'attachment': 'смета', 'org_name': None, 'city': 'Москва',
                'signer_name': 'Сидоров Сидор Сидорович', 'signer_position': 'Начальник\nцентра',
                'executor_name': 'Иванов Иван Иванович', 'executor_phone': '12-34'}
        data = docx.build_docx(memo, [{'head_position': 'Начальнику', 'head_name': 'Петров Пётр Петрович',
                                       'head_name_dative': 'Петрову П.П.'}], org_name='АО «НТЦ»')
        z = zipfile.ZipFile(io.BytesIO(data))
        self.assertEqual(z.testzip(), None)
        xml = z.read('word/document.xml').decode('utf-8')
        minidom.parseString(xml)  # корректный XML
        for text in ('АО «НТЦ»', 'СЛУЖЕБНАЯ ЗАПИСКА', '№НТЦ-0001-26', 'Тема &lt;&amp;&gt;', 'Петрову П.П.',
                     'Уважаемый Пётр Петрович!', 'Приложение: ', 'Сидоров С.С.'):
            self.assertIn(text, xml)
        self.assertNotIn('Исп. Иванов И.И.', xml)
        self.assertIn('Исп. Иванов И.И.', z.read('word/footer1.xml').decode('utf-8'))

    def test_preview_placeholders(self):
        blocks = docx.memo_blocks({'reg_number': 'X'}, [], preview=True)
        self.assertTrue(any(isinstance(b.get('text'), dict) for b in blocks))


class StoreTest(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp()
        self.store = Store(self.root)
        self.store.open()

    def tearDown(self):
        self.store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def test_create_memo_and_document(self):
        seed(self.store)
        memo_id, err = actions.create_memo(self.store, memo_data(), [1])
        self.assertIsNone(err)
        memo = self.store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        self.assertEqual(memo['reg_number'], 'НТЦ-0001-26')
        self.assertEqual(memo['file_path'], 'Служебные записки/2026/НТЦ-0001-26/НТЦ-0001-26.docx')
        self.assertTrue(self.store.document_path(memo['file_path']))
        # Нумерация по годам
        actions.create_memo(self.store, memo_data(), [1])
        second, _ = actions.create_memo(self.store, memo_data('2027-01-01'), [1])
        regs = [m['reg_number'] for m in self.store.all('SELECT * FROM memos ORDER BY id')]
        self.assertEqual(regs, ['НТЦ-0001-26', 'НТЦ-0002-26', 'НТЦ-0001-27'])
        # Повторное формирование - по прежнему пути
        before = self.store.one('SELECT file_path FROM memos WHERE id = ?', (second,))['file_path']
        self.assertEqual(actions.regenerate_document(self.store, second), before)
        actions.delete_memo(self.store, second)
        self.assertIsNone(self.store.one('SELECT * FROM memos WHERE id = ?', (second,)))

    def test_validation(self):
        seed(self.store)
        with self.assertRaises(StoreError):
            actions.create_memo(self.store, memo_data(), [])
        with self.assertRaises(StoreError):
            actions.create_memo(self.store, memo_data(body='  '), [1])
        with self.assertRaises(StoreError):
            actions.save_directory_row(self.store, 'topics', {'name': 'Закупка'})  # дубликат
        actions.save_directory_row(self.store, 'employees', {'full_name': 'Без Должности Иванович'})
        with self.assertRaises(StoreError):
            actions.set_signer(self.store, 2, True)  # подписант уже назначен
        actions.set_signer(self.store, 1, False)
        with self.assertRaises(StoreError):
            actions.set_signer(self.store, 2, True)  # нет должности
        with self.assertRaises(StoreError):
            actions.save_settings(self.store, '', '', 'A/B', 'Docs')
        with self.assertRaises(StoreError):
            actions.save_settings(self.store, '', '', 'НТЦ', '../outside')

    def test_changed_by_other_user(self):
        other = Store(self.root)
        other.open()
        self.assertFalse(self.store.changed())
        actions.save_directory_row(other, 'topics', {'name': 'Новая'})
        self.assertTrue(self.store.changed())
        self.assertFalse(self.store.changed())
        other.close()

    def test_search_and_stats(self):
        seed(self.store)
        actions.create_memo(self.store, memo_data(body='Про ОБОРУДОВАНИЕ'), [1])
        actions.create_memo(self.store, memo_data('2026-05-01', 'Другое'), [1])
        self.assertEqual(len(actions.search_memos(self.store, 'оборудование')), 1)
        self.assertEqual(len(actions.search_memos(self.store, year=2026)), 2)
        self.assertEqual(actions.month_counts(self.store, 2026)[2], 1)
        self.assertEqual(actions.group_counts(self.store, 'addressee', 2026), [{'label': 'Отдел кадров', 'n': 2}])


class MigrationTest(unittest.TestCase):
    def test_old_database(self):
        """База первой браузерной версии: без новых колонок и таблицы адресатов записки."""
        root = tempfile.mkdtemp()
        os.makedirs(os.path.join(root, 'data'))
        db = sqlite3.connect(os.path.join(root, 'data', 'memos.db'))
        db.executescript("""
            CREATE TABLE employees (id INTEGER PRIMARY KEY AUTOINCREMENT, full_name TEXT NOT NULL,
              phone TEXT NOT NULL DEFAULT '');
            CREATE TABLE addressees (id INTEGER PRIMARY KEY AUTOINCREMENT, department TEXT NOT NULL,
              head_position TEXT NOT NULL, head_name TEXT NOT NULL);
            CREATE TABLE topics (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE);
            CREATE TABLE memos (id INTEGER PRIMARY KEY AUTOINCREMENT, number INTEGER NOT NULL,
              year INTEGER NOT NULL, memo_date TEXT NOT NULL, topic_id INTEGER, topic_name TEXT NOT NULL,
              addressee_id INTEGER, addressee_department TEXT NOT NULL, addressee_position TEXT NOT NULL,
              addressee_name TEXT NOT NULL, body TEXT NOT NULL, executor_id INTEGER,
              executor_name TEXT NOT NULL, executor_phone TEXT NOT NULL DEFAULT '', file_path TEXT,
              created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')), UNIQUE (year, number));
            CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            INSERT INTO settings VALUES ('number_prefix', 'ОТД');
            INSERT INTO memos (number, year, memo_date, topic_name, addressee_department, addressee_position,
              addressee_name, body, executor_name) VALUES (3, 2025, '2025-02-01', 'Тема', 'Отдел', 'Начальнику',
              'Петров Пётр Петрович', 'Текст', 'Иванов Иван Иванович');
        """)
        db.commit()
        db.close()
        store = Store(root)
        store.open()
        memo = store.one('SELECT * FROM memos')
        self.assertEqual(memo['reg_number'], 'ОТД-0003-25')
        self.assertIsNone(memo['org_name'])
        self.assertEqual(store.memo_addressees(memo['id'])[0]['head_name'], 'Петров Пётр Петрович')
        self.assertTrue(actions.regenerate_document(store, memo['id']))
        store.close()
        shutil.rmtree(root, ignore_errors=True)


def _worker(root, count, out):
    store = Store(root)
    store.open()
    ids = []
    for _ in range(count):
        memo_id, err = actions.create_memo(store, memo_data(), [1])
        ids.append(memo_id)
    store.close()
    out.put(ids)


class ConcurrencyTest(unittest.TestCase):
    def test_parallel_writers_get_unique_numbers(self):
        root = tempfile.mkdtemp()
        store = Store(root)
        store.open()
        seed(store)
        store.close()
        out = multiprocessing.Queue()
        procs = [multiprocessing.Process(target=_worker, args=(root, 8, out)) for _ in range(4)]
        for p in procs:
            p.start()
        ids = [i for _ in procs for i in out.get(timeout=120)]
        for p in procs:
            p.join()
            self.assertEqual(p.exitcode, 0)
        store = Store(root)
        store.open()
        numbers = [r['number'] for r in store.all('SELECT number FROM memos ORDER BY number')]
        self.assertEqual(numbers, list(range(1, 33)))
        self.assertEqual(len(set(ids)), 32)
        files = os.listdir(os.path.join(root, 'Служебные записки', '2026'))
        self.assertEqual(len(files), 32)
        self.assertFalse(os.path.exists(store.lock_path))
        store.close()
        shutil.rmtree(root, ignore_errors=True)

    def test_stale_lock_is_broken(self):
        root = tempfile.mkdtemp()
        path = os.path.join(root, 'memos.lock')
        with open(path, 'w') as f:
            json.dump({'token': 'old', 'time': (time.time() - 120) * 1000}, f)
        with FileLock(path, timeout=2):
            with open(path) as f:
                self.assertNotEqual(json.load(f)['token'], 'old')
        self.assertFalse(os.path.exists(path))
        shutil.rmtree(root, ignore_errors=True)

    def test_busy_lock_times_out(self):
        root = tempfile.mkdtemp()
        path = os.path.join(root, 'memos.lock')
        with open(path, 'w') as f:
            json.dump({'token': 'other', 'time': time.time() * 1000, 'user': 'ivanov'}, f)
        with self.assertRaises(StoreError) as ctx:
            FileLock(path, timeout=0.5).acquire()
        self.assertIn('ivanov', str(ctx.exception))
        shutil.rmtree(root, ignore_errors=True)


if __name__ == '__main__':
    unittest.main()

