"""Хранилище: база SQLite data/memos.db и документы Word в общей папке.

Программа работает с файлом базы напрямую (модуль sqlite3). Запись идёт так:
создать файл-блокировку data/memos.lock -> транзакция в базе -> удалить
блокировку. Файл блокировки создаётся атомарно (O_CREAT | O_EXCL), поэтому
на сетевой папке два компьютера не могут захватить его одновременно, а номера
СЗ не повторяются. Формат блокировки тот же, что у браузерной версии.
"""
import getpass
import json
import os
import random
import socket
import sqlite3
import time
import uuid

from . import docx

DATA_DIR = 'data'
DB_FILE = 'memos.db'
LOCK_FILE = 'memos.lock'
LOCK_STALE_S = 60  # блокировка старше минуты считается брошенной
LOCK_TIMEOUT_S = 20
DEFAULT_OUTPUT_DIR = 'Служебные записки'
DEFAULT_NUMBER_PREFIX = 'НТЦ'

# Справочные значения копируются в записку на момент создания,
# чтобы правка справочников не меняла уже выпущенные документы.
SCHEMA = """
CREATE TABLE IF NOT EXISTS employees (
  id        INTEGER PRIMARY KEY AUTOINCREMENT,
  full_name TEXT NOT NULL,
  phone     TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS addressees (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  department    TEXT NOT NULL,
  head_position TEXT NOT NULL,
  head_name     TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS topics (
  id   INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE
);
CREATE TABLE IF NOT EXISTS memos (
  id                   INTEGER PRIMARY KEY AUTOINCREMENT,
  number               INTEGER NOT NULL,
  year                 INTEGER NOT NULL,
  reg_number           TEXT,
  memo_date            TEXT NOT NULL,
  topic_id             INTEGER,
  topic_name           TEXT NOT NULL,
  addressee_id         INTEGER,
  addressee_department TEXT NOT NULL,
  addressee_position   TEXT NOT NULL,
  addressee_name       TEXT NOT NULL,
  body                 TEXT NOT NULL,
  executor_id          INTEGER,
  executor_name        TEXT NOT NULL,
  executor_phone       TEXT NOT NULL DEFAULT '',
  file_path            TEXT,
  created_at           TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
  UNIQUE (year, number)
);
CREATE TABLE IF NOT EXISTS settings (
  key   TEXT PRIMARY KEY,
  value TEXT NOT NULL
);
-- Коды сохранений браузерной версии (она по ним проверяет, что запись не затёрли)
CREATE TABLE IF NOT EXISTS write_log (
  id TEXT PRIMARY KEY,
  at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);
-- Адресаты записки (их может быть несколько), в порядке выбора
CREATE TABLE IF NOT EXISTS memo_addressees (
  memo_id       INTEGER NOT NULL,
  ord           INTEGER NOT NULL,
  addressee_id  INTEGER,
  department    TEXT NOT NULL,
  head_position TEXT NOT NULL,
  head_name     TEXT NOT NULL,
  PRIMARY KEY (memo_id, ord)
);
"""

# Колонки, добавленные после первой версии: (таблица, колонка, определение)
ADDED_COLUMNS = [
    ('memos', 'reg_number', 'TEXT'),
    ('memos', 'attachment', "TEXT NOT NULL DEFAULT ''"),
    ('memos', 'signer_id', 'INTEGER'),
    ('memos', 'signer_name', "TEXT NOT NULL DEFAULT ''"),
    ('memos', 'signer_position', "TEXT NOT NULL DEFAULT ''"),
    ('memos', 'city', "TEXT NOT NULL DEFAULT ''"),
    ('memos', 'org_name', 'TEXT'),  # NULL у записок прошлых версий: шапка берётся из настроек
    ('employees', 'position', "TEXT NOT NULL DEFAULT ''"),
    ('addressees', 'head_name_dative', "TEXT NOT NULL DEFAULT ''"),
    ('employees', 'is_signer', 'INTEGER NOT NULL DEFAULT 0'),  # подписант - не больше одного
    ('memo_addressees', 'head_name_dative', "TEXT NOT NULL DEFAULT ''"),
]


class StoreError(Exception):
    """Ошибка, текст которой показывается пользователю."""


def split_path(path):
    return [p for p in str(path).replace('\\', '/').split('/') if p and p != '.']


class FileLock:
    """Блокировка на время записи: файл, который создаётся только если его нет."""

    def __init__(self, path, stale=LOCK_STALE_S, timeout=LOCK_TIMEOUT_S):
        self.path = path
        self.stale = stale
        self.timeout = timeout
        self.token = None

    def _read(self):
        try:
            with open(self.path, encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return None
        except (OSError, ValueError):
            # Файл только что создан и ещё пуст или занят: смотрим на время изменения
            try:
                return {'time': os.path.getmtime(self.path) * 1000}
            except OSError:
                return None

    def _is_stale(self):
        current = self._read()
        return current is not None and time.time() * 1000 - (current.get('time') or 0) > self.stale * 1000

    def _break(self):
        # Переименование удаётся только одному из тех, кто заметил брошенную блокировку
        aside = '%s.%s.stale' % (self.path, uuid.uuid4().hex)
        try:
            os.rename(self.path, aside)
            os.remove(aside)
        except OSError:
            pass

    def acquire(self):
        token = uuid.uuid4().hex
        deadline = time.monotonic() + self.timeout
        denied = False
        while True:
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            except FileExistsError:
                denied = False
                if self._is_stale():
                    self._break()
                    continue
            except PermissionError:
                # В Windows так бывает и пока чужая блокировка удаляется
                denied = True
            else:
                with os.fdopen(fd, 'w', encoding='utf-8') as f:
                    json.dump({'token': token, 'time': int(time.time() * 1000),
                               'host': socket.gethostname(), 'user': _user()}, f)
                self.token = token
                return
            if time.monotonic() > deadline:
                if denied and not os.path.exists(self.path):
                    raise StoreError('Нет прав на запись в папку %s. Проверьте права доступа.'
                                     % os.path.dirname(self.path))
                raise StoreError('База занята другим пользователем%s. Повторите через минуту.'
                                 % _holder(self._read()))
            time.sleep(0.1 + random.random() * 0.2)

    def release(self):
        current = self._read()
        if current and current.get('token') == self.token:
            try:
                os.remove(self.path)
            except OSError:
                pass  # блокировка устареет сама
        self.token = None

    def __enter__(self):
        self.acquire()
        return self

    def __exit__(self, *exc):
        self.release()


def _user():
    try:
        return getpass.getuser()
    except Exception:
        return ''


def _holder(lock):
    if not lock or not (lock.get('user') or lock.get('host')):
        return ''
    return ' (%s)' % ', '.join(x for x in (lock.get('user'), lock.get('host')) if x)


class Store:
    def __init__(self, root):
        self.root = os.path.abspath(root)
        self.data_dir = os.path.join(self.root, DATA_DIR)
        self.db_path = os.path.join(self.data_dir, DB_FILE)
        self.lock_path = os.path.join(self.data_dir, LOCK_FILE)
        self.db = None
        self._version = None

    # ---------- База ----------

    def open(self):
        if not os.path.isdir(self.root):
            raise StoreError('Папка %s недоступна.' % self.root)
        try:
            os.makedirs(self.data_dir, exist_ok=True)
        except OSError as e:
            raise StoreError('Не удалось создать папку %s: %s' % (self.data_dir, e))
        # Создание базы и добавление колонок для базы прошлой версии - под блокировкой
        with self._lock():
            self.db = sqlite3.connect(self.db_path, timeout=LOCK_TIMEOUT_S, isolation_level=None)
            self.db.row_factory = sqlite3.Row
            # WAL на сетевой папке не работает
            self.db.execute('PRAGMA journal_mode = DELETE')
            self._prepare()
        self._version = self._data_version()

    def close(self):
        if self.db:
            self.db.close()
            self.db = None

    def _lock(self):
        return FileLock(self.lock_path)

    def _columns(self, table):
        return [r['name'] for r in self.db.execute('PRAGMA table_info(%s)' % table)]

    def _prepare(self):
        """Создаёт недостающие таблицы и колонки в базе, созданной прошлой версией."""
        self.db.executescript(SCHEMA)
        self.db.execute('BEGIN IMMEDIATE')
        try:
            for table, column, definition in ADDED_COLUMNS:
                if column not in self._columns(table):
                    self.db.execute('ALTER TABLE %s ADD COLUMN %s %s' % (table, column, definition))
            prefix = self.setting('number_prefix', DEFAULT_NUMBER_PREFIX)
            for r in self.all('SELECT id, number, year FROM memos WHERE reg_number IS NULL'):
                self.db.execute('UPDATE memos SET reg_number = ? WHERE id = ?',
                                (docx.format_reg_number(prefix, r['number'], r['year']), r['id']))
            # У записок прошлых версий был один адресат в колонках memos
            self.db.execute(
                'INSERT INTO memo_addressees (memo_id, ord, addressee_id, department, head_position, head_name) '
                'SELECT id, 0, addressee_id, addressee_department, addressee_position, addressee_name FROM memos '
                'WHERE id NOT IN (SELECT memo_id FROM memo_addressees)')
            self.db.execute('COMMIT')
        except BaseException:
            self.db.execute('ROLLBACK')
            raise

    def _data_version(self):
        return self.db.execute('PRAGMA data_version').fetchone()[0]

    def changed(self):
        """True, если с прошлой проверки базу изменил другой пользователь."""
        version = self._data_version()
        if version == self._version:
            return False
        self._version = version
        return True

    def all(self, sql, params=()):
        return [dict(r) for r in self.db.execute(sql, params)]

    def one(self, sql, params=()):
        rows = self.all(sql, params)
        return rows[0] if rows else None

    def run(self, sql, params=()):
        return self.db.execute(sql, params).lastrowid

    def write(self, fn):
        """Выполняет fn(store) в транзакции под блокировкой и возвращает её результат."""
        with self._lock():
            self.db.execute('BEGIN IMMEDIATE')
            try:
                result = fn(self)
                self.db.execute('COMMIT')
            except BaseException:
                self.db.execute('ROLLBACK')
                raise
        self._version = self._data_version()
        return result

    # ---------- Настройки ----------

    def setting(self, key, fallback=''):
        row = self.one('SELECT value FROM settings WHERE key = ?', (key,))
        return row['value'] if row else fallback

    def set_setting(self, key, value):
        self.run('INSERT INTO settings (key, value) VALUES (?, ?) '
                 'ON CONFLICT(key) DO UPDATE SET value = excluded.value', (key, value))

    @property
    def output_dir(self):
        return self.setting('output_dir', DEFAULT_OUTPUT_DIR)

    @property
    def city(self):
        return self.setting('city', '')

    @property
    def org_name(self):
        return self.setting('org_name', '')

    @property
    def number_prefix(self):
        return self.setting('number_prefix', DEFAULT_NUMBER_PREFIX)

    @property
    def signer(self):
        return self.one('SELECT * FROM employees WHERE is_signer = 1 ORDER BY id LIMIT 1')

    def memo_addressees(self, memo_id):
        return self.all('SELECT * FROM memo_addressees WHERE memo_id = ? ORDER BY ord', (memo_id,))

    # ---------- Файлы документов ----------

    def abs_path(self, rel_path):
        return os.path.join(self.root, *split_path(rel_path))

    def document_path(self, rel_path):
        """Полный путь к документу Word или None, если файла нет."""
        if not rel_path:
            return None
        path = self.abs_path(rel_path)
        return path if os.path.isfile(path) else None

    def save_document(self, memo, overwrite_path=None):
        """Записывает документ Word и возвращает его путь внутри общей папки."""
        data = docx.build_docx(memo, self.memo_addressees(memo['id']), self.org_name)
        if overwrite_path:
            rel_parts = split_path(overwrite_path)
        else:
            # <папка документов>/<год>/<номер СЗ>/<номер СЗ>.docx
            base = docx.file_base_name(memo)
            rel_dir = split_path(self.output_dir) + [str(memo['year']), base]
            directory = os.path.join(self.root, *rel_dir)
            name = base + '.docx'
            n = 2
            while os.path.exists(os.path.join(directory, name)):
                name = '%s (%d).docx' % (base, n)
                n += 1
            rel_parts = rel_dir + [name]
        path = os.path.join(self.root, *rel_parts)
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, 'wb') as f:
                f.write(data)
        except PermissionError:
            raise StoreError('Не удалось записать %s: файл открыт в Word или нет прав на запись. '
                             'Закройте документ и повторите.' % path)
        return '/'.join(rel_parts)
