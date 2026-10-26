// Хранилище: база SQLite (sql.js) в файле data/memos.db общей папки.
//
// Каждый пользователь держит копию базы в памяти браузера. Запись идёт так:
// захватить файл-блокировку data/memos.lock -> перечитать базу с диска ->
// применить изменения -> записать файл -> снять блокировку. Так изменения
// разных пользователей не затирают друг друга.
(function (root) {
  'use strict';

  const DATA_DIR = 'data';
  const DB_FILE = 'memos.db';
  const LOCK_FILE = 'memos.lock';
  const LOCK_STALE_MS = 60000; // блокировка старше минуты считается брошенной
  const LOCK_TIMEOUT_MS = 20000;
  const DEFAULT_OUTPUT_DIR = 'Служебные записки';

  // Справочные значения копируются в записку на момент создания,
  // чтобы правка справочников не меняла уже выпущенные документы.
  const SCHEMA = `
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
  `;

  // Колонки, добавленные после первой версии: [таблица, колонка, определение]
  const ADDED_COLUMNS = [
    ['memos', 'reg_number', 'TEXT'],
    ['memos', 'attachment', "TEXT NOT NULL DEFAULT ''"],
    ['memos', 'signer_id', 'INTEGER'],
    ['memos', 'signer_name', "TEXT NOT NULL DEFAULT ''"],
    ['memos', 'signer_position', "TEXT NOT NULL DEFAULT ''"],
    ['memos', 'city', "TEXT NOT NULL DEFAULT ''"],
    ['memos', 'org_name', 'TEXT'], // NULL у записок прошлых версий: шапка берётся из настроек
    ['employees', 'position', "TEXT NOT NULL DEFAULT ''"],
    ['addressees', 'head_name_dative', "TEXT NOT NULL DEFAULT ''"],
    ['employees', 'is_signer', 'INTEGER NOT NULL DEFAULT 0'], // подписант - не больше одного
    ['memo_addressees', 'head_name_dative', "TEXT NOT NULL DEFAULT ''"],
  ];

  const DEFAULT_NUMBER_PREFIX = 'НТЦ';

  function columns(db, table) {
    const cols = [];
    const stmt = db.prepare(`PRAGMA table_info(${table})`);
    while (stmt.step()) cols.push(stmt.getAsObject().name);
    stmt.free();
    return cols;
  }

  // Создаёт недостающие таблицы и колонки в базе, созданной прошлой версией
  function prepare(db) {
    db.exec(SCHEMA);
    for (const [table, column, def] of ADDED_COLUMNS) {
      if (!columns(db, table).includes(column)) {
        db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${def}`);
      }
    }
    const prefixRow = db.exec("SELECT value FROM settings WHERE key = 'number_prefix'");
    const prefix = prefixRow.length ? prefixRow[0].values[0][0] : DEFAULT_NUMBER_PREFIX;
    const old = db.exec('SELECT id, number, year FROM memos WHERE reg_number IS NULL');
    if (old.length) {
      for (const [id, number, year] of old[0].values) {
        db.run('UPDATE memos SET reg_number = ? WHERE id = ?',
          [root.SZDocx.formatRegNumber(prefix, number, year), id]);
      }
    }
    // У записок прошлых версий был один адресат в колонках memos
    db.exec(`INSERT INTO memo_addressees (memo_id, ord, addressee_id, department, head_position, head_name)
      SELECT id, 0, addressee_id, addressee_department, addressee_position, addressee_name FROM memos
      WHERE id NOT IN (SELECT memo_id FROM memo_addressees)`);
  }

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function isNotFound(e) {
    return e && (e.name === 'NotFoundError' || e.code === 'ENOENT');
  }

  function splitPath(path) {
    return String(path).split(/[\\/]+/).filter((p) => p && p !== '.');
  }

  async function getDir(dir, path, create) {
    for (const part of splitPath(path)) {
      dir = await dir.getDirectoryHandle(part, { create });
    }
    return dir;
  }

  async function readFile(dir, name) {
    try {
      return await (await dir.getFileHandle(name)).getFile();
    } catch (e) {
      if (isNotFound(e)) return null;
      throw e;
    }
  }

  async function writeFile(dir, name, data) {
    const handle = await dir.getFileHandle(name, { create: true });
    const writable = await handle.createWritable();
    await writable.write(data);
    await writable.close();
  }

  class Store {
    constructor(SQL, rootHandle) {
      this.SQL = SQL;
      this.root = rootHandle;
      this.db = null;
      this.stamp = null;
    }

    async open() {
      this.dataDir = await getDir(this.root, DATA_DIR, true);
      if (!(await this.refresh(true))) {
        // Первый запуск: создаём пустую базу
        await this.write(() => {});
      }
    }

    // Перечитывает базу, если файл изменился. Возвращает true, если база обновлена.
    async refresh(force) {
      const file = await readFile(this.dataDir, DB_FILE);
      if (!file) return false;
      const stamp = `${file.lastModified}:${file.size}`;
      if (!force && this.db && stamp === this.stamp) return false;
      const db = new this.SQL.Database(new Uint8Array(await file.arrayBuffer()));
      prepare(db);
      if (this.db) this.db.close();
      this.db = db;
      this.stamp = stamp;
      return true;
    }

    all(sql, params = []) {
      const stmt = this.db.prepare(sql);
      try {
        stmt.bind(params);
        const rows = [];
        while (stmt.step()) rows.push(stmt.getAsObject());
        return rows;
      } finally {
        stmt.free();
      }
    }

    one(sql, params = []) {
      return this.all(sql, params)[0] || null;
    }

    run(sql, params = []) {
      this.db.run(sql, params);
      return this.one('SELECT last_insert_rowid() AS id').id;
    }

    setting(key, fallback) {
      const row = this.one('SELECT value FROM settings WHERE key = ?', [key]);
      return row ? row.value : fallback;
    }

    get outputDir() {
      return this.setting('output_dir', DEFAULT_OUTPUT_DIR);
    }

    get city() {
      return this.setting('city', '');
    }

    get signer() {
      return this.one('SELECT * FROM employees WHERE is_signer = 1 ORDER BY id LIMIT 1');
    }

    memoAddressees(memoId) {
      return this.all('SELECT * FROM memo_addressees WHERE memo_id = ? ORDER BY ord', [memoId]);
    }

    get orgName() {
      return this.setting('org_name', '');
    }

    get numberPrefix() {
      return this.setting('number_prefix', DEFAULT_NUMBER_PREFIX);
    }

    setSetting(key, value) {
      this.run('INSERT INTO settings (key, value) VALUES (?, ?) ' +
        'ON CONFLICT(key) DO UPDATE SET value = excluded.value', [key, value]);
    }

    // Выполняет fn(store) в транзакции на свежей копии базы и сохраняет файл.
    async write(fn) {
      const token = await this._lock();
      try {
        await this.refresh(true);
        if (!this.db) {
          this.db = new this.SQL.Database();
          prepare(this.db);
        }
        this.db.exec('BEGIN');
        let result;
        try {
          result = fn(this);
          this.db.exec('COMMIT');
        } catch (e) {
          this.db.exec('ROLLBACK');
          throw e;
        }
        await writeFile(this.dataDir, DB_FILE, this.db.export());
        const file = await readFile(this.dataDir, DB_FILE);
        this.stamp = `${file.lastModified}:${file.size}`;
        return result;
      } catch (e) {
        this.stamp = null; // при следующем чтении перечитать с диска
        throw e;
      } finally {
        await this._unlock(token);
      }
    }

    async _readLock() {
      const file = await readFile(this.dataDir, LOCK_FILE);
      if (!file) return null;
      try {
        return JSON.parse(await file.text());
      } catch {
        return { time: file.lastModified };
      }
    }

    async _lock() {
      const token = Math.random().toString(36).slice(2) + Date.now().toString(36);
      const deadline = Date.now() + LOCK_TIMEOUT_MS;
      for (;;) {
        const current = await this._readLock();
        if (!current || Date.now() - (current.time || 0) > LOCK_STALE_MS) {
          await writeFile(this.dataDir, LOCK_FILE,
            JSON.stringify({ token, time: Date.now() }));
          // Двое могли записать блокировку одновременно: побеждает тот,
          // чей токен остался в файле.
          await sleep(200);
          const check = await this._readLock();
          if (check && check.token === token) return token;
        }
        if (Date.now() > deadline) {
          throw new Error('База занята другим пользователем. Повторите через минуту.');
        }
        await sleep(300 + Math.random() * 500);
      }
    }

    async _unlock(token) {
      try {
        const current = await this._readLock();
        if (current && current.token === token) {
          await this.dataDir.removeEntry(LOCK_FILE);
        }
      } catch {
        // Блокировка устареет сама
      }
    }

    // ---------- Файлы документов ----------

    async saveDocument(memo, overwritePath) {
      const bytes = root.SZDocx.buildDocx(memo, {
        orgName: this.orgName, addressees: this.memoAddressees(memo.id),
      });
      if (overwritePath) {
        const parts = splitPath(overwritePath);
        const dir = await getDir(this.root, parts.slice(0, -1).join('/'), true);
        await writeFile(dir, parts[parts.length - 1], bytes);
        return overwritePath;
      }
      // <папка документов>/<год>/<номер СЗ>/<номер СЗ>.docx
      const base = root.SZDocx.fileBaseName(memo);
      const relDir = [...splitPath(this.outputDir), String(memo.year), base];
      const dir = await getDir(this.root, relDir.join('/'), true);
      let name = `${base}.docx`;
      for (let n = 2; await readFile(dir, name); n++) name = `${base} (${n}).docx`;
      await writeFile(dir, name, bytes);
      return [...relDir, name].join('/');
    }

    async readDocument(relPath) {
      if (!relPath) return null;
      const parts = splitPath(relPath);
      try {
        const dir = await getDir(this.root, parts.slice(0, -1).join('/'), false);
        return await readFile(dir, parts[parts.length - 1]);
      } catch (e) {
        if (isNotFound(e)) return null;
        throw e;
      }
    }
  }

  root.SZStore = { Store, DEFAULT_OUTPUT_DIR, DEFAULT_NUMBER_PREFIX };
})(typeof window !== 'undefined' ? window : globalThis);
