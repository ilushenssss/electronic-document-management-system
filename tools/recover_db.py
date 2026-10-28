"""Проверка и восстановление базы служебных записок (data/memos.db).

Запуск (программа «Служебные записки» должна быть закрыта у всех пользователей):

    environment\\python\\python.exe tools\\recover_db.py "\\\\server\\Общие\\СЗ\\data\\memos.db"

Без пути берётся data\\memos.db рядом с программой.

Исходный файл не изменяется. Скрипт:
1. Копирует базу вместе со служебными файлами SQLite (memos.db-journal, memos.db-wal)
   в папку «восстановление-<дата>» рядом с базой и проверяет копию.
2. Если база исправна - сообщает об этом и больше ничего не делает.
3. Иначе переносит всё, что читается, в новый файл memos-восстановлена.db:
   сначала таблицами целиком, а где не выходит - по одной строке, пропуская испорченные.
"""
import datetime
import os
import shutil
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from sz.store import Store  # noqa: E402

# Порядок важен только для отчёта: справочники, затем записки
TABLES = ['settings', 'topics', 'employees', 'addressees', 'memos', 'memo_addressees', 'write_log']
SIDE_FILES = ('-journal', '-wal', '-shm')
SCAN_GAP = 5000  # столько пропущенных номеров строк подряд - конец таблицы


def check(db):
    try:
        rows = [r[0] for r in db.execute('PRAGMA integrity_check')]
    except sqlite3.DatabaseError as e:
        return [str(e)]
    return [] if rows == ['ok'] else rows


def columns(db, table):
    try:
        return [r[1] for r in db.execute('PRAGMA table_info(%s)' % table)]
    except sqlite3.DatabaseError:
        return []


def copy_table(src, dst, table):
    """Переносит строки таблицы. Возвращает (перенесено, не прочитано)."""
    dst_cols = columns(dst, table)
    cols = [c for c in (columns(src, table) or dst_cols) if c in dst_cols]
    if not cols:
        return 0, None
    names = ', '.join(cols)
    insert = 'INSERT OR IGNORE INTO %s (%s) VALUES (%s)' % (table, names, ', '.join('?' * len(cols)))
    try:
        rows = src.execute('SELECT %s FROM %s' % (names, table)).fetchall()
        dst.executemany(insert, rows)
        return len(rows), 0
    except sqlite3.DatabaseError:
        pass
    # Таблица целиком не читается: по одной строке
    saved = failed = 0
    rowid = 0
    gap = 0
    try:
        last = src.execute('SELECT MAX(rowid) FROM %s' % table).fetchone()[0]
    except sqlite3.DatabaseError:
        last = None
    while (rowid < last) if last else (gap < SCAN_GAP):
        rowid += 1
        try:
            row = src.execute('SELECT %s FROM %s WHERE rowid = ?' % (names, table), (rowid,)).fetchone()
        except sqlite3.DatabaseError:
            failed += 1
            gap = 0
            continue
        if row is None:
            gap += 1
            continue
        gap = 0
        dst.execute(insert, row)
        saved += 1
    return saved, failed


def main():
    if len(sys.argv) > 1:
        path = os.path.abspath(sys.argv[1])
    else:
        path = os.path.join(os.path.dirname(HERE), 'data', 'memos.db')
    if not os.path.isfile(path):
        sys.exit('Нет файла %s' % path)
    if os.path.exists(os.path.join(os.path.dirname(path), 'memos.lock')):
        print('Внимание: есть data\\memos.lock - возможно, кто-то сейчас сохраняет данные.\n'
              'Закройте программу у всех пользователей и запустите проверку ещё раз.\n')

    # Работаем только с копией: исходные файлы остаются как есть
    base = os.path.join(os.path.dirname(path), 'восстановление-%s' % datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))
    work = base
    n = 2
    while os.path.exists(work):
        work = '%s-%d' % (base, n)
        n += 1
    os.makedirs(work)
    copy = os.path.join(work, 'memos.db')
    shutil.copy2(path, copy)
    for suffix in SIDE_FILES:
        if os.path.exists(path + suffix):
            print('Найден служебный файл SQLite %s - он скопирован вместе с базой.' % os.path.basename(path + suffix))
            shutil.copy2(path + suffix, copy + suffix)
    print('Копия для проверки: %s\n' % copy)

    src = sqlite3.connect(copy)  # при открытии SQLite сам доведёт до конца прерванную запись
    problems = check(src)
    if not problems:
        counts = []
        for t in ('memos', 'employees', 'addressees', 'topics'):
            try:
                counts.append('%s: %d' % (t, src.execute('SELECT COUNT(*) FROM %s' % t).fetchone()[0]))
            except sqlite3.DatabaseError:
                pass
        src.close()
        print('База ИСПРАВНА (integrity_check: ok). %s' % ', '.join(counts))
        print('Сообщение об ошибке выдавала другая программа. Исходную базу можно не трогать;\n'
              'папку %s можно удалить.' % work)
        return
    print('База ПОВРЕЖДЕНА. Первые сообщения проверки:')
    for p in problems[:10]:
        print('  ', p)

    # Новая пустая база той же структуры, что создаёт программа
    root = os.path.join(work, 'новая')
    os.makedirs(root)
    store = Store(root)
    store.open()
    store.close()
    target = os.path.join(work, 'memos-восстановлена.db')
    shutil.move(os.path.join(root, 'data', 'memos.db'), target)
    shutil.rmtree(root, ignore_errors=True)

    dst = sqlite3.connect(target)
    dst.execute('DELETE FROM memo_addressees')  # программа могла заполнить их при создании
    print('\nПеренос данных:')
    for table in TABLES:
        saved, failed = copy_table(src, dst, table)
        if failed is None:
            print('  %-16s не читается совсем' % table)
        else:
            print('  %-16s перенесено %d%s' % (table, saved, ', не прочитано %d' % failed if failed else ''))
    dst.commit()
    ok = not check(dst)
    dst.close()
    src.close()
    print('\nВосстановленная база: %s (%s)' % (target, 'проверка пройдена' if ok else 'проверка НЕ пройдена'))
    print('Дальше:\n'
          '  1. Переименуйте повреждённую data\\memos.db, например в memos-повреждена.db.\n'
          '  2. Скопируйте memos-восстановлена.db в data\\ под именем memos.db.\n'
          '  3. Откройте программу и проверьте реестр. Записки, которых не хватает, можно\n'
          '     сверить с документами Word в папке «Служебные записки».')


if __name__ == '__main__':
    main()
