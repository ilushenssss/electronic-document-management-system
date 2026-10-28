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
4. Сверяет реестр с папками документов Word (<папка документов>/<год>/<номер СЗ>/). Для
   каждого документа без записи добавляет запись «нет данных» с тем же номером и ссылкой
   на документ: номера не будут выданы повторно, а записку можно открыть в Word из реестра.
"""
import datetime
import os
import re
import shutil
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from sz.store import DEFAULT_OUTPUT_DIR, Store, split_path  # noqa: E402

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
    gap = 0  # попыток подряд без прочитанной строки (пропуски и ошибки)
    errors = 0  # ошибки после последней прочитанной строки
    try:
        last = src.execute('SELECT MAX(rowid) FROM %s' % table).fetchone()[0]
    except sqlite3.DatabaseError:
        last = None
    while (rowid < last) if last else (gap < SCAN_GAP):
        rowid += 1
        try:
            row = src.execute('SELECT %s FROM %s WHERE rowid = ?' % (names, table), (rowid,)).fetchone()
        except sqlite3.DatabaseError:
            errors += 1
            gap += 1
            continue
        if row is None:
            gap += 1
            continue
        gap = 0
        failed += errors
        errors = 0
        dst.execute(insert, row)
        saved += 1
    if last:
        failed += errors
    if not saved and errors:
        return 0, None  # не прочиталось ни одной строки
    return saved, failed


def word_documents(program_root, output_dir):
    """Документы Word по папкам <папка документов>/<год>/<номер СЗ>/: {номер СЗ: (год, номер, путь)}."""
    base_parts = split_path(output_dir)
    base = os.path.join(program_root, *base_parts)
    found = {}
    if not os.path.isdir(base):
        return found
    for year in sorted(os.listdir(base)):
        year_dir = os.path.join(base, year)
        if not (year.isdigit() and os.path.isdir(year_dir)):
            continue
        for name in sorted(os.listdir(year_dir)):
            m = re.match(r'^.+-(\d+)-\d{2}$', name)
            folder = os.path.join(year_dir, name)
            if not m or not os.path.isdir(folder):
                continue
            files = sorted(f for f in os.listdir(folder) if f.lower().endswith('.docx'))
            if not files:
                continue
            main_file = name + '.docx' if name + '.docx' in files else files[0]
            found[name] = (int(year), int(m.group(1)), '/'.join(base_parts + [year, name, main_file]))
    return found


def add_missing(dst, program_root):
    """Записи «нет данных» для документов Word, которых нет в реестре. Возвращает их номера."""
    row = dst.execute("SELECT value FROM settings WHERE key = 'output_dir'").fetchone()
    docs = word_documents(program_root, row[0] if row else DEFAULT_OUTPUT_DIR)
    known = {r[0] for r in dst.execute('SELECT reg_number FROM memos')}
    taken = {(r[0], r[1]) for r in dst.execute('SELECT year, number FROM memos')}
    added = []
    for reg, (year, number, rel) in sorted(docs.items(), key=lambda d: (d[1][0], d[1][1])):
        if reg in known or (year, number) in taken:
            continue
        mtime = os.path.getmtime(os.path.join(program_root, *split_path(rel)))
        dst.execute(
            'INSERT INTO memos (number, year, reg_number, memo_date, topic_name, addressee_department, '
            'addressee_position, addressee_name, body, executor_name, file_path) '
            "VALUES (?, ?, ?, ?, 'нет данных (восстановление)', '—', '—', '—', ?, '—', ?)",
            (number, year, reg, datetime.date.fromtimestamp(mtime).isoformat(),
             'Запись не прочиталась из повреждённой базы. Текст записки - в документе Word.', rel))
        added.append(reg)
    return added


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
    # Документы Word, которых нет в восстановленном реестре
    program_root = os.path.dirname(os.path.dirname(path))
    added = add_missing(dst, program_root)
    dst.commit()
    if added:
        print('\nДокументов Word без записи в реестре: %d. Для них добавлены записи «нет данных»\n'
              '(дата - дата файла), чтобы номера не выдавались повторно:' % len(added))
        print('  ' + ', '.join(added[:40]) + (' и ещё %d' % (len(added) - 40) if len(added) > 40 else ''))
    ok = not check(dst)
    dst.close()
    src.close()
    print('\nВосстановленная база: %s (%s)' % (target, 'проверка пройдена' if ok else 'проверка НЕ пройдена'))
    print('Дальше:\n'
          '  1. Переименуйте повреждённую data\\memos.db, например в memos-повреждена.db.\n'
          '  2. Скопируйте memos-восстановлена.db в data\\ под именем memos.db.\n'
          '  3. Откройте программу и проверьте реестр. У записей «нет данных» откройте документ Word:\n'
          '     в нём текст записки. Записки, для которых не было и документа Word, не восстановить.')


if __name__ == '__main__':
    main()
