r"""Перенос «Номер СЗ 1С» из Excel в существующие СЗ по колонке «Номер СЗ».

Пример PowerShell (закройте приложение у пользователей перед импортом):
  & ".\environment\python\python.exe" ".\tools\import_sed_numbers.py" "C:\журнал.xlsx" "."

Вместо общей папки можно указать data/memos.db. По умолчанию берётся папка
программы. --dry-run только формирует отчёт; --yes подтверждает запись без вопроса.
Существующие номера очищаются для пустых и нераспознанных значений Excel.
Word, PDF, пути сканов, номера СЗ и остальные поля не изменяются.
"""
import argparse
import csv
import datetime
import os
from pathlib import Path
import re
import sqlite3
import sys
import unicodedata
import uuid

PROGRAM_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROGRAM_DIR))
from sz.store import FileLock, StoreError  # noqa: E402

SEPARATOR = r'[/\\|¦_.,:;?*\-\u2010-\u2015\u2212\u2044\u2215\s]+'
NUMBER = re.compile(r'([0-9]+)' + SEPARATOR + r'([0-9]+)' + SEPARATOR +
                    r'([0-9]+)' + SEPARATOR + r'СЗ', re.IGNORECASE)
REG = re.compile(r'(.+?)\s*[-\u2010-\u2015\u2212]\s*([0-9]{1,6})\s*'
                 r'[-\u2010-\u2015\u2212]\s*([0-9]{2})')


def text(value):
    return unicodedata.normalize('NFKC', str(value if value is not None else '')).strip()


def normalize_sed(value):
    value = text(value)
    marker = re.sub(r'[\s.]+', '', value.casefold())
    if marker in ('вбумаге', 'вотд40'):
        return '', 1, 'В бумаге'
    match = NUMBER.fullmatch(value)
    if match:
        number = '/'.join(match.groups()) + '/СЗ'
        if len(number) <= 120:
            return number, 0, 'номер'
    return '', 0, 'пусто' if not value else 'нераспознано → пусто'


def reg_key(value):
    match = REG.fullmatch(text(value))
    if not match:
        return None
    return (' '.join(match[1].split()).casefold(), int(match[2]), int(match[3]))


def read_excel(path, sheet=None):
    if Path(path).suffix.lower() not in ('.xlsx', '.xlsm'):
        raise StoreError('Сохраните журнал в формате .xlsx или .xlsm.')
    try:
        import openpyxl
    except ImportError:
        raise StoreError('Нет openpyxl. Запускайте скрипт через environment/python/python.exe.')
    def load(data_only):
        try:
            return openpyxl.load_workbook(path, read_only=True, data_only=data_only)
        except Exception as error:
            raise StoreError('Не удалось прочитать Excel: %s' % error) from error

    formulas = load(False)
    cached = None
    try:
        cached = load(True)
        name = sheet or formulas.sheetnames[0]
        if name not in formulas.sheetnames:
            raise StoreError('В Excel нет листа «%s».' % name)
        wanted = ['номер сз', 'номер сз 1с']
        columns = None
        entries = []
        for line, (raw, values) in enumerate(zip(formulas[name].iter_rows(),
                                                cached[name].iter_rows(values_only=True)), 1):
            if columns is None:
                keys = [re.sub(r'\s+', ' ', text(c.value).casefold()).replace('1c', '1с') for c in raw]
                if all(key in keys for key in wanted):
                    if any(keys.count(key) != 1 for key in wanted):
                        raise StoreError('В строке заголовков повторяются нужные колонки.')
                    columns = [keys.index(key) for key in wanted]
                elif line >= 30:
                    raise StoreError('В первых 30 строках не найдены «Номер СЗ» и «Номер СЗ 1С».')
                continue
            reg, sed = [values[i] if i < len(values) else None for i in columns]
            uncached = any(raw[i].data_type == 'f' and values[i] is None
                           for i in columns if i < len(raw))
            if not text(reg) and not text(sed) and not uncached:
                continue
            entries.append({'line': line, 'reg': reg, 'raw': sed,
                            'error': 'Формула без сохранённого результата: пересохраните Excel.' if uncached else ''})
        if columns is None:
            raise StoreError('Не найдены колонки «Номер СЗ» и «Номер СЗ 1С».')
        return name, entries
    finally:
        formulas.close()
        if cached:
            cached.close()


def connect(path):
    if not Path(path).is_file():
        raise StoreError('База не найдена: %s. Новая база создаваться не будет.' % path)
    db = sqlite3.connect(str(path), timeout=20, isolation_level=None)
    db.row_factory = sqlite3.Row
    required = {'id', 'reg_number', 'sed_number', 'in_paper', 'scan_path'}
    if not required.issubset({r['name'] for r in db.execute('PRAGMA table_info(memos)')}):
        db.close()
        raise StoreError('В базе нет нужных полей. Сначала запустите актуальную версию приложения.')
    return db


def snapshot(row):
    return tuple(row[key] for key in ('reg_number', 'sed_number', 'in_paper', 'scan_path'))


def make_plan(db, entries):
    index = {}
    for row in db.execute('SELECT id, reg_number, sed_number, in_paper, scan_path FROM memos'):
        key = reg_key(row['reg_number'])
        if key:
            index.setdefault(key, []).append(dict(row))
    grouped, report, plan = {}, [], []
    for entry in entries:
        result = dict(entry, status='', old='', new='')
        report.append(result)
        key = reg_key(entry['reg'])
        if entry['error']:
            result['status'] = entry['error']
        elif not key:
            result['status'] = 'Пропуск: неверный номер СЗ'
        elif len(index.get(key, [])) != 1:
            result['status'] = 'Пропуск: СЗ не найдена' if key not in index else 'Пропуск: неоднозначный номер в БД'
        else:
            result['memo'] = index[key][0]
            result['normalized'] = normalize_sed(entry['raw'])
            grouped.setdefault(key, []).append(result)
    for group in grouped.values():
        variants = {item['normalized'][:2] for item in group}
        if len(variants) > 1:
            for item in group:
                item['status'] = 'Пропуск: противоречивые дубликаты в Excel'
            continue
        for i, item in enumerate(group):
            memo = item['memo']
            number, paper, reason = item['normalized']
            item['old'] = 'В бумаге' if memo['in_paper'] else memo['sed_number']
            item['new'] = 'В бумаге' if paper else number
            if i:
                item['status'] = 'Пропуск: повтор той же записи'
            elif (memo['sed_number'], memo['in_paper']) == (number, paper):
                item['status'] = 'Без изменений'
            else:
                item['status'] = 'Изменить: ' + reason
                plan.append({'id': memo['id'], 'before': snapshot(memo), 'number': number, 'paper': paper})
    return plan, report


def write_report(path, report):
    with open(path, 'w', encoding='utf-8-sig', newline='') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(['Строка Excel', 'Номер СЗ', 'Номер СЗ 1С (исходный)', 'Было', 'Станет', 'Результат'])
        for item in report:
            # Значения Excel — только текст, без активных формул при открытии CSV.
            values = [item['line'], item['reg'], item['raw'], item['old'], item['new'], item['status']]
            writer.writerow(["'" + str(v) if str(v or '').lstrip().startswith(('=', '+', '-', '@')) else v
                             for v in values])


def apply_plan(path, plan):
    if not plan:
        return None
    path = Path(path)
    stamp = datetime.datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:8]
    backup = path.with_name('memos-before-sed-import-%s.db' % stamp)
    temporary = str(backup) + '.tmp'
    with FileLock(str(path.parent / 'memos.lock')):
        db = connect(path)
        try:
            # backup() вызывается до BEGIN IMMEDIATE, иначе SQLite может ждать
            # окончания своей же транзакции. Общая блокировка удерживается всё время.
            target = sqlite3.connect(temporary)
            try:
                db.backup(target)
                if target.execute('PRAGMA integrity_check').fetchall() != [('ok',)]:
                    raise StoreError('Резервная копия не прошла проверку. Импорт отменён.')
            finally:
                target.close()
            os.replace(temporary, backup)
            db.execute('BEGIN IMMEDIATE')
            try:
                for change in plan:
                    row = db.execute('SELECT reg_number, sed_number, in_paper, scan_path FROM memos WHERE id = ?',
                                     (change['id'],)).fetchone()
                    if row is None or snapshot(row) != change['before']:
                        raise StoreError('Данные СЗ изменились после отчёта. Импорт целиком отменён; запустите заново.')
                db.executemany('UPDATE memos SET sed_number = ?, in_paper = ? WHERE id = ?',
                               [(item['number'], item['paper'], item['id']) for item in plan])
                db.execute('COMMIT')
            except BaseException:
                db.execute('ROLLBACK')
                raise
        finally:
            db.close()
            if os.path.exists(temporary):
                os.unlink(temporary)
    return backup


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('excel', help='Файл .xlsx/.xlsm')
    parser.add_argument('database', nargs='?', default=str(PROGRAM_DIR), help='Общая папка или файл memos.db')
    parser.add_argument('--sheet', help='Название листа; по умолчанию первый лист')
    parser.add_argument('--report', help='Путь к полному CSV-отчёту')
    parser.add_argument('--dry-run', action='store_true', help='Только отчёт, без записи в БД')
    parser.add_argument('--yes', action='store_true', help='Подтвердить все изменения, включая очистку номеров')
    args = parser.parse_args(argv)
    db_path = Path(args.database).resolve()
    if db_path.is_dir():
        db_path = db_path / 'data' / 'memos.db'
    db = connect(db_path)
    try:
        db.execute('PRAGMA query_only = ON')
        sheet, entries = read_excel(args.excel, args.sheet)
        plan, report = make_plan(db, entries)
    finally:
        db.close()
    stamp = datetime.datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:8]
    report_path = Path(args.report) if args.report else db_path.parent / ('sed-import-report-%s.csv' % stamp)
    if report_path.resolve() in (db_path.resolve(), Path(args.excel).resolve()):
        raise StoreError('Отчёт должен быть отдельным файлом, а не исходным Excel или базой.')
    write_report(report_path, report)
    cleared = sum(not item['number'] and not item['paper'] for item in plan)
    print('База: %s\nЛист: %s\nПолный отчёт: %s' % (db_path, sheet, report_path))
    print('Изменений: %d; из них очисток: %d; отметок «В бумаге»: %d.' % (
        len(plan), cleared, sum(item['paper'] for item in plan)))
    for item in report[:30]:
        print('Строка %s | %s | %r → %r | %s' % (item['line'], item['reg'], item['old'], item['new'], item['status']))
    if args.dry_run or not plan:
        print('База не изменена.')
        return 0
    print('ВНИМАНИЕ: пустые и нераспознанные значения очистят прежний номер/отметку в соответствующих СЗ.')
    if not args.yes and input('Применить изменения? (да/нет): ').strip().casefold() not in ('да', 'д', 'yes', 'y'):
        print('База не изменена.')
        return 0
    backup = apply_plan(db_path, plan)
    print('Обновлено СЗ: %d. Резервная копия: %s' % (len(plan), backup))
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except (StoreError, OSError, sqlite3.Error, ValueError) as error:
        sys.exit('Ошибка: %s' % error)
