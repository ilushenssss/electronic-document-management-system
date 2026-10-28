"""Перенос журнала служебных записок из Excel (.xlsx) в базу программы.

Запуск (в PowerShell; пути замените своими):

    & "\\\\server\\СЗ\\environment\\python\\python.exe" "\\\\server\\СЗ\\tools\\import_journal.py" "C:\\журнал.xlsx" "\\\\server\\СЗ"

Второй параметр - общая папка программы (где лежит data\\memos.db) или сам файл memos.db.
Без него берётся папка, где лежит программа.

Берётся первый лист. Колонки ищутся по заголовкам (регистр и лишние пробелы не важны):
  «Объект»             - тематика работ (новая тематика добавляется в справочник)
  «Номер СЗ»           - номер, например НТЦ-0001-26; год берётся из номера
  «Краткое содержание» - текст записки
  «Дата»               - дата записки
  «Ответственный»      - исполнитель (ищется в справочнике «Сотрудники»)
  «Кому»               - адресаты через «;», «,» или с новой строки (ищутся в справочнике «Адресаты»)
Колонки «Номер СЗ в ИСУ КБ» и «Номер СЗ 1С» и все прочие не используются.

Сначала скрипт показывает, что будет внесено, и спрашивает подтверждение. Перед записью
делается копия базы (data\\memos-до-импорта-<дата>.db). Номера, которые уже есть в базе,
пропускаются, поэтому запускать повторно безопасно. Документы Word не создаются: их можно
сформировать в программе кнопкой «Сформировать Word заново».
"""
import datetime
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROGRAM_DIR = os.path.dirname(HERE)
sys.path.insert(0, PROGRAM_DIR)

from sz import docx  # noqa: E402
from sz.store import Store, StoreError  # noqa: E402

COLUMNS = {
    'topic': 'Объект',
    'reg': 'Номер СЗ',
    'body': 'Краткое содержание',
    'date': 'Дата',
    'executor': 'Ответственный',
    'to': 'Кому',
}
HEADER_SEARCH_ROWS = 30  # в скольких первых строках искать заголовки
EXCEL_EPOCH = datetime.date(1899, 12, 30)


def norm(text):
    """Строка для сравнения: без регистра, «ё» как «е», пробелы схлопнуты, «П. П.» как «П.П.»."""
    s = str(text or '').lower().replace('ё', 'е').replace('\xa0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return re.sub(r'\.\s+(?=\S\.)', '.', s)


def header_key(text):
    return re.sub(r'[^0-9a-zа-я]+', ' ', norm(text)).strip()


# ---------- Чтение журнала ----------

def read_sheet(path):
    """Строки первого листа; значения объединённых ячеек повторяются во всех их клетках."""
    try:
        import openpyxl
    except ImportError:
        sys.exit('Нет библиотеки openpyxl. Запускайте скрипт через python.exe из папки environment.')
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    for rng in ws.merged_cells.ranges:
        value = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if r - 1 < len(rows):
                    row = rows[r - 1]
                    while len(row) < c:
                        row.append(None)
                    row[c - 1] = value
    return ws.title, rows


def find_header(rows):
    """Номер строки заголовков и номера нужных колонок."""
    wanted = {key: header_key(title) for key, title in COLUMNS.items()}
    for i, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
        keys = [header_key(v) for v in row]
        found = {key: keys.index(w) for key, w in wanted.items() if w in keys}
        if len(found) == len(wanted):
            return i, found
    titles = [str(v) for row in rows[:HEADER_SEARCH_ROWS] for v in row if v not in (None, '')]
    sys.exit('Не найдены заголовки колонок: %s.\nВ первых строках листа есть: %s'
             % (', '.join('«%s»' % t for t in COLUMNS.values()), ', '.join('«%s»' % t for t in titles[:40])))


def parse_reg(value):
    """«НТЦ-0001-26» (также «НТЦ-1-26», «НТЦ 0001 - 26», с длинным тире) -> (префикс, номер, год)."""
    m = re.match(r'^\s*(.*?\S)\s*[-–—]\s*(\d{1,6})\s*[-–—]\s*(\d{2})\s*$', str(value or ''))
    if not m:
        return None
    return m.group(1), int(m.group(2)), 2000 + int(m.group(3))


def parse_date(value):
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)) and 20000 < value < 80000:  # число вместо даты в Excel
        return EXCEL_EPOCH + datetime.timedelta(days=int(value))
    s = re.sub(r'\s*г\.?\s*$', '', str(value or '').strip())  # «01.02.2026 г.»
    for fmt in ('%d.%m.%Y', '%d.%m.%y', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def cell_text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return '\n'.join(docx.split_lines(str(value)))


# ---------- Сопоставление со справочниками ----------

class Directory:
    def __init__(self, store):
        self.topics = {norm(t['name']): t for t in store.all('SELECT * FROM topics')}
        self.employees = store.all('SELECT * FROM employees')
        self.addressees = store.all('SELECT * FROM addressees')

    def employee(self, text):
        n = norm(text)
        if not n:
            return None
        for e in self.employees:
            if n in (norm(e['full_name']), norm(docx.short_name(e['full_name']))):
                return e
        # Только фамилия, если такой сотрудник один
        same = [e for e in self.employees if norm(e['full_name']).split(' ')[0] == n]
        return same[0] if len(same) == 1 else None

    def addressee(self, text):
        n = norm(text)
        if not n:
            return None
        for a in self.addressees:
            keys = {norm(a['department']), norm(a['head_name']), norm(docx.short_name(a['head_name'])),
                    norm(a['head_name_dative']), norm('%s %s' % (a['head_position'], a['head_name_dative']))}
            if n in keys - {''}:
                return a
        # Отдел или Фамилия И.О. внутри текста, если подходит ровно один адресат
        same = [a for a in self.addressees
                if any(k and k in n for k in (norm(a['department']), norm(a['head_name_dative'])))]
        return same[0] if len(same) == 1 else None


def split_addressees(text):
    return [p.strip() for p in re.split(r'[;,\n]+', text or '') if p.strip()]


# ---------- Разбор строк ----------

def parse_rows(rows, header, cols, directory, existing):
    memos, errors, warnings = [], [], []
    already = 0
    new_topics, unknown_people, unknown_to = set(), set(), set()
    seen = {}
    for i in range(header + 1, len(rows)):
        row = rows[i]
        line = i + 1  # номер строки в Excel

        def get(key):
            c = cols[key]
            return row[c] if c < len(row) else None

        if all(get(k) in (None, '') for k in COLUMNS):
            continue
        reg = parse_reg(get('reg'))
        if not reg:
            errors.append('строка %d: непонятный номер СЗ «%s»' % (line, get('reg') or ''))
            continue
        prefix, number, year = reg
        reg_number = docx.format_reg_number(prefix, number, year)
        date = parse_date(get('date'))
        if not date:
            errors.append('строка %d (%s): непонятная дата «%s»' % (line, reg_number, get('date') or ''))
            continue
        if (year, number) in seen:
            errors.append('строка %d: номер %s уже был в строке %d' % (line, reg_number, seen[(year, number)]))
            continue
        seen[(year, number)] = line
        if (year, number) in existing:
            already += 1
            continue
        if date.year != year:
            warnings.append('строка %d (%s): дата %s из другого года, год взят из номера'
                            % (line, reg_number, date.strftime('%d.%m.%Y')))

        topic_text = cell_text(get('topic')).replace('\n', ' ')
        topic = directory.topics.get(norm(topic_text))
        if topic_text and not topic:
            new_topics.add(topic_text)
        executor_text = cell_text(get('executor')).replace('\n', ' ')
        executor = directory.employee(executor_text)
        if executor_text and not executor:
            unknown_people.add(executor_text)
        addressees = []
        for part in split_addressees(cell_text(get('to'))):
            a = directory.addressee(part)
            if not a:
                unknown_to.add(part)
            # Кого нет в справочнике - как в журнале: и отдел, и строка адресата в документе
            addressees.append(a or {'id': None, 'department': part, 'head_position': part, 'head_name': '',
                                    'head_name_dative': ''})
        for key, title in (('topic', topic_text), ('body', cell_text(get('body'))), ('to', addressees)):
            if not title:
                warnings.append('строка %d (%s): пустая колонка «%s»' % (line, reg_number, COLUMNS[key]))
        memos.append({
            'line': line, 'number': number, 'year': year, 'reg_number': reg_number,
            'memo_date': date.isoformat(), 'topic_text': topic_text, 'topic': topic,
            'body': cell_text(get('body')), 'executor_text': executor_text, 'executor': executor,
            'addressees': addressees,
        })
    return {'memos': memos, 'errors': errors, 'warnings': warnings, 'already': already,
            'new_topics': new_topics, 'unknown_people': unknown_people, 'unknown_to': unknown_to}


# ---------- Запись ----------

def write_memos(store, memos):
    def insert(s):
        topic_ids = {norm(t['name']): t['id'] for t in s.all('SELECT * FROM topics')}
        for m in memos:
            if s.one('SELECT 1 AS x FROM memos WHERE year = ? AND number = ?', (m['year'], m['number'])):
                continue  # успели внести, пока шёл просмотр
            topic_id = None
            if m['topic_text']:
                key = norm(m['topic_text'])
                if key not in topic_ids:
                    topic_ids[key] = s.run('INSERT INTO topics (name) VALUES (?)', (m['topic_text'],))
                topic_id = topic_ids[key]
            e = m['executor'] or {}
            addrs = m['addressees']

            def joined(key):
                return '; '.join(a[key] for a in addrs if a[key])

            memo_id = s.run(
                'INSERT INTO memos (number, year, reg_number, memo_date, topic_id, topic_name, addressee_id, '
                'addressee_department, addressee_position, addressee_name, body, executor_id, executor_name, '
                'executor_phone) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                (m['number'], m['year'], m['reg_number'], m['memo_date'], topic_id, m['topic_text'],
                 addrs[0]['id'] if addrs else None, joined('department'), joined('head_position'),
                 joined('head_name'), m['body'], e.get('id'), e.get('full_name') or m['executor_text'],
                 e.get('phone') or ''))
            for i, a in enumerate(addrs):
                s.run('INSERT INTO memo_addressees (memo_id, ord, addressee_id, department, head_position, '
                      'head_name, head_name_dative) VALUES (?, ?, ?, ?, ?, ?, ?)',
                      (memo_id, i, a['id'], a['department'], a['head_position'], a['head_name'],
                       a['head_name_dative']))
    store.write(insert)


def resolve_root(arg):
    if not arg:
        return PROGRAM_DIR
    path = os.path.abspath(arg)
    if os.path.isfile(path):  # указан сам memos.db
        return os.path.dirname(os.path.dirname(path))
    return path


def show_list(title, items, limit=30):
    """Списки строк журнала - по порядку строк, множества имён - по алфавиту."""
    if not items:
        return
    items = sorted(items) if isinstance(items, set) else list(items)
    print('\n%s (%d):' % (title, len(items)))
    for item in items[:limit]:
        print('  -', item)
    if len(items) > limit:
        print('  ... и ещё %d' % (len(items) - limit))


def main():
    args = [a for a in sys.argv[1:] if a != '--yes']
    if not args:
        sys.exit(__doc__)
    xlsx = os.path.abspath(args[0])
    if not os.path.isfile(xlsx):
        sys.exit('Нет файла %s' % xlsx)
    root = resolve_root(args[1] if len(args) > 1 else None)
    store = Store(root)
    try:
        store.open()
    except StoreError as e:
        sys.exit(str(e))
    print('База: %s' % store.db_path)

    sheet, rows = read_sheet(xlsx)
    header, cols = find_header(rows)
    print('Журнал: %s, лист «%s», заголовки в строке %d' % (xlsx, sheet, header + 1))
    existing = {(r['year'], r['number']) for r in store.all('SELECT year, number FROM memos')}
    res = parse_rows(rows, header, cols, Directory(store), existing)
    memos, errors = res['memos'], res['errors']

    print('\nБудет внесено записок: %d' % len(memos))
    for m in memos[:10]:
        print('  %s  %s  %s  | %s | %s' % (
            m['reg_number'], m['memo_date'], m['topic_text'][:30], m['executor_text'][:25],
            '; '.join(a['department'] for a in m['addressees'])[:40]))
    if len(memos) > 10:
        print('  ...')
    if res['already']:
        print('Уже есть в базе и будут пропущены: %d' % res['already'])
    show_list('Новые тематики - будут добавлены в справочник', res['new_topics'])
    show_list('Исполнители, которых нет в справочнике «Сотрудники» - будут записаны как в журнале',
              res['unknown_people'])
    show_list('Адресаты, которых нет в справочнике «Адресаты» - будут записаны как в журнале', res['unknown_to'])
    show_list('Предупреждения', res['warnings'])
    show_list('ОШИБКИ - эти строки не будут внесены', errors, limit=100)

    # Номер строки с ошибкой не попадёт в базу, и программа может выдать его повторно
    if errors:
        print('\nИсправьте ошибки в журнале и запустите скрипт снова, иначе номера этих строк\n'
              'программа может выдать повторно.')
    if not memos:
        print('\nВносить нечего.')
        return
    if '--yes' not in sys.argv:
        answer = input('\nВнести %d записок в базу? (да/нет): ' % len(memos)).strip().lower()
        if answer not in ('да', 'д', 'y', 'yes'):
            print('Ничего не внесено.')
            return

    backup = os.path.join(store.data_dir, 'memos-до-импорта-%s.db' % datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))
    shutil.copy2(store.db_path, backup)
    write_memos(store, memos)
    store.close()
    print('Внесено записок: %d. Копия базы до импорта: %s' % (len(memos), backup))


if __name__ == '__main__':
    main()
