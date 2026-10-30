"""Действия с данными: создание СЗ, справочники, настройки. Окна только вызывают их."""
import re

from . import docx
from .store import StoreError

# Описание справочников: заголовок, поля, сортировка.
# Поле: (имя, подпись, обязательное, многострочное - переносы сохраняются и печатаются)
DIRECTORIES = {
    'employees': {
        'title': 'Сотрудники',
        'fields': [('full_name', 'ФИО полностью', True, False),
                   ('position', 'Должность', False, True),
                   ('phone', 'Номер телефона', False, False)],
        'signer': True,  # колонка «Подписант»
        'order': 'full_name',
    },
    'addressees': {
        'title': 'Адресаты',
        'fields': [('department', 'Отдел', True, False),
                   ('head_position', 'Должность в документе, например «Начальнику отдела»', True, True),
                   ('head_name', 'ФИО руководителя полностью', True, False),
                   ('head_name_dative', 'Фамилия И.О. в дательном падеже, например «Петрову П.П.»',
                    True, False)],
        'order': 'department',
    },
    'topics': {
        'title': 'Тематики работ',
        'fields': [('name', 'Название', True, False)],
        'order': 'name',
    },
}


def next_number(store, year):
    return store.one('SELECT COALESCE(MAX(number), 0) + 1 AS n FROM memos WHERE year = ?', (year,))['n']


def create_memo(store, data, addressee_ids):
    """Сохраняет новую СЗ в реестре и формирует документ Word.

    data: topic_id, executor_id, memo_date (ГГГГ-ММ-ДД), body, attachment и, если текст
    оформлен, body_rich - список абзацев (см. docx.normalize_rich); тогда body не нужен.
    Возвращает (id записки, текст ошибки создания Word или None).
    """
    rich = docx.normalize_rich(data.get('body_rich')) if data.get('body_rich') is not None else None
    body = docx.rich_plain_text(rich) if rich is not None else (data.get('body') or '').strip()
    body_rich = docx.dump_rich(rich) if rich else None
    if not data.get('topic_id'):
        raise StoreError('Выберите тематику работ.')
    if not addressee_ids:
        raise StoreError('Выберите хотя бы одного адресата.')
    if not body:
        raise StoreError('Заполните текст служебной записки.')
    if not data.get('executor_id'):
        raise StoreError('Выберите исполнителя.')
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', data.get('memo_date') or ''):
        raise StoreError('Укажите дату.')
    year = int(data['memo_date'][:4])

    def insert(s):
        topic = s.one('SELECT * FROM topics WHERE id = ?', (data['topic_id'],))
        executor = s.one('SELECT * FROM employees WHERE id = ?', (data['executor_id'],))
        signer = s.signer
        addrs = [s.one('SELECT * FROM addressees WHERE id = ?', (i,)) for i in addressee_ids]
        if not signer:
            raise StoreError('Подписант не назначен. Отметьте его галочкой в справочнике «Сотрудники».')
        if not topic or not executor or not all(addrs):
            raise StoreError('Выбранную запись справочника кто-то удалил. Выберите заново.')
        no_dative = next((a for a in addrs if not a['head_name_dative']), None)
        if no_dative:
            raise StoreError('У адресата %s не заполнена Фамилия И.О. в дательном падеже. '
                             'Заполните её в справочнике «Адресаты».' % no_dative['head_name'])
        if not signer['position']:
            raise StoreError('У подписанта %s не указана должность. '
                             'Заполните её в справочнике «Сотрудники».' % signer['full_name'])
        # Номер считается под блокировкой, поэтому у двух пользователей он не совпадёт
        number = next_number(s, year)

        # Колонки addressee_* - сводка для поиска; сами адресаты - в memo_addressees
        def joined(key):
            return '; '.join(a[key] for a in addrs)

        memo_id = s.run(
            'INSERT INTO memos (number, year, reg_number, memo_date, topic_id, topic_name, '
            'addressee_id, addressee_department, addressee_position, addressee_name, '
            'body, attachment, signer_id, signer_name, signer_position, '
            'executor_id, executor_name, executor_phone, city, org_name, body_rich) '
            'VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
            (number, year, docx.format_reg_number(s.number_prefix, number, year), data['memo_date'],
             topic['id'], topic['name'], addrs[0]['id'], joined('department'), joined('head_position'),
             joined('head_name'), body, (data.get('attachment') or '').strip(), signer['id'],
             signer['full_name'], signer['position'], executor['id'], executor['full_name'],
             executor['phone'], s.city, s.org_name, body_rich))
        for i, a in enumerate(addrs):
            s.run('INSERT INTO memo_addressees (memo_id, ord, addressee_id, department, head_position, '
                  'head_name, head_name_dative) VALUES (?, ?, ?, ?, ?, ?, ?)',
                  (memo_id, i, a['id'], a['department'], a['head_position'], a['head_name'],
                   a['head_name_dative']))
        return memo_id

    memo_id = store.write(insert)
    try:
        regenerate_document(store, memo_id, new_file=True)
    except Exception as e:
        return memo_id, str(e)
    return memo_id, None


def regenerate_document(store, memo_id, new_file=False):
    """Формирует документ Word записки (заново - по прежнему пути) и запоминает путь."""
    memo = store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
    if not memo:
        raise StoreError('Записку удалили из реестра.')
    path = store.save_document(memo, None if new_file else memo['file_path'])
    if path != memo['file_path']:
        store.write(lambda s: s.run('UPDATE memos SET file_path = ? WHERE id = ?', (path, memo_id)))
    return path


def delete_memo(store, memo_id):
    def delete(s):
        s.run('DELETE FROM memo_addressees WHERE memo_id = ?', (memo_id,))
        s.run('DELETE FROM memos WHERE id = ?', (memo_id,))
    store.write(delete)


def directory_rows(store, kind):
    return store.all('SELECT * FROM %s ORDER BY %s' % (kind, DIRECTORIES[kind]['order']))


def clean_values(kind, raw):
    """Проверяет и очищает значения полей справочника."""
    values = {}
    for name, label, required, multiline in DIRECTORIES[kind]['fields']:
        text = raw.get(name) or ''
        values[name] = '\n'.join(docx.split_lines(text)) if multiline else text.strip()
        if required and not values[name]:
            raise StoreError('Заполните поле «%s».' % label)
    return values


def save_directory_row(store, kind, raw, row_id=None):
    values = clean_values(kind, raw)
    cols = list(values)

    def save(s):
        if row_id is not None:
            s.run('UPDATE %s SET %s WHERE id = ?' % (kind, ', '.join('%s = ?' % c for c in cols)),
                  [values[c] for c in cols] + [row_id])
        else:
            s.run('INSERT INTO %s (%s) VALUES (%s)' % (kind, ', '.join(cols), ', '.join('?' * len(cols))),
                  [values[c] for c in cols])
    try:
        store.write(save)
    except Exception as e:
        if 'UNIQUE' in str(e):
            raise StoreError('Такая запись уже есть в справочнике.')
        raise


def delete_directory_row(store, kind, row_id):
    store.write(lambda s: s.run('DELETE FROM %s WHERE id = ?' % kind, (row_id,)))


def set_signer(store, employee_id, on):
    """Подписант один: пока он назначен, другого назначить нельзя."""
    def update(s):
        if on:
            current = s.signer
            if current and current['id'] != employee_id:
                raise StoreError('Подписант уже назначен: %s. Сначала снимите галочку у него.'
                                 % current['full_name'])
            emp = s.one('SELECT * FROM employees WHERE id = ?', (employee_id,))
            if not emp:
                raise StoreError('Сотрудник удалён.')
            if not emp['position']:
                raise StoreError('У сотрудника %s не указана должность: она печатается рядом с подписью. '
                                 'Сначала заполните её.' % emp['full_name'])
        s.run('UPDATE employees SET is_signer = ? WHERE id = ?', (1 if on else 0, employee_id))
    store.write(update)


BAD_CHARS = re.compile(r'[\\/:*?"<>|]')


def save_settings(store, org_name, city, prefix, output_dir):
    org_name = '\n'.join(docx.split_lines(org_name))
    prefix = prefix.strip()
    output_dir = output_dir.strip().replace('\\', '/')
    if not prefix or BAD_CHARS.search(prefix):
        raise StoreError('Префикс не может быть пустым и не может содержать символы \\ / : * ? " < > |')
    if not output_dir or any(p == '..' or re.search(r'[:*?"<>|]', p) for p in output_dir.split('/')):
        raise StoreError('Укажите папку внутри общей папки, без «..» и символов : * ? " < > |')

    def save(s):
        s.set_setting('org_name', org_name)
        s.set_setting('city', city.strip())
        s.set_setting('number_prefix', prefix)
        s.set_setting('output_dir', output_dir)
    store.write(save)


def search_memos(store, q='', topic_id=None, year=None):
    sql = 'SELECT * FROM memos WHERE 1=1'
    params = []
    if topic_id:
        sql += ' AND topic_id = ?'
        params.append(int(topic_id))
    if year:
        sql += ' AND year = ?'
        params.append(int(year))
    memos = store.all(sql + ' ORDER BY year DESC, number DESC', params)
    if q:
        # LIKE в SQLite не учитывает регистр кириллицы, поэтому ищем в Python
        needle = q.lower()
        keys = ('reg_number', 'topic_name', 'addressee_department', 'addressee_name',
                'executor_name', 'signer_name', 'body', 'attachment')
        memos = [m for m in memos if needle in ' '.join(str(m.get(k) or '') for k in keys).lower()]
    return memos


# Адресатов у записки может быть несколько, поэтому они считаются по memo_addressees
GROUP_QUERIES = {
    'topic': 'SELECT topic_name AS label, COUNT(*) AS n FROM memos WHERE year = ? GROUP BY label',
    'executor': 'SELECT executor_name AS label, COUNT(*) AS n FROM memos WHERE year = ? GROUP BY label',
    'addressee': 'SELECT a.department AS label, COUNT(*) AS n FROM memo_addressees a '
                 'JOIN memos m ON m.id = a.memo_id WHERE m.year = ? GROUP BY label',
}


def group_counts(store, kind, year, limit=7):
    rows = store.all(GROUP_QUERIES[kind] + ' ORDER BY n DESC, label', (year,))
    if len(rows) <= limit + 1:
        return rows
    rest = sum(r['n'] for r in rows[limit:])
    return rows[:limit] + [{'label': 'Остальные (%d)' % (len(rows) - limit), 'n': rest}]


def month_counts(store, year):
    counts = [0] * 12
    for r in store.all('SELECT CAST(substr(memo_date, 6, 2) AS INTEGER) AS m, COUNT(*) AS n '
                       'FROM memos WHERE year = ? GROUP BY m', (year,)):
        counts[r['m'] - 1] = r['n']
    return counts


