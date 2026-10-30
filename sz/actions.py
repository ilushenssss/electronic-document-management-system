"""Действия с данными: создание СЗ, справочники, настройки. Окна только вызывают их."""
import re
import datetime
import hashlib
import json

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


class MemoConflict(StoreError):
    """Записка изменена или удалена после открытия формы."""


def load_memo_for_edit(store, memo_id):
    # Единый снимок записи и её адресатов, без изменения схемы базы.
    store.db.execute('SAVEPOINT edit_snapshot')
    try:
        memo = store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        if not memo:
            raise MemoConflict('СЗ удалена из реестра. Введённые данные сохранены в форме.')
        addressees = store.memo_addressees(memo_id)
        payload = json.dumps([memo, addressees], ensure_ascii=False, sort_keys=True)
        version = hashlib.sha256(payload.encode('utf-8')).hexdigest()
        return {'memo': memo, 'addressees': addressees, 'version': version}
    finally:
        store.db.execute('RELEASE SAVEPOINT edit_snapshot')


def update_memo(store, snapshot, expected_version, data, addressee_keys):
    """Обновить существующую СЗ и Word; отрицательные ключи — сохранённые значения.

    Справочники могут меняться независимо: сохранённые значения не подменяются
    текущими. Ошибка Word возвращается отдельно после успешного COMMIT базы.
    """
    memo_id = snapshot['memo']['id']
    rich = docx.normalize_rich(data.get('body_rich')) if data.get('body_rich') is not None else None
    body = docx.rich_plain_text(rich) if rich is not None else (data.get('body') or '').strip()
    if not body:
        raise StoreError('Заполните текст служебной записки.')
    if not addressee_keys:
        raise StoreError('Выберите хотя бы одного адресата.')
    try:
        date = datetime.date.fromisoformat(data.get('memo_date') or '')
    except ValueError:
        raise StoreError('Укажите корректную дату.')
    if date.year != snapshot['memo']['year']:
        raise StoreError('Дату можно изменить только в пределах года номера СЗ.')

    def update(s):
        current = load_memo_for_edit(s, memo_id)
        if current['version'] != expected_version:
            raise MemoConflict('СЗ уже изменена другим пользователем. Перечитайте её перед сохранением. '
                               'Ваш текст остаётся в форме.')
        old = snapshot['memo']
        topic_id = data.get('topic_id')
        executor_id = data.get('executor_id')
        if topic_id == -1:
            topic = {'id': old['topic_id'], 'name': old['topic_name']}
        else:
            topic = s.one('SELECT * FROM topics WHERE id = ?', (topic_id,))
        if executor_id == -1:
            executor = {'id': old['executor_id'], 'full_name': old['executor_name'], 'phone': old['executor_phone']}
        else:
            executor = s.one('SELECT * FROM employees WHERE id = ?', (executor_id,))
        if not topic or not executor:
            raise StoreError('Выберите тематику и исполнителя. Выбранная запись могла быть удалена.')
        addrs = []
        for key in addressee_keys:
            if key < 0:
                index = -key - 1
                if index >= len(snapshot['addressees']):
                    raise StoreError('Перечитайте список адресатов СЗ.')
                addr = snapshot['addressees'][index]
                addr = dict(addr, id=addr['addressee_id'])
            else:
                addr = s.one('SELECT * FROM addressees WHERE id = ?', (key,))
                if addr and not addr['head_name_dative']:
                    raise StoreError('У выбранного адресата не заполнена Фамилия И.О. в дательном падеже.')
            if not addr:
                raise StoreError('Выбранного адресата удалили. Выберите заново.')
            addrs.append(addr)
        values = {'memo_date': date.isoformat(), 'body': body,
                  'body_rich': docx.dump_rich(rich) if rich else None,
                  'attachment': (data.get('attachment') or '').strip(),
                  'topic_id': topic['id'], 'topic_name': topic['name'],
                  'executor_id': executor['id'], 'executor_name': executor['full_name'],
                  'executor_phone': executor['phone'], 'addressee_id': addrs[0]['id']}
        for field, key in [('addressee_department', 'department'), ('addressee_position', 'head_position'),
                           ('addressee_name', 'head_name')]:
            values[field] = '; '.join(a[key] for a in addrs)
        s.run('UPDATE memos SET %s WHERE id = ?' % ', '.join(key + ' = ?' for key in values),
              list(values.values()) + [memo_id])
        s.run('DELETE FROM memo_addressees WHERE memo_id = ?', (memo_id,))
        for order, addr in enumerate(addrs):
            s.run('INSERT INTO memo_addressees (memo_id, ord, addressee_id, department, head_position, '
                  'head_name, head_name_dative) VALUES (?, ?, ?, ?, ?, ?, ?)',
                  (memo_id, order, addr['id'], addr['department'], addr['head_position'],
                   addr['head_name'], addr['head_name_dative']))
        return {'document_error': None}

    def publish(result):
        memo = store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        try:
            path = store.save_document(memo, memo['file_path'])
            if path != memo['file_path']:
                store.run('UPDATE memos SET file_path = ? WHERE id = ?', (path, memo_id))
        except Exception as e:
            result['document_error'] = str(e)
        result['version'] = load_memo_for_edit(store, memo_id)['version']

    return store.write(update, after_commit=publish)


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


def _memo_filter(q='', topic_id=None, year=None, executor_name=None, addressee_department=None):
    sql = ' WHERE 1=1'
    params = []
    if topic_id:
        sql += ' AND topic_id = ?'
        params.append(int(topic_id))
    if year:
        sql += ' AND year = ?'
        params.append(int(year))
    if executor_name:
        sql += ' AND executor_name = ?'
        params.append(executor_name)
    if addressee_department:
        sql += (' AND EXISTS (SELECT 1 FROM memo_addressees a '
                'WHERE a.memo_id = memos.id AND a.department = ?)')
        params.append(addressee_department)
    if q:
        keys = ('reg_number', 'topic_name', 'addressee_department', 'addressee_name',
                'executor_name', 'signer_name', 'body', 'attachment')
        text = " || ' ' || ".join("COALESCE(%s, '')" % key for key in keys)
        sql += ' AND instr(unicode_lower(%s), ?) > 0' % text
        params.append(q.lower())
    return sql, params


def search_memos(store, q='', topic_id=None, year=None, executor_name=None, addressee_department=None):
    where, params = _memo_filter(q, topic_id, year, executor_name, addressee_department)
    return store.all('SELECT * FROM memos' + where + ' ORDER BY year DESC, number DESC', params)


def search_memos_page(store, q='', topic_id=None, year=None, executor_name=None,
                      addressee_department=None, page=0, page_size=50):
    """Одна страница реестра; фильтрация и подсчёт выполняются до LIMIT.

    Номер страницы начинается с нуля и ограничивается существующими страницами.
    """
    page_size = max(1, int(page_size))
    where, params = _memo_filter(q, topic_id, year, executor_name, addressee_department)
    # Подсчёт и выборка видят один снимок при одновременной работе пользователей.
    store.db.execute('SAVEPOINT registry_page')
    try:
        total = store.one('SELECT COUNT(*) AS n FROM memos')['n']
        matched = store.one('SELECT COUNT(*) AS n FROM memos' + where, params)['n']
        page_count = max(1, (matched + page_size - 1) // page_size)
        page = min(max(0, int(page)), page_count - 1)
        memos = store.all('SELECT * FROM memos' + where +
                          ' ORDER BY year DESC, number DESC LIMIT ? OFFSET ?',
                          params + [page_size, page * page_size])
        return {'memos': memos, 'total': total, 'matched': matched,
                'page': page, 'page_count': page_count}
    finally:
        store.db.execute('RELEASE SAVEPOINT registry_page')


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


