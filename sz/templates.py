"""Шаблоны и подстановка полей без выполнения кода и изменения исходного шаблона."""
import re

from . import docx
from .store import StoreError

KINDS = {'memo': 'Полная СЗ', 'fragment': 'Формулировка'}
FIELD = re.compile(r'(?<!\{)\{([\w][\w \-]{0,79})\}(?!\})', re.UNICODE)


def paragraphs(row):
    return docx.parse_rich(row.get('body_rich')) or [
        {'indent': True, 'runs': [{'text': line}]} for line in row.get('body', '').splitlines()]


def fields(content):
    found = {}
    for paragraph in content:
        text = ''.join(run['text'] for run in paragraph['runs'])
        for match in FIELD.finditer(text):
            name = match.group(1).strip()
            found.setdefault(name.casefold(), name)
    return found


def render(content, values):
    """Одна подстановка; оформление замены берётся у открывающей скобки."""
    values = {key.strip().casefold(): str(value).strip() for key, value in values.items()}
    for key, name in fields(content).items():
        if not values.get(key):
            raise StoreError('Заполните поле «%s».' % name)
    result = []
    for paragraph in content:
        runs = paragraph['runs']
        text = ''.join(run['text'] for run in runs)
        out = []

        def copy_slice(start, end):
            offset = 0
            for run in runs:
                left, right = max(start, offset), min(end, offset + len(run['text']))
                if left < right:
                    out.append(dict(run, text=run['text'][left-offset:right-offset]))
                offset += len(run['text'])

        start = 0
        for match in FIELD.finditer(text):
            copy_slice(start, match.start())
            offset = 0
            for run in runs:
                if offset + len(run['text']) > match.start():
                    out.append(dict(run, text=values[match.group(1).strip().casefold()]))
                    break
                offset += len(run['text'])
            start = match.end()
        copy_slice(start, len(text))
        result.append(dict(paragraph, runs=out))
    return docx.normalize_rich(result)


def list_templates(store, active_only=False, topic_id=None):
    sql = 'SELECT t.*, p.name AS topic_name FROM templates t LEFT JOIN topics p ON p.id=t.topic_id'
    where, args = [], []
    if active_only:
        where.append('t.active=1')
    if topic_id is not None:
        where.append('(t.topic_id IS NULL OR t.topic_id=? OR (t.active=1 AND t.for_configuration_control=1))')
        args.append(topic_id)
    if where:
        sql += ' WHERE ' + ' AND '.join(where)
    return store.all(sql + ' ORDER BY unicode_lower(t.name), t.id', args)


def configuration_control_template(store):
    return store.one('SELECT * FROM templates WHERE active=1 AND for_configuration_control=1')


def get_template(store, row_id):
    row = store.one('SELECT * FROM templates WHERE id=?', (row_id,))
    if not row:
        raise StoreError('Шаблон удалён другим пользователем. Обновите список.')
    return row


def save_template(store, raw, row_id=None, expected_version=None):
    name = (raw.get('name') or '').strip()
    kind = raw.get('kind')
    content = docx.normalize_rich(raw.get('paragraphs') or [])
    body = docx.rich_plain_text(content)
    if not name:
        raise StoreError('Укажите название шаблона.')
    if kind not in KINDS:
        raise StoreError('Выберите вид шаблона.')
    if not body.strip():
        raise StoreError('Введите текст шаблона.')
    topic_id = raw.get('topic_id') or None
    active = bool(raw.get('active', True))
    linked = active and bool(raw.get('for_configuration_control'))
    values = [name, kind, body, docx.dump_rich(content), topic_id, int(active), int(linked)]

    def save(s):
        if row_id is not None:
            current = get_template(s, row_id)
            if current['version'] != expected_version:
                raise StoreError('Шаблон изменён другим пользователем. Ваш текст сохранён в форме. '
                                 'Перечитайте шаблон перед сохранением.')
        if any(r['id'] != row_id and r['name'].casefold() == name.casefold()
               for r in s.all('SELECT id, name FROM templates')):
            raise StoreError('Шаблон с таким названием уже существует. Измените название.')
        if topic_id is not None and not s.one('SELECT id FROM topics WHERE id=?', (topic_id,)):
            raise StoreError('Тематика удалена. Выберите другую тематику или «Любая тематика».')
        if linked:
            s.run('UPDATE templates SET for_configuration_control=0, version=version+1 '
                  'WHERE for_configuration_control=1 AND id<>?', (row_id or 0,))
        if row_id is None:
            s.run('INSERT INTO templates(name,kind,body,body_rich,topic_id,active,for_configuration_control) '
                  'VALUES(?,?,?,?,?,?,?)', values)
            return s.one('SELECT last_insert_rowid() AS id')['id']
        s.run('UPDATE templates SET name=?,kind=?,body=?,body_rich=?,topic_id=?,active=?, '
              'for_configuration_control=?,version=version+1 WHERE id=?', values + [row_id])
        return row_id
    return store.write(save)


def delete_template(store, row_id, expected_version):
    def delete(s):
        if get_template(s, row_id)['version'] != expected_version:
            raise StoreError('Шаблон изменён другим пользователем. Перечитайте его перед удалением.')
        s.run('DELETE FROM templates WHERE id=?', (row_id,))
    store.write(delete)


EXAMPLES = [
    ('Согласование закупки', 'memo',
     'Прошу согласовать закупку {наименование} в количестве {количество} для {цель}.'),
    ('Направление документов', 'memo',
     'Направляю {документ} для {цель}.\nДата: {дата}. Исполнитель: {исполнитель}.'),
    ('Производственная необходимость', 'fragment',
     'В связи с производственной необходимостью прошу {действие}.'),
]


def add_examples(store):
    count = 0
    for name, kind, text in EXAMPLES:
        if any(row['name'].casefold() == name.casefold() for row in list_templates(store)):
            continue
        save_template(store, {'name': name, 'kind': kind,
                             'paragraphs': [{'indent': True, 'runs': [{'text': line}]} for line in text.splitlines()]})
        count += 1
    return count
