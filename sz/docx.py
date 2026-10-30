"""Формирование Word-документа (.docx) служебной записки без внешних библиотек."""
import io
import json
import re
import zipfile

FONT = 'Arial'
FONT_SIZE = 14  # pt
CM = 567  # твипов в сантиметре


def twips(cm):
    """Сантиметры в твипы, половина округляется вверх, как в браузерной версии."""
    return int(cm * CM + 0.5)


TEXT_WIDTH = 11906 - twips(2.5) - twips(1)  # ширина текста между полями
MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
              'августа', 'сентября', 'октября', 'ноября', 'декабря']


def format_date(iso):
    y, m, d = iso[:10].split('-')
    return '%s.%s.%s' % (d, m, y)


def format_long_date(iso):
    """"14" января 2026 г."""
    y, m, d = iso[:10].split('-')
    return '"%s" %s %s г.' % (d, MONTHS_GEN[int(m) - 1], y)


def format_reg_number(prefix, number, year):
    """НТЦ-0001-26: префикс, порядковый номер с ведущими нулями, две цифры года."""
    return '%s-%04d-%s' % (prefix, int(number), str(year)[-2:])


def split_lines(text):
    return [line.strip() for line in re.split(r'\r?\n', str(text or '')) if line.strip()]


def short_name(full):
    """Иванов Иван Иванович -> Иванов И.И."""
    parts = str(full or '').split()
    if len(parts) < 2:
        return ' '.join(parts)
    if '.' in parts[1]:
        return '%s %s' % (parts[0], ''.join(parts[1:]))
    return '%s %s' % (parts[0], ''.join(p[0].upper() + '.' for p in parts[1:]))


def addressee_name(a):
    """Фамилия И.О. адресата в дательном падеже (заполняется в справочнике вручную).
    У адресатов, внесённых до появления этого поля, - краткая форма из полного ФИО."""
    return a.get('head_name_dative') or short_name(a.get('head_name'))


def greeting(addressees):
    """Обращение: по имени и отчеству одного адресата, «Уважаемые Коллеги!» для нескольких.
    Род определяется по окончанию отчества."""
    if len(addressees) > 1:
        return 'Уважаемые Коллеги!'
    if not addressees:
        return ''
    parts = str(addressees[0].get('head_name') or '').split()
    if len(parts) < 3 or any('.' in p for p in parts):
        return 'Уважаемый(ая) коллега!'
    patronymic = parts[2].lower()
    if re.search(r'(вич|ич|оглы|улы)$', patronymic):
        word = 'Уважаемый'
    elif re.search(r'(вна|чна|кызы|гызы)$', patronymic):
        word = 'Уважаемая'
    else:
        word = 'Уважаемый(ая)'
    return '%s %s %s!' % (word, parts[1], parts[2])


# ---------- Текст записки с оформлением ----------
#
# Хранится в memos.body_rich как JSON-список абзацев:
#   [{"align": "both", "indent": true,
#     "runs": [{"text": "...", "bold": true, "italic": true, "underline": true, "size": 12}]}]
# align: left | center | right | both (по ширине); indent - отступ первой строки;
# "\n" внутри text - перенос строки без нового абзаца. Ложные признаки не пишутся.
# Обычный текст записки (memos.body) хранится всегда - по нему идёт поиск.

RICH_SIZES = (10, 11, 12, 13, 14, 16)
ALIGNS = ('left', 'center', 'right', 'both')
RUN_FLAGS = ('bold', 'italic', 'underline')


def _clean_run(r):
    if not isinstance(r, dict) or not isinstance(r.get('text'), str):
        return None
    text = r['text'].replace('\r\n', '\n').replace('\r', '\n')
    if not text:
        return None
    run = {'text': text}
    for flag in RUN_FLAGS:
        if r.get(flag):
            run[flag] = True
    if r.get('size') in RICH_SIZES and r['size'] != FONT_SIZE:
        run['size'] = r['size']
    return run


def normalize_rich(paragraphs):
    """Оставляет только допустимое оформление, убирает пустые абзацы и пробелы по краям абзаца,
    склеивает соседние фрагменты с одинаковым оформлением."""
    out = []
    for p in paragraphs if isinstance(paragraphs, list) else []:
        if not isinstance(p, dict):
            continue
        runs = []
        for r in filter(None, map(_clean_run, p.get('runs') or [])):
            if runs and {k: v for k, v in runs[-1].items() if k != 'text'} == \
                    {k: v for k, v in r.items() if k != 'text'}:
                runs[-1]['text'] += r['text']
            else:
                runs.append(r)
        if runs:
            runs[0]['text'] = runs[0]['text'].lstrip()
            runs[-1]['text'] = runs[-1]['text'].rstrip()
            runs = [r for r in runs if r['text']]
        if runs:
            out.append({'align': p.get('align') if p.get('align') in ALIGNS else 'both',
                        'indent': bool(p.get('indent')), 'runs': runs})
    return out


def parse_rich(value):
    """Абзацы с оформлением из memos.body_rich или None, если оформления нет."""
    if not value:
        return None
    try:
        return normalize_rich(json.loads(value)) or None
    except ValueError:
        return None


def rich_plain_text(paragraphs):
    return '\n'.join(''.join(r['text'] for r in p['runs']) for p in paragraphs)


def rich_is_plain(paragraphs):
    """True, если оформление не отличается от обычного текста (по ширине, с отступом)."""
    return all(p['align'] == 'both' and p['indent'] and
               all(set(r) == {'text'} and '\n' not in r['text'] for r in p['runs'])
               for p in paragraphs)


def dump_rich(paragraphs):
    """JSON для memos.body_rich; None, если записке оформление не нужно."""
    paragraphs = normalize_rich(paragraphs)
    if not paragraphs or rich_is_plain(paragraphs):
        return None
    return json.dumps(paragraphs, ensure_ascii=False, separators=(',', ':'))


def memo_blocks(m, addressees, preview=False):
    """Содержимое записки как список блоков. По нему строятся и документ Word,
    и лист на экране, поэтому они не расходятся.

      {text | runs, bold, align, first_line, size} - абзац;
          runs - список {text, bold, italic, underline, size}
      {split: [слева, справа], size}               - строка с двумя краями
      {blank: True}                                - пустая строка

    В режиме preview незаполненные значения заменяются словарями {'ph': 'подсказка'}.
    """
    def v(value, hint):
        return value or ({'ph': hint} if preview else '')

    # Смысловые блоки по порядку; между ними ровно одна пустая строка
    sections = []

    def add(*blocks):
        if blocks:
            sections.append(list(blocks))

    org = split_lines(m.get('org_name'))
    if org:
        add(*[{'text': line, 'bold': True, 'align': 'center'} for line in org])
    elif preview:
        add({'text': {'ph': 'Наименование организации — укажите в настройках'}, 'align': 'center'})

    add({'text': 'СЛУЖЕБНАЯ ЗАПИСКА', 'bold': True, 'align': 'center'})

    if m.get('city'):
        add({'text': m['city'], 'align': 'center'})
    elif preview:
        add({'text': {'ph': 'Город — укажите в настройках'}, 'align': 'center'})

    date = format_long_date(m['memo_date']) if m.get('memo_date') else v('', 'дата')
    add({'split': [date, '№%s' % m.get('reg_number', '')]})

    # Адресаты: должность и Фамилия И.О. отдельными строками
    if addressees:
        blocks = []
        for a in addressees:
            blocks.append({'text': a['head_position'], 'align': 'right'})
            # У адресатов из перенесённого журнала, которых нет в справочнике, ФИО нет
            if addressee_name(a):
                blocks.append({'text': addressee_name(a), 'align': 'right'})
        add(*blocks)
    elif preview:
        add({'text': {'ph': 'Должность адресата'}, 'align': 'right'},
            {'text': {'ph': 'Фамилия И.О.'}, 'align': 'right'})

    add({'text': v(m.get('topic_name'), 'Тематика работ')})

    add({'text': greeting(addressees) or v('', 'Уважаемый Имя Отчество!'), 'align': 'center'})

    rich = parse_rich(m.get('body_rich'))
    body = split_lines(m.get('body'))
    if rich:
        # По левому краю - без явного выравнивания, как у остальных абзацев
        add(*[{'runs': p['runs'], 'align': None if p['align'] == 'left' else p['align'],
               'first_line': p['indent'], 'line_spacing': 360} for p in rich])
    elif body:
        add(*[{'text': line, 'align': 'both', 'first_line': True, 'line_spacing': 360} for line in body])
    elif preview:
        add({'text': {'ph': 'Текст служебной записки'}, 'align': 'both', 'first_line': True,
             'line_spacing': 360})

    attachment = split_lines(m.get('attachment'))
    if attachment:
        add(*[{'runs': [{'text': '' if i else 'Приложение: '}, {'text': line}], 'align': 'both'}
              for i, line in enumerate(attachment)])

    # Должность подписанта может быть в несколько строк; Фамилия И.О. - справа на последней
    if m.get('signer_name') or preview:
        position = split_lines(m.get('signer_position'))
        name = v(short_name(m.get('signer_name')), 'Фамилия И.О.')
        last = position[-1] if position else v('', 'Должность подписанта')
        add(*([{'text': line} for line in position[:-1]] + [{'split': [last, name]}]))

    executor_name = m.get('executor_name')
    executor = [{'text': v(executor_name and 'Исп. %s' % short_name(executor_name),
                           'Исп. Фамилия И.О.'), 'size': 10, 'footer': True}]
    if m.get('executor_phone'):
        executor.append({'text': 'Тел. %s' % m['executor_phone'], 'size': 10, 'footer': True})
    add(*executor)

    out = []
    for i, blocks in enumerate(sections):
        if i:
            out.append({'blank': True})
        out.extend(blocks)
    return out


_CONTROL = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def xml(text):
    text = (str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))
    return _CONTROL.sub('', text)  # управляющие символы недопустимы в XML


def run_xml(run, size=FONT_SIZE):
    """Размер задаётся у каждого фрагмента явно, чтобы не зависеть от настроек Word.
    Порядок элементов w:rPr задан схемой Word: b, i, sz, u."""
    size = run.get('size') or size
    rpr = (('<w:b/><w:bCs/>' if run.get('bold') else '') +
           ('<w:i/><w:iCs/>' if run.get('italic') else '') +
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size * 2, size * 2) +
           ('<w:u w:val="single"/>' if run.get('underline') else ''))
    text = run['text']
    if text == '\t':
        content = '<w:tab/>'
    else:
        content = '<w:br/>'.join('<w:t xml:space="preserve">%s</w:t>' % xml(part)
                                 for part in text.split('\n'))
    return '<w:r><w:rPr>%s</w:rPr>%s</w:r>' % (rpr, content)


def block_xml(blk):
    ppr = []
    if blk.get('blank'):
        runs = []
    elif blk.get('split'):
        # Правый край - через табуляцию по правому краю текста
        ppr.append('<w:tabs><w:tab w:val="right" w:pos="%d"/></w:tabs>' % TEXT_WIDTH)
        runs = [{'text': blk['split'][0]}, {'text': '\t'}, {'text': blk['split'][1]}]
    else:
        runs = blk.get('runs') or [{'text': blk.get('text'), 'bold': blk.get('bold')}]
    if blk.get('first_line'):
        ppr.append('<w:ind w:firstLine="%d"/>' % twips(1.25))
    if blk.get('line_spacing'):
        ppr.append('<w:spacing w:after="0" w:line="%d" w:lineRule="auto"/>' % blk['line_spacing'])
    if blk.get('align'):
        ppr.append('<w:jc w:val="%s"/>' % blk['align'])
    size = blk.get('size') or FONT_SIZE
    rpr_default = '<w:rPr><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (size * 2, size * 2)
    body = ''.join(run_xml(r, size) for r in runs if r['text'])
    return '<w:p><w:pPr>%s%s</w:pPr>%s</w:p>' % (''.join(ppr), rpr_default, body)


def document_xml(memo, addressees):
    blocks = memo_blocks(memo, addressees)
    while blocks and (blocks[-1].get('footer') or blocks[-1].get('blank')):
        blocks.pop()
    body = ''.join(block_xml(b) for b in blocks)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body>%s'
            '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
            'w:header="0" w:footer="%d" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>'
            % (body, twips(1.5), twips(1), twips(1), twips(2.5), twips(0.2)))


def footer_xml(memo):
    """Исполнитель у нижнего края только последней страницы.

    PAGE/NUMPAGES вычисляет Word после разбиения текста на страницы;
    поэтому подбирать пустые строки под объём записки не требуется.
    """
    def field(instruction):
        return ('<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
                '<w:r><w:instrText xml:space="preserve">%s</w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="end"/></w:r>') % xml(instruction)

    paragraphs = []
    for b in memo_blocks(memo, []):
        if not b.get('footer') or not b.get('text'):
            continue
        text = b['text']
        # Кавычки и обратные слеши экранируются также для языка полей Word.
        literal = text.replace('\\', '\\\\').replace('"', '\\"')
        paragraphs.append(
            '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
            '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>'
            + field(' PAGE ') +
            '<w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>'
            + field(' NUMPAGES ') +
            '<w:r><w:instrText xml:space="preserve">%s</w:instrText></w:r>' % xml(' "%s" "" ' % literal) +
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run_xml({'text': text}, 10) +
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + ''.join(paragraphs) + '</w:ftr>')


STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
    '<w:sz w:val="{s}"/><w:szCs w:val="{s}"/><w:lang w:val="ru-RU"/>'
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
    '</w:pPr></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
    '<w:sz w:val="{s}"/><w:szCs w:val="{s}"/></w:rPr></w:style>'
    '</w:styles>').format(f=FONT, s=FONT_SIZE * 2)

CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '</Types>')

ROOT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>')

DOCUMENT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    '</Relationships>')


def build_docx(memo, addressees, org_name=''):
    """Возвращает содержимое файла .docx (bytes)."""
    memo = dict(memo)
    # Записки прошлых версий не хранят шапку - берётся текущая из настроек
    if memo.get('org_name') is None:
        memo['org_name'] = org_name
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES_XML)
        z.writestr('_rels/.rels', ROOT_RELS_XML)
        z.writestr('word/document.xml', document_xml(memo, addressees))
        z.writestr('word/_rels/document.xml.rels', DOCUMENT_RELS_XML)
        z.writestr('word/styles.xml', STYLES_XML)
        z.writestr('word/footer1.xml', footer_xml(memo))
        z.writestr('word/settings.xml',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                   '<w:updateFields w:val="true"/></w:settings>')
    return buf.getvalue()


def safe_name(text, limit=60):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', ' ', str(text))
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'^[ .]+|[ .]+$', '', s)[:limit]
    s = re.sub(r'[ .]+$', '', s)
    return s or 'без темы'


def file_base_name(memo):
    """Имя папки и файла документа: номер СЗ, например НТЦ-0001-26."""
    return safe_name(memo['reg_number'])

