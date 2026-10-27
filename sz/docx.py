"""Формирование Word-документа (.docx) служебной записки без внешних библиотек."""
import io
import re
import zipfile

FONT = 'Calibri'
FONT_SIZE = 14  # pt
CM = 567  # твипов в сантиметре


def twips(cm):
    """Сантиметры в твипы, половина округляется вверх, как в браузерной версии."""
    return int(cm * CM + 0.5)


TEXT_WIDTH = 11906 - twips(3) - twips(1.5)  # ширина текста между полями
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


def memo_blocks(m, addressees, preview=False):
    """Содержимое записки как список блоков. По нему строятся и документ Word,
    и лист на экране, поэтому они не расходятся.

      {text | runs, bold, align, first_line, size} - абзац
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
            blocks.append({'text': addressee_name(a), 'align': 'right'})
        add(*blocks)
    elif preview:
        add({'text': {'ph': 'Должность адресата'}, 'align': 'right'},
            {'text': {'ph': 'Фамилия И.О.'}, 'align': 'right'})

    add({'text': v(m.get('topic_name'), 'Тематика работ')})

    add({'text': greeting(addressees) or v('', 'Уважаемый Имя Отчество!'), 'align': 'center'})

    body = split_lines(m.get('body'))
    if body:
        add(*[{'text': line, 'align': 'both', 'first_line': True} for line in body])
    elif preview:
        add({'text': {'ph': 'Текст служебной записки'}, 'align': 'both', 'first_line': True})

    attachment = split_lines(m.get('attachment'))
    if attachment:
        add(*[{'runs': [['' if i else 'Приложение: '], [line]], 'align': 'both'}
              for i, line in enumerate(attachment)])

    # Должность подписанта может быть в несколько строк; Фамилия И.О. - справа на последней
    if m.get('signer_name') or preview:
        position = split_lines(m.get('signer_position'))
        name = v(short_name(m.get('signer_name')), 'Фамилия И.О.')
        last = position[-1] if position else v('', 'Должность подписанта')
        add(*([{'text': line} for line in position[:-1]] + [{'split': [last, name]}]))

    executor_name = m.get('executor_name')
    executor = [{'text': v(executor_name and 'Исп. %s' % short_name(executor_name),
                           'Исп. Фамилия И.О.'), 'size': 10}]
    if m.get('executor_phone'):
        executor.append({'text': 'Тел. %s' % m['executor_phone'], 'size': 10})
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


def run_xml(text, bold, size=FONT_SIZE):
    """Размер задаётся у каждого фрагмента явно, чтобы не зависеть от настроек Word."""
    rpr = ('<w:b/><w:bCs/>' if bold else '') + \
        '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size * 2, size * 2)
    content = '<w:tab/>' if text == '\t' else '<w:t xml:space="preserve">%s</w:t>' % xml(text)
    return '<w:r><w:rPr>%s</w:rPr>%s</w:r>' % (rpr, content)


def block_xml(blk):
    ppr = []
    if blk.get('blank'):
        runs = []
    elif blk.get('split'):
        # Правый край - через табуляцию по правому краю текста
        ppr.append('<w:tabs><w:tab w:val="right" w:pos="%d"/></w:tabs>' % TEXT_WIDTH)
        runs = [[blk['split'][0]], ['\t'], [blk['split'][1]]]
    else:
        runs = blk.get('runs') or [[blk.get('text'), blk.get('bold')]]
    if blk.get('first_line'):
        ppr.append('<w:ind w:firstLine="%d"/>' % twips(1.25))
    if blk.get('align'):
        ppr.append('<w:jc w:val="%s"/>' % blk['align'])
    size = blk.get('size') or FONT_SIZE
    rpr_default = '<w:rPr><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (size * 2, size * 2)
    body = ''.join(run_xml(r[0], r[1] if len(r) > 1 else False, size) for r in runs if r[0])
    return '<w:p><w:pPr>%s%s</w:pPr>%s</w:p>' % (''.join(ppr), rpr_default, body)


def document_xml(memo, addressees):
    body = ''.join(block_xml(b) for b in memo_blocks(memo, addressees))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>%s'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
            'w:header="708" w:footer="708" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>'
            % (body, twips(2), twips(1.5), twips(2), twips(3)))


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
