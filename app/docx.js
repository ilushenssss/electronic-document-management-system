// Формирование Word-документа (.docx) служебной записки без внешних библиотек.
(function (root) {
  'use strict';

  const FONT = 'Arial';
  const FONT_SIZE = 28; // в полупунктах: 14 pt
  const CM = 567; // твипов в сантиметре

  const TEXT_WIDTH = 11906 - 3 * CM - Math.round(1.5 * CM); // ширина текста между полями
  const MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля',
    'августа', 'сентября', 'октября', 'ноября', 'декабря'];

  function formatDate(iso) {
    const [y, m, d] = iso.split('-');
    return `${d}.${m}.${y}`;
  }

  // "14" января 2026 г.
  function formatLongDate(iso) {
    const [y, m, d] = iso.split('-');
    return `"${d}" ${MONTHS_GEN[Number(m) - 1]} ${y} г.`;
  }

  // НТЦ-0001-26: префикс, порядковый номер с ведущими нулями, две цифры года
  function formatRegNumber(prefix, number, year) {
    return `${prefix}-${String(number).padStart(4, '0')}-${String(year).slice(-2)}`;
  }

  function splitLines(text) {
    return String(text || '').split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  }

  // Иванов Иван Иванович -> Иванов И.И.
  function shortName(full) {
    const parts = String(full || '').trim().split(/\s+/).filter(Boolean);
    if (parts.length < 2) return parts.join(' ');
    if (parts[1].includes('.')) return `${parts[0]} ${parts.slice(1).join('')}`;
    return `${parts[0]} ${parts.slice(1).map((p) => `${p[0].toUpperCase()}.`).join('')}`;
  }

  // Фамилия И.О. адресата в дательном падеже (заполняется в справочнике вручную).
  // У адресатов, внесённых до появления этого поля, - краткая форма из полного ФИО.
  function addresseeName(a) {
    return a.head_name_dative || shortName(a.head_name);
  }

  // Обращение: по имени и отчеству одного адресата, «Уважаемые Коллеги!» для нескольких.
  // Род определяется по окончанию отчества.
  function greeting(addressees) {
    if (addressees.length > 1) return 'Уважаемые Коллеги!';
    if (!addressees.length) return '';
    const parts = String(addressees[0].head_name || '').trim().split(/\s+/);
    if (parts.length < 3 || parts.some((p) => p.includes('.'))) return 'Уважаемый(ая) коллега!';
    const patronymic = parts[2].toLowerCase();
    const word = /(вич|ич|оглы|улы)$/.test(patronymic) ? 'Уважаемый'
      : /(вна|чна|кызы|гызы)$/.test(patronymic) ? 'Уважаемая' : 'Уважаемый(ая)';
    return `${word} ${parts[1]} ${parts[2]}!`;
  }

  // ---------- Текст записки с оформлением ----------
  //
  // Хранится в memos.body_rich как JSON-список абзацев:
  //   [{"align": "both", "indent": true,
  //     "runs": [{"text": "...", "bold": true, "italic": true, "underline": true, "size": 12}]}]
  // align: left | center | right | both (по ширине); indent - отступ первой строки;
  // "\n" внутри text - перенос строки без нового абзаца. Ложные признаки не пишутся.
  // Обычный текст записки (memos.body) хранится всегда - по нему идёт поиск.
  // Тот же формат у настольной версии (sz/docx.py).

  const RICH_SIZES = [10, 11, 12, 13, 14, 16];
  const ALIGNS = ['left', 'center', 'right', 'both'];
  const RUN_FLAGS = ['bold', 'italic', 'underline'];

  function cleanRun(r) {
    if (!r || typeof r !== 'object' || typeof r.text !== 'string') return null;
    const text = r.text.replace(/\r\n?/g, '\n');
    if (!text) return null;
    const run = { text };
    for (const flag of RUN_FLAGS) if (r[flag]) run[flag] = true;
    if (RICH_SIZES.includes(r.size) && r.size !== FONT_SIZE / 2) run.size = r.size;
    return run;
  }

  function sameFormat(a, b) {
    return RUN_FLAGS.every((f) => !!a[f] === !!b[f]) && (a.size || 0) === (b.size || 0);
  }

  // Оставляет только допустимое оформление, убирает пустые абзацы и пробелы по краям абзаца,
  // склеивает соседние фрагменты с одинаковым оформлением
  function normalizeRich(paragraphs) {
    const out = [];
    for (const p of Array.isArray(paragraphs) ? paragraphs : []) {
      if (!p || typeof p !== 'object') continue;
      let runs = [];
      for (const r of (Array.isArray(p.runs) ? p.runs : []).map(cleanRun).filter(Boolean)) {
        const last = runs[runs.length - 1];
        if (last && sameFormat(last, r)) last.text += r.text;
        else runs.push(r);
      }
      if (runs.length) {
        runs[0].text = runs[0].text.replace(/^\s+/, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
        runs = runs.filter((r) => r.text);
      }
      if (runs.length) out.push({ align: ALIGNS.includes(p.align) ? p.align : 'both', indent: !!p.indent, runs });
    }
    return out;
  }

  // Абзацы с оформлением из memos.body_rich или null, если оформления нет
  function parseRich(value) {
    if (!value) return null;
    try {
      const paragraphs = normalizeRich(JSON.parse(value));
      return paragraphs.length ? paragraphs : null;
    } catch {
      return null;
    }
  }

  function richPlainText(paragraphs) {
    return paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
  }

  // Оформление не отличается от обычного текста (по ширине, с отступом)
  function richIsPlain(paragraphs) {
    return paragraphs.every((p) => p.align === 'both' && p.indent &&
      p.runs.every((r) => Object.keys(r).length === 1 && !r.text.includes('\n')));
  }

  // JSON для memos.body_rich; null, если записке оформление не нужно
  function dumpRich(paragraphs) {
    const clean = normalizeRich(paragraphs);
    if (!clean.length || richIsPlain(clean)) return null;
    return JSON.stringify(clean);
  }

  // Содержимое записки как список блоков. По нему строятся и документ Word,
  // и лист на экране, поэтому они не расходятся.
  //   { text | runs, bold, align, firstLine, size } - абзац;
  //       runs - список { text, bold, italic, underline, size }
  //   { split: [слева, справа], size }                 - строка с двумя краями
  //   { blank: true }                                  - пустая строка
  // В режиме preview незаполненные значения заменяются объектами { ph: 'подсказка' }.
  function memoBlocks(m, addressees, { preview = false } = {}) {
    const v = (value, hint) => value || (preview ? { ph: hint } : '');
    // Смысловые блоки по порядку; между ними ровно одна пустая строка
    const sections = [];
    const add = (...blocks) => { if (blocks.length) sections.push(blocks); };

    const org = splitLines(m.org_name);
    if (org.length) add(...org.map((line) => ({ text: line, bold: true, align: 'center' })));
    else if (preview) add({ text: { ph: 'Наименование организации — укажите в настройках' }, align: 'center' });

    add({ text: 'СЛУЖЕБНАЯ ЗАПИСКА', bold: true, align: 'center' });

    if (m.city) add({ text: m.city, align: 'center' });
    else if (preview) add({ text: { ph: 'Город — укажите в настройках' }, align: 'center' });

    add({ split: [m.memo_date ? formatLongDate(m.memo_date) : v('', 'дата'), `№${m.reg_number}`] });

    // Адресаты: должность и Фамилия И.О. отдельными строками
    if (addressees.length) {
      // У адресатов из перенесённого журнала, которых нет в справочнике, ФИО нет
      add(...addressees.flatMap((a) => [
        { text: a.head_position, align: 'right' },
        ...(addresseeName(a) ? [{ text: addresseeName(a), align: 'right' }] : [])]));
    } else if (preview) {
      add({ text: { ph: 'Должность адресата' }, align: 'right' }, { text: { ph: 'Фамилия И.О.' }, align: 'right' });
    }

    add({ text: v(m.topic_name, 'Тематика работ') });

    add({ text: greeting(addressees) || v('', 'Уважаемый Имя Отчество!'), align: 'center' });

    const rich = parseRich(m.body_rich);
    const body = splitLines(m.body);
    if (rich) {
      // По левому краю - без явного выравнивания, как у остальных абзацев
      add(...rich.map((p) => ({ runs: p.runs, align: p.align === 'left' ? undefined : p.align, firstLine: p.indent })));
    } else if (body.length) add(...body.map((line) => ({ text: line, align: 'both', firstLine: true })));
    else if (preview) add({ text: { ph: 'Текст служебной записки' }, align: 'both', firstLine: true });

    const attachment = splitLines(m.attachment);
    if (attachment.length) {
      add(...attachment.map((line, i) => ({ runs: [{ text: i ? '' : 'Приложение: ' }, { text: line }], align: 'both' })));
    }

    // Должность подписанта может быть в несколько строк; Фамилия И.О. - справа на последней
    if (m.signer_name || preview) {
      const position = splitLines(m.signer_position);
      const name = v(shortName(m.signer_name), 'Фамилия И.О.');
      add(...position.slice(0, -1).map((line) => ({ text: line })),
        { split: [position.length ? position[position.length - 1] : v('', 'Должность подписанта'), name] });
    }

    const executor = [{ text: v(m.executor_name && `Исп. ${shortName(m.executor_name)}`, 'Исп. Фамилия И.О.'), size: 10 }];
    if (m.executor_phone) executor.push({ text: `Тел. ${m.executor_phone}`, size: 10 });
    add(...executor);

    return sections.flatMap((blocks, i) => (i ? [{ blank: true }, ...blocks] : blocks));
  }


  function xml(text) {
    return String(text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      // Управляющие символы недопустимы в XML
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  }

  // Размер задаётся у каждого фрагмента явно, чтобы не зависеть от настроек Word.
  // Порядок элементов w:rPr задан схемой Word: b, i, sz, u.
  function run(r, blockSize = FONT_SIZE / 2) {
    const size = r.size || blockSize;
    const rPr = (r.bold ? '<w:b/><w:bCs/>' : '') + (r.italic ? '<w:i/><w:iCs/>' : '') +
      `<w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/>` +
      (r.underline ? '<w:u w:val="single"/>' : '');
    const content = r.text === '\t' ? '<w:tab/>'
      : r.text.split('\n').map((part) => `<w:t xml:space="preserve">${xml(part)}</w:t>`).join('<w:br/>');
    return `<w:r><w:rPr>${rPr}</w:rPr>${content}</w:r>`;
  }

  function blockXml(blk) {
    const pPr = [];
    let runs;
    if (blk.blank) {
      runs = [];
    } else if (blk.split) {
      // Правый край - через табуляцию по правому краю текста
      pPr.push(`<w:tabs><w:tab w:val="right" w:pos="${TEXT_WIDTH}"/></w:tabs>`);
      runs = [{ text: blk.split[0] }, { text: '\t' }, { text: blk.split[1] }];
    } else {
      runs = blk.runs || [{ text: blk.text, bold: blk.bold }];
    }
    if (blk.firstLine) pPr.push(`<w:ind w:firstLine="${Math.round(1.25 * CM)}"/>`);
    if (blk.align) pPr.push(`<w:jc w:val="${blk.align}"/>`);
    const size = blk.size || FONT_SIZE / 2;
    const rPrDefault = `<w:rPr><w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/></w:rPr>`;
    return `<w:p>${pPr.length || rPrDefault ? `<w:pPr>${pPr.join('')}${rPrDefault}</w:pPr>` : ''}` +
      runs.filter((r) => r.text).map((r) => run(r, size)).join('') + '</w:p>';
  }

  function documentXml(memo, addressees) {
    const body = memoBlocks(memo, addressees).map(blockXml).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:body>${body}` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      `<w:pgMar w:top="${2 * CM}" w:right="${Math.round(1.5 * CM)}" ` +
      `w:bottom="${2 * CM}" w:left="${3 * CM}" w:header="708" w:footer="708" w:gutter="0"/>` +
      '</w:sectPr></w:body></w:document>';
  }

  const STYLES_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    `<w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:eastAsia="${FONT}" w:cs="${FONT}"/>` +
    `<w:sz w:val="${FONT_SIZE}"/><w:szCs w:val="${FONT_SIZE}"/><w:lang w:val="ru-RU"/>` +
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' +
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>' +
    '</w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
    `<w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:eastAsia="${FONT}" w:cs="${FONT}"/>` +
    `<w:sz w:val="${FONT_SIZE}"/><w:szCs w:val="${FONT_SIZE}"/></w:rPr></w:style>` +
    '</w:styles>';

  const CONTENT_TYPES_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>';

  const ROOT_RELS_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  const DOCUMENT_RELS_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  // ---------- ZIP без сжатия (метод STORE) ----------

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(data) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function zip(files) {
    const enc = new TextEncoder();
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

    const chunks = [];
    const central = [];
    let offset = 0;

    for (const [name, content] of files) {
      const nameBytes = enc.encode(name);
      const data = typeof content === 'string' ? enc.encode(content) : content;
      const crc = crc32(data);

      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0x0800, true); // имена в UTF-8
      local.setUint16(8, 0, true);
      local.setUint16(10, dosTime, true);
      local.setUint16(12, dosDate, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, data.length, true);
      local.setUint32(22, data.length, true);
      local.setUint16(26, nameBytes.length, true);
      local.setUint16(28, 0, true);
      chunks.push(new Uint8Array(local.buffer), nameBytes, data);

      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true);
      cd.setUint16(4, 20, true);
      cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x0800, true);
      cd.setUint16(10, 0, true);
      cd.setUint16(12, dosTime, true);
      cd.setUint16(14, dosDate, true);
      cd.setUint32(16, crc, true);
      cd.setUint32(20, data.length, true);
      cd.setUint32(24, data.length, true);
      cd.setUint16(28, nameBytes.length, true);
      cd.setUint32(42, offset, true);
      central.push(new Uint8Array(cd.buffer), nameBytes);

      offset += 30 + nameBytes.length + data.length;
    }

    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);

    const parts = [...chunks, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
    let pos = 0;
    for (const p of parts) { out.set(p, pos); pos += p.length; }
    return out;
  }

  function buildDocx(memo, opts = {}) {
    return zip([
      ['[Content_Types].xml', CONTENT_TYPES_XML],
      ['_rels/.rels', ROOT_RELS_XML],
      // Записки прошлых версий не хранят шапку - берётся текущая из настроек
      ['word/document.xml', documentXml({ ...memo, org_name: memo.org_name ?? opts.orgName },
        opts.addressees || [])],
      ['word/_rels/document.xml.rels', DOCUMENT_RELS_XML],
      ['word/styles.xml', STYLES_XML],
    ]);
  }

  function safeName(text, limit = 60) {
    let s = String(text).replace(/[\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
    s = s.replace(/^[ .]+|[ .]+$/g, '').slice(0, limit).replace(/[ .]+$/, '');
    return s || 'без темы';
  }

  // Имя папки и файла документа: номер СЗ, например НТЦ-0001-26
  function fileBaseName(memo) {
    return safeName(memo.reg_number);
  }

  root.SZDocx = {
    addresseeName, buildDocx, dumpRich, fileBaseName, formatDate, formatLongDate, formatRegNumber, memoBlocks,
    normalizeRich, parseRich, richPlainText, RICH_SIZES, shortName, splitLines,
  };
})(typeof window !== 'undefined' ? window : globalThis);
