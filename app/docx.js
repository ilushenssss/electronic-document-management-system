// Формирование Word-документа (.docx) служебной записки без внешних библиотек.
(function (root) {
  'use strict';

  const FONT = 'Calibri';
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

  // Содержимое записки как список блоков. По нему строятся и документ Word,
  // и лист на экране, поэтому они не расходятся.
  //   { text | runs, bold, align, firstLine, size } - абзац
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
      add(...addressees.flatMap((a) => [
        { text: a.head_position, align: 'right' }, { text: addresseeName(a), align: 'right' }]));
    } else if (preview) {
      add({ text: { ph: 'Должность адресата' }, align: 'right' }, { text: { ph: 'Фамилия И.О.' }, align: 'right' });
    }

    add({ text: v(m.topic_name, 'Тематика работ') });

    add({ text: greeting(addressees) || v('', 'Уважаемый Имя Отчество!'), align: 'center' });

    const body = splitLines(m.body);
    if (body.length) add(...body.map((line) => ({ text: line, align: 'both', firstLine: true })));
    else if (preview) add({ text: { ph: 'Текст служебной записки' }, align: 'both', firstLine: true });

    const attachment = splitLines(m.attachment);
    if (attachment.length) {
      add(...attachment.map((line, i) => ({ runs: [[i ? '' : 'Приложение: '], [line]], align: 'both' })));
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

  // Размер задаётся у каждого фрагмента явно, чтобы не зависеть от настроек Word
  function run(text, bold, size = FONT_SIZE / 2) {
    const rPr = (bold ? '<w:b/><w:bCs/>' : '') +
      `<w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/>`;
    const content = text === '\t' ? '<w:tab/>' : `<w:t xml:space="preserve">${xml(text)}</w:t>`;
    return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${content}</w:r>`;
  }

  function blockXml(blk) {
    const pPr = [];
    let runs;
    if (blk.blank) {
      runs = [];
    } else if (blk.split) {
      // Правый край - через табуляцию по правому краю текста
      pPr.push(`<w:tabs><w:tab w:val="right" w:pos="${TEXT_WIDTH}"/></w:tabs>`);
      runs = [[blk.split[0]], ['\t'], [blk.split[1]]];
    } else {
      runs = blk.runs || [[blk.text, blk.bold]];
    }
    if (blk.firstLine) pPr.push(`<w:ind w:firstLine="${Math.round(1.25 * CM)}"/>`);
    if (blk.align) pPr.push(`<w:jc w:val="${blk.align}"/>`);
    const size = blk.size || FONT_SIZE / 2;
    const rPrDefault = `<w:rPr><w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/></w:rPr>`;
    return `<w:p>${pPr.length || rPrDefault ? `<w:pPr>${pPr.join('')}${rPrDefault}</w:pPr>` : ''}` +
      runs.filter(([t]) => t).map(([t, b]) => run(t, b, size)).join('') + '</w:p>';
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
    addresseeName, buildDocx, fileBaseName, formatDate, formatLongDate, formatRegNumber, memoBlocks,
    shortName, splitLines,
  };
})(typeof window !== 'undefined' ? window : globalThis);
