// Формирование Word-документа (.docx) служебной записки без внешних библиотек.
(function (root) {
  'use strict';

  const FONT = 'Calibri';
  const FONT_SIZE = 28; // в полупунктах: 14 pt
  const CM = 567; // твипов в сантиметре

  function formatDate(iso) {
    const [y, m, d] = iso.split('-');
    return `${d}.${m}.${y}`;
  }

  // НТЦ-0001-26: префикс, порядковый номер с ведущими нулями, две цифры года
  function formatRegNumber(prefix, number, year) {
    return `${prefix}-${String(number).padStart(4, '0')}-${String(year).slice(-2)}`;
  }

  function xml(text) {
    return String(text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      // Управляющие символы недопустимы в XML
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  }

  function run(text, bold) {
    return `<w:r>${bold ? '<w:rPr><w:b/><w:bCs/></w:rPr>' : ''}` +
      `<w:t xml:space="preserve">${xml(text)}</w:t></w:r>`;
  }

  // runs: строка или массив [текст, жирный]
  function para(runs, opts = {}) {
    const pPr = [];
    if (opts.after) pPr.push(`<w:spacing w:after="${opts.after}"/>`);
    if (opts.left || opts.firstLine) {
      pPr.push(`<w:ind${opts.left ? ` w:left="${opts.left}"` : ''}` +
        `${opts.firstLine ? ` w:firstLine="${opts.firstLine}"` : ''}/>`);
    }
    if (opts.borderBottom) {
      pPr.push('<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="6" w:color="000000"/></w:pBdr>');
    }
    if (opts.align) pPr.push(`<w:jc w:val="${opts.align}"/>`);
    const list = typeof runs === 'string' ? [[runs, opts.bold]] : runs;
    return `<w:p>${pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : ''}` +
      list.filter(([t]) => t).map(([t, b]) => run(t, b)).join('') + '</w:p>';
  }

  function documentXml(memo, orgName) {
    const body = [];

    // Шапка: наименование организации, каждая строка по центру, снизу черта
    const orgLines = String(orgName || '').split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    orgLines.forEach((line, i) => {
      const last = i === orgLines.length - 1;
      body.push(para(line, { bold: true, align: 'center', borderBottom: last, after: last ? 480 : 0 }));
    });

    for (const line of [memo.addressee_position, memo.addressee_department,
      memo.addressee_name]) {
      body.push(para(line, { left: Math.round(9.5 * CM) }));
    }
    body.push(para(''), para(''));
    body.push(para('СЛУЖЕБНАЯ ЗАПИСКА', { bold: true, align: 'center' }));
    body.push(para(`№ ${memo.reg_number} от ${formatDate(memo.memo_date)}`,
      { align: 'center', after: 240 }));
    body.push(para([['Тема: ', true], [memo.topic_name, false]], { after: 240 }));

    for (const line of memo.body.split(/\r?\n/)) {
      if (line.trim()) {
        body.push(para(line.trim(),
          { align: 'both', firstLine: Math.round(1.25 * CM) }));
      }
    }

    body.push(para(''), para(''));
    body.push(para(`Ответственный исполнитель: ${memo.executor_name}`));
    if (memo.executor_phone) body.push(para(`Тел.: ${memo.executor_phone}`));

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:body>${body.join('')}` +
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
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
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
      ['word/document.xml', documentXml(memo, opts.orgName)],
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

  root.SZDocx = { buildDocx, fileBaseName, formatDate, formatRegNumber };
})(typeof window !== 'undefined' ? window : globalThis);
