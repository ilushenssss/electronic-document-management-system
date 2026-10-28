// Проверка формирования документа с оформлением текста. Запуск: node tests/docx-test.js
'use strict';

const assert = require('assert');
require('../app/docx.js');

const D = globalThis.SZDocx;

const RICH = [
  { align: 'both', indent: true, runs: [{ text: 'Прошу ' }, { text: 'срочно', bold: true, underline: true },
    { text: ' согласовать', italic: true, size: 12 }] },
  { align: 'center', indent: false, runs: [{ text: 'Первая строка\nвторая строка' }] },
];

// Абзацы чистятся: пустые убираются, недопустимое оформление отбрасывается, фрагменты склеиваются
assert.deepStrictEqual(D.normalizeRich([
  { align: 'weird', runs: [{ text: '  а' }, { text: 'б' }, { text: 'в ', bold: 1 }, { text: 'x', size: 99, color: 'red' }, { text: 5 }] },
  { runs: [{ text: '   ' }] },
  'мусор',
]), [{ align: 'both', indent: false, runs: [{ text: 'аб' }, { text: 'в ', bold: true }, { text: 'x' }] }]);

// Текст без оформления хранится как раньше - без body_rich
assert.strictEqual(D.dumpRich([{ align: 'both', indent: true, runs: [{ text: 'Абзац' }] }]), null);
assert.deepStrictEqual(D.parseRich(D.dumpRich(RICH)), RICH);
assert.strictEqual(D.parseRich('{битый'), null);
assert.strictEqual(D.richPlainText(RICH), 'Прошу срочно согласовать\nПервая строка\nвторая строка');

function documentXml(memo) {
  const zip = Buffer.from(D.buildDocx(memo, { addressees: [] }));
  // ZIP без сжатия: ищем word/document.xml по локальным заголовкам
  for (let i = 0; zip.readUInt32LE(i) === 0x04034b50;) {
    const nameLen = zip.readUInt16LE(i + 26);
    const size = zip.readUInt32LE(i + 18);
    const name = zip.slice(i + 30, i + 30 + nameLen).toString();
    if (name === 'word/document.xml') return zip.slice(i + 30 + nameLen, i + 30 + nameLen + size).toString();
    i += 30 + nameLen + size;
  }
  throw new Error('нет word/document.xml');
}

const xml = documentXml({ reg_number: 'X', memo_date: '2026-01-01', body: 'игнорируется', body_rich: D.dumpRich(RICH) });
assert.ok(!xml.includes('игнорируется'));
// Порядок элементов оформления задан схемой Word: b, i, sz, u
assert.ok(xml.includes('<w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/><w:u w:val="single"/></w:rPr>' +
  '<w:t xml:space="preserve">срочно</w:t>'));
assert.ok(xml.includes('<w:rPr><w:i/><w:iCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'));
assert.ok(xml.includes('Первая строка</w:t><w:br/><w:t xml:space="preserve">вторая строка'));
assert.ok(xml.includes('<w:jc w:val="center"/>'));

// Без body_rich - обычные абзацы по ширине с отступом
const plain = documentXml({ reg_number: 'X', body: 'Абзац один\nАбзац два' });
assert.ok(plain.includes('<w:ind w:firstLine="709"/><w:jc w:val="both"/>'));
assert.ok(plain.includes('>Абзац два<'));

console.log('Проверки документа пройдены');
