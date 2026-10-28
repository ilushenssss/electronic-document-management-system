// Проверка в Chromium: поле «Текст» с оформлением, предпросмотр, сохранение записки и документ Word.
//
// Запуск (нужны Node.js и Playwright с Chromium): node tests/browser-test.js
// Общая папка подменяется хранилищем браузера (OPFS), поэтому окно выбора папки не появляется.
'use strict';

const assert = require('assert');
const fs = require('fs');
const http = require('http');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = path.resolve(__dirname, '..');
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' };

function serve() {
  const server = http.createServer((req, res) => {
    const file = path.join(ROOT, decodeURIComponent(new URL(req.url, 'http://x').pathname));
    if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      res.writeHead(404);
      res.end();
      return;
    }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream' });
    fs.createReadStream(file).pipe(res);
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server)));
}

// Выделяет в поле текст sub
function selectText(page, sub) {
  return page.evaluate((s) => {
    const edit = document.querySelector('.rich-edit');
    edit.focus();
    const walker = document.createTreeWalker(edit, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = walker.nextNode())) {
      const i = n.data.indexOf(s);
      if (i >= 0) {
        const r = document.createRange();
        r.setStart(n, i);
        r.setEnd(n, i + s.length);
        getSelection().removeAllRanges();
        getSelection().addRange(r);
        return true;
      }
    }
    throw new Error(`Нет текста «${s}»`);
  }, sub);
}

const sheetBody = (page) => page.evaluate(() => {
  const ps = [...document.querySelectorAll('#preview .s-indent, #preview .s-p')];
  return ps.map((p) => p.innerHTML).join('\n');
});

const pressed = (page, sel) => page.getAttribute(`.rich-bar ${sel}`, 'aria-pressed');

async function main() {
  const server = await serve();
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });

  // Папка - внутреннее хранилище браузера с файлом SZ.html
  await page.addInitScript(() => {
    window.showDirectoryPicker = async () => {
      const dir = await navigator.storage.getDirectory();
      await dir.getFileHandle('SZ.html', { create: true });
      return dir;
    };
  });
  await page.goto(`http://127.0.0.1:${server.address().port}/SZ.html`);
  await page.waitForSelector('#login');

  // Справочники - сразу в базу
  await page.evaluate(async () => {
    const dir = await navigator.storage.getDirectory();
    for await (const name of dir.keys()) await dir.removeEntry(name, { recursive: true });
    await dir.getFileHandle('SZ.html', { create: true });
    const s = new window.SZStore.Store(window.SZ_SQL, dir);
    await s.open();
    await s.write((st) => {
      st.run("INSERT INTO topics (name) VALUES ('Закупка')");
      st.run("INSERT INTO employees (full_name, position, phone, is_signer) VALUES ('Сидоров Сидор Сидорович', 'Главный инженер', '12-34', 1)");
      st.run("INSERT INTO addressees (department, head_position, head_name, head_name_dative) VALUES ('Отдел кадров', 'Начальнику отдела кадров', 'Петров Пётр Петрович', 'Петрову П.П.')");
      st.setSetting('org_name', 'АО «НТЦ»');
      st.setSetting('city', 'Москва');
      // Хранилище OPFS в Chromium не создаёт файлы с русскими именами (в настоящей папке
      // Windows таких ограничений нет), поэтому в тесте номер латиницей
      st.setSetting('number_prefix', 'NTC');
    });
  });
  await page.click('#login');
  await page.waitForSelector('#side:not([hidden])');
  await page.evaluate(() => { location.hash = '#/new'; });
  await page.waitForSelector('.rich-edit');

  await page.selectOption('select[name=topic_id]', { index: 1 });
  await page.selectOption('select[name=executor_id]', { index: 1 });
  await page.check('.pick input');

  // Набор: абзацы и перенос строки
  await page.click('.rich-edit');
  await page.keyboard.type('Прошу срочно согласовать закупку.');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Перечень работ');
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.type('вторая строка');
  let html = await sheetBody(page);
  assert.match(html, /Перечень работ<br>вторая строка/, 'перенос строки на листе');

  // Полужирный и подчёркнутый для выделения
  await selectText(page, 'срочно');
  await page.click('[data-flag=bold]');
  await page.click('[data-flag=underline]');
  assert.strictEqual(await pressed(page, '[data-flag=bold]'), 'true');
  html = await sheetBody(page);
  assert.match(html, /<u><b>срочно<\/b><\/u>/, 'полужирный и подчёркнутый на листе');

  // Размер шрифта
  await selectText(page, 'согласовать');
  await page.selectOption('.rich-size', '12');
  html = await sheetBody(page);
  assert.match(html, /<span style="font-size:0\.8571em">согласовать<\/span>/, 'размер 12 на листе');
  await selectText(page, 'согласовать');
  assert.strictEqual(await page.inputValue('.rich-size'), '12', 'список размеров показывает размер выделения');

  // Выравнивание по центру и без отступа для второго абзаца
  await selectText(page, 'Перечень');
  await page.click('[data-align=center]');
  await page.click('[data-cmd=indent]');
  assert.strictEqual(await pressed(page, '[data-align=center]'), 'true');
  assert.strictEqual(await pressed(page, '[data-cmd=indent]'), 'false');

  // Ctrl+B без выделения: следующий набранный текст полужирный (раскладка не важна)
  await page.click('.rich-edit p:last-child');
  await page.keyboard.press('End');
  await page.keyboard.press('Control+KeyB');
  assert.strictEqual(await pressed(page, '[data-flag=bold]'), 'true');
  await page.keyboard.type(' важно');
  html = await sheetBody(page);
  assert.match(html, /<b> важно<\/b>/, 'Ctrl+B перед набором');

  // Ctrl+Z отменяет набор, Ctrl+Y возвращает
  await page.keyboard.press('Control+KeyZ');
  html = await sheetBody(page);
  assert.doesNotMatch(html, /важно/, 'отмена набора');
  await page.keyboard.press('Control+KeyY');
  assert.match(await sheetBody(page), /важно/, 'повтор набора');
  // Отмена изменения размера
  await selectText(page, 'закупку');
  await page.selectOption('.rich-size', '16');
  assert.match(await sheetBody(page), /font-size:1\.1429em">закупку/);
  await page.keyboard.press('Control+KeyZ');
  assert.doesNotMatch(await sheetBody(page), /font-size:1\.1429em/, 'отмена размера');

  // Вставка из Word: остаются только полужирный, курсив и подчёркивание
  await page.click('.rich-edit p:first-child');
  await page.keyboard.press('End');
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData('text/html', '<html><body><!--StartFragment--><p class=MsoNormal style="font-family:Comic Sans MS;' +
      'font-size:20pt;color:red;text-align:center">Из <b>Word</b> <i>с</i>\n <u>оформлением</u></p>\n' +
      '<p class=MsoNormal><span style="font-size:18pt">второй</span></p><!--EndFragment--></body></html>');
    dt.setData('text/plain', 'Из Word с оформлением\nвторой');
    document.querySelector('.rich-edit').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  html = await sheetBody(page);
  assert.match(html, /закупку\.Из <b>Word<\/b> <i>с<\/i> <u>оформлением<\/u>/, 'вставка из Word');
  assert.doesNotMatch(html, /Comic|color|20pt|18pt/, 'чужое оформление отброшено');

  if (process.env.SHOT_NEW) {
    await selectText(page, 'срочно');
    await page.screenshot({ path: process.env.SHOT_NEW, fullPage: true });
  }

  // Сохранение
  await page.click('#memo-form button[type=submit]');
  await page.waitForFunction(() => location.hash.startsWith('#/memo/'));
  await page.waitForSelector('.flash');
  const flashText = await page.innerText('.flash');
  assert.ok(await page.$('.flash-success'), 'сообщение: ' + flashText);
  const saved = await page.evaluate(async () => {
    const dir = await navigator.storage.getDirectory();
    const s = new window.SZStore.Store(window.SZ_SQL, dir);
    await s.open();
    const memo = s.one('SELECT * FROM memos');
    const file = await s.readDocument(memo.file_path);
    return { memo, docx: new TextDecoder().decode(await file.arrayBuffer()) };
  });
  const rich = JSON.parse(saved.memo.body_rich);
  assert.deepStrictEqual(rich, [
    { align: 'both', indent: true, runs: [
      { text: 'Прошу ' }, { text: 'срочно', bold: true, underline: true }, { text: ' ' },
      { text: 'согласовать', size: 12 }, { text: ' закупку.Из ' }, { text: 'Word', bold: true }, { text: ' ' },
      { text: 'с', italic: true }, { text: ' ' }, { text: 'оформлением', underline: true }] },
    { align: 'both', indent: true, runs: [{ text: 'второй' }] },
    { align: 'center', indent: false, runs: [{ text: 'Перечень работ\nвторая строка' }, { text: ' важно', bold: true }] },
  ], 'body_rich в базе');
  assert.ok(saved.memo.body.startsWith('Прошу срочно согласовать закупку.'), 'обычный текст для поиска');
  // Документ Word (ZIP без сжатия - XML читается как есть)
  assert.ok(saved.docx.includes('<w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">срочно</w:t>'), 'Word: полужирный и подчёркнутый');
  assert.ok(saved.docx.includes('<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">согласовать'), 'Word: размер 12');
  assert.ok(saved.docx.includes('Перечень работ</w:t><w:br/><w:t xml:space="preserve">вторая строка'), 'Word: перенос строки');
  assert.ok(saved.docx.includes('w:ascii="Arial"'), 'Word: шрифт Arial');
  // Лист на странице записки
  const memoSheet = await page.innerHTML('.sheet');
  assert.match(memoSheet, /<u><b>срочно<\/b><\/u>/, 'лист на странице записки');

  if (process.env.SHOT) await page.screenshot({ path: process.env.SHOT });
  assert.deepStrictEqual(errors, [], 'ошибки в консоли');
  await browser.close();
  server.close();
  console.log('Все проверки в браузере пройдены');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
