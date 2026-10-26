// Интерфейс: подключение к общей папке, страницы и обработчики форм.
(function () {
  'use strict';

  const { Store, DEFAULT_OUTPUT_DIR, DEFAULT_NUMBER_PREFIX } = window.SZStore;
  const { addresseeName, formatDate, formatRegNumber, memoBlocks, shortName, splitLines } = window.SZDocx;

  // Описание справочников: заголовок, поля (имя, подпись, обязательное), сортировка
  const DIRECTORIES = {
    employees: {
      title: 'Сотрудники',
      // 4-й элемент поля - многострочное (переносы сохраняются и печатаются)
      fields: [['full_name', 'ФИО полностью', true], ['position', 'Должность', false, true],
        ['phone', 'Номер телефона', false]],
      signer: true, // колонка «Подписант»
      order: 'full_name',
    },
    addressees: {
      title: 'Адресаты',
      fields: [['department', 'Отдел', true],
        ['head_position', 'Должность в документе, например «Начальнику отдела»', true],
        ['head_name', 'ФИО руководителя полностью', true],
        ['head_name_dative', 'Фамилия И.О. в дательном падеже, например «Петрову П.П.»', true]],
      order: 'department',
    },
    topics: {
      title: 'Тематики работ',
      fields: [['name', 'Название', true]],
      order: 'name',
    },
  };

  const MONTHS = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
  const MONTHS_FULL = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август',
    'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const MONTHS_IN = ['январе', 'феврале', 'марте', 'апреле', 'мае', 'июне', 'июле', 'августе',
    'сентябре', 'октябре', 'ноябре', 'декабре'];

  const main = document.getElementById('main');
  const side = document.getElementById('side');
  const nav = document.getElementById('nav');
  let store = null;
  let rootHandle = null;
  let flashes = [];
  let editId = null; // редактируемая запись справочника
  let listFilter = { q: '', topic: '', year: '' };
  let statsYear = new Date().getFullYear();
  let statsAsTable = false;

  // ---------- Утилиты ----------

  function esc(value) {
    return String(value ?? '').replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    })[c]);
  }

  function flash(message, kind = 'success') {
    flashes.push([kind, message]);
  }

  function flashHtml() {
    const html = flashes.map(([k, m]) => `<div class="flash flash-msg flash-${k}" role="status">${esc(m)}</div>`).join('');
    flashes = [];
    return html;
  }

  // Показывает сообщения над страницей, не перерисовывая её (введённые данные остаются)
  function showFlashes() {
    for (const el of main.querySelectorAll('.flash-msg')) el.remove();
    main.insertAdjacentHTML('afterbegin', flashHtml());
    window.scrollTo(0, 0);
  }

  function showError(e) {
    console.error(e);
    flash(e && e.message ? e.message : String(e), 'error');
    showFlashes();
  }

  function today() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  function options(rows, selected, label) {
    return rows.map((r) => `<option value="${r.id}" ${String(r.id) === String(selected) ? 'selected' : ''}>${esc(label(r))}</option>`).join('');
  }

  function nextNumber(year) {
    return store.one('SELECT COALESCE(MAX(number), 0) + 1 AS n FROM memos WHERE year = ?', [year]).n;
  }

  // Выполняет сохранение и перерисовывает страницу. Если пользователь за это
  // время ушёл на другую страницу, её не трогаем: сообщение покажется позже.
  async function busy(button, action) {
    const hash = location.hash;
    const text = button.textContent;
    button.disabled = true;
    button.textContent = 'Сохранение…';
    try {
      await action();
      if (location.hash === hash) render();
    } catch (e) {
      console.error(e);
      flash(e && e.message ? e.message : String(e), 'error');
      if (location.hash === hash) showFlashes();
    } finally {
      button.disabled = false;
      button.textContent = text;
    }
  }

  function download(file, name) {
    const url = URL.createObjectURL(file);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  // ---------- Лист служебной записки (как будет в Word) ----------

  // Лист строится из тех же блоков, что и документ Word (SZDocx.memoBlocks)
  function sheetHtml(m, addressees, { preview = false } = {}) {
    const t = (x) => (x && typeof x === 'object' ? `<span class="ph">${esc(x.ph)}</span>` : esc(x));
    const html = memoBlocks(m, addressees, { preview }).map((b) => {
      if (b.blank) return '<p class="s-p">&nbsp;</p>';
      const cls = ['s-p'];
      if (b.align) cls.push(`s-${b.align}`);
      if (b.firstLine) cls.push('s-indent');
      if (b.size) cls.push('s-small');
      if (b.split) {
        cls.push('s-split');
        return `<p class="${cls.join(' ')}"><span>${t(b.split[0])}</span><span>${t(b.split[1])}</span></p>`;
      }
      const runs = b.runs || [[b.text, b.bold]];
      return `<p class="${cls.join(' ')}">${runs.map(([x, bold]) => (bold ? `<b>${t(x)}</b>` : t(x))).join('')}</p>`;
    }).join('');
    return `<article class="sheet" aria-label="Лист служебной записки">${html}</article>`;
  }

  // Записки прошлых версий не хранят шапку - показываем текущую из настроек
  function memoForSheet(memo) {
    return { ...memo, org_name: memo.org_name ?? store.orgName };
  }

  // ---------- Подсказки при наведении ----------

  const tip = document.getElementById('tip');

  function placeTip(x, y) {
    const pad = 12;
    const r = tip.getBoundingClientRect();
    let left = x + pad;
    let top = y + pad;
    if (left + r.width > window.innerWidth - 8) left = x - r.width - pad;
    if (top + r.height > window.innerHeight - 8) top = y - r.height - pad;
    tip.style.left = `${Math.max(8, left)}px`;
    tip.style.top = `${Math.max(8, top)}px`;
  }

  function showTip(target, x, y) {
    tip.textContent = target.dataset.tip;
    tip.hidden = false;
    placeTip(x, y);
  }

  document.addEventListener('mouseover', (e) => {
    const t = e.target.closest && e.target.closest('[data-tip]');
    if (t) showTip(t, e.clientX, e.clientY);
    else tip.hidden = true;
  });
  document.addEventListener('mousemove', (e) => {
    if (!tip.hidden) placeTip(e.clientX, e.clientY);
  });
  document.addEventListener('focusin', (e) => {
    const t = e.target.closest && e.target.closest('[data-tip]');
    if (!t) { tip.hidden = true; return; }
    const r = t.getBoundingClientRect();
    showTip(t, r.left + r.width / 2, r.top);
  });
  document.addEventListener('focusout', () => { tip.hidden = true; });

  // ---------- Подключение к папке ----------

  const IDB_NAME = 'sz-memos';

  function idb(mode, fn) {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore('kv');
      req.onerror = () => reject(req.error);
      req.onsuccess = () => {
        try {
          const tx = req.result.transaction('kv', mode);
          const r = fn(tx.objectStore('kv'));
          tx.oncomplete = () => resolve(r && r.result);
          tx.onerror = () => reject(tx.error);
        } catch (e) {
          reject(e);
        }
      };
    });
  }

  const savedHandle = () => idb('readonly', (s) => s.get('root')).catch(() => null);
  const saveHandle = (h) => idb('readwrite', (s) => s.put(h, 'root')).catch(() => {});

  // Папка, где лежит SZ.html, в привычном виде: \\сервер\папка или C:\папка
  function folderOfThisPage() {
    const path = decodeURIComponent(location.pathname).replace(/\/[^/]*$/, '');
    if (location.host) return `\\\\${location.host}${path.replace(/\//g, '\\')}`;
    return path.replace(/^\/([A-Za-z]:)/, '$1').replace(/\//g, '\\');
  }

  // Имя папки для сравнения с выбранной: «СЗ», «Z:» и т. п.
  const folderKey = (name) => String(name).replace(/[\\/:]+$/g, '').replace(/[\\/:]/g, '').toLowerCase();

  function isPageFolder(handle) {
    const parts = folderOfThisPage().split('\\').filter(Boolean);
    return folderKey(handle.name) === folderKey(parts[parts.length - 1] || '');
  }

  // Браузер не даёт доступ к папке по пути, только через окно выбора.
  // Поэтому при первом входе на компьютере папку выбирают один раз, дальше
  // «Войти» только спрашивает разрешение для запомненной папки.
  async function pickFolder() {
    let picked;
    try {
      picked = await window.showDirectoryPicker({ id: 'sz-memos', mode: 'readwrite' });
    } catch (e) {
      if (e.name !== 'AbortError') renderConnect(null, e.message);
      return;
    }
    let hasPage = true;
    try {
      await picked.getFileHandle('SZ.html');
    } catch {
      hasPage = false;
    }
    if (!hasPage) {
      renderConnect(null, `В папке «${picked.name}» нет SZ.html. Нажмите «Войти» ещё раз и выберите папку ${folderOfThisPage()}`);
      return;
    }
    await saveHandle(picked);
    await connect(picked);
  }

  async function login(handle) {
    if (!handle) {
      await pickFolder();
      return;
    }
    try {
      if (await handle.requestPermission({ mode: 'readwrite' }) === 'granted') {
        await connect(handle);
      } else {
        renderConnect(handle, 'Без разрешения на доступ к папке войти нельзя. Нажмите «Войти» и выберите «Разрешить».');
      }
    } catch (e) {
      // Папку переместили или удалили: выбираем заново
      console.error(e);
      await pickFolder();
    }
  }

  function renderConnect(handle, message) {
    side.hidden = true;
    store = null;
    main.innerHTML = `
      <section class="connect">
        <h1>Служебные записки</h1>
        ${message ? `<div class="flash flash-error">${esc(message)}</div>` : ''}
        <p>Реестр и документы Word хранятся в папке</p>
        <p class="path">${esc(folderOfThisPage())}</p>
        <div class="actions">
          <button class="btn primary btn-login" id="login">Войти</button>
        </div>
        ${handle
          ? `<p class="muted small">Браузер спросит разрешение на доступ к папке. Чтобы он больше
              не спрашивал, выберите «Разрешать при каждом посещении».</p>`
          : `<p class="muted small">Первый вход на этом компьютере: браузер откроет окно выбора папки.
              Выберите папку, указанную выше, и нажмите «Выбрать папку», затем «Разрешить».
              Это нужно сделать один раз, дальше достаточно нажимать «Войти».</p>`}
      </section>`;
    document.getElementById('login').onclick = () => login(handle);
  }

  async function connect(handle) {
    main.innerHTML = '<div class="loading">Загрузка реестра…</div>';
    try {
      const s = new Store(window.SZ_SQL, handle);
      await s.open();
      store = s;
      rootHandle = handle;
    } catch (e) {
      console.error(e);
      renderConnect(handle, `Не удалось открыть базу: ${e.message}`);
      return;
    }
    side.hidden = false;
    if (!location.hash) location.hash = '#/list';
    render();
  }

  function renderSide() {
    const org = store.orgName.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    document.getElementById('side-org').textContent = org.join(' ');
    document.getElementById('side-foot').innerHTML =
      `<div class="side-folder" title="Общая папка">${esc(folderOfThisPage())}</div>`;
  }

  // ---------- Маршрутизация ----------

  function route() {
    const parts = location.hash.replace(/^#\/?/, '').split('/');
    return { page: parts[0] || 'list', arg: parts[1] };
  }

  function render() {
    if (!store) return;
    const { page, arg } = route();
    for (const a of nav.querySelectorAll('a')) {
      const active = a.dataset.page === page || (a.dataset.page === 'list' && page === 'memo');
      a.classList.toggle('active', active);
      if (active) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    }
    tip.hidden = true;
    try {
      renderSide();
      if (page === 'new') renderNew();
      else if (page === 'dirs') renderDirs(arg in DIRECTORIES || arg === 'settings' ? arg : 'employees');
      else if (page === 'memo') renderMemo(Number(arg)).catch(showError);
      else if (page === 'stats') renderStats();
      else renderList();
    } catch (e) {
      console.error(e);
      main.innerHTML = `<div class="flash flash-error">Ошибка: ${esc(e.message)}</div>`;
    }
    window.scrollTo(0, 0);
  }

  // ---------- Новая СЗ ----------

  function renderNew() {
    const topics = store.all('SELECT * FROM topics ORDER BY name');
    const addressees = store.all('SELECT * FROM addressees ORDER BY department, head_name');
    const employees = store.all('SELECT * FROM employees ORDER BY full_name');
    const signer = store.signer;
    const missing = [
      !topics.length && '<a href="#/dirs/topics">тематики работ</a>',
      !addressees.length && '<a href="#/dirs/addressees">адресатов</a>',
      !employees.length && '<a href="#/dirs/employees">сотрудников</a>',
    ].filter(Boolean);

    main.innerHTML = `
      ${flashHtml()}
      <header class="page-head"><h1>Новая служебная записка</h1></header>
      ${missing.length ? `<div class="flash flash-error">Чтобы создать записку, добавьте в справочники ${missing.join(', ')}.</div>` : ''}
      ${employees.length && !signer ? '<div class="flash flash-error">Подписант не назначен. Отметьте его галочкой в справочнике <a href="#/dirs/employees">«Сотрудники»</a>.</div>' : ''}
      <div class="compose">
        <form class="panel form" id="memo-form">
          <label class="field"><span>Тематика работ</span>
            <select name="topic_id" required><option value="">Выберите тематику</option>
              ${options(topics, '', (t) => t.name)}</select></label>
          <fieldset class="field picker">
            <legend>Кому</legend>
            ${addressees.length > 6 ? '<input type="search" class="picker-filter" id="addr-filter" placeholder="Найти отдел или фамилию" aria-label="Найти адресата">' : ''}
            <div class="picker-list">
              ${addressees.map((a) => `
                <label class="pick" data-text="${esc(`${a.department} ${a.head_position} ${a.head_name}`.toLocaleLowerCase('ru'))}">
                  <input type="checkbox" value="${a.id}">
                  <span class="pick-text">${esc(a.department)}<small>${esc(`${a.head_position} ${addresseeName(a)}`)}</small></span>
                  <span class="pick-order" aria-hidden="true"></span>
                </label>`).join('')}
            </div>
            <small>Можно выбрать несколько. В записке адресаты идут в порядке выбора.</small>
          </fieldset>
          <label class="field field-date"><span>Дата</span>
            <input type="date" name="memo_date" value="${today()}" required></label>
          <label class="field"><span>Текст</span>
            <textarea name="body" rows="9" required placeholder="Каждый абзац с новой строки"></textarea></label>
          <label class="field"><span>Приложение <i>необязательно</i></span>
            <textarea name="attachment" rows="2" placeholder="Например: смета на 2 л. в 1 экз."></textarea></label>
          <div class="field"><span>Подписант</span>
            <div class="signer-line">${signer
              ? `${esc(shortName(signer.full_name))}<small class="pre">${esc(signer.position) || 'должность не указана'}</small>`
              : '<small>не назначен</small>'}
              <a href="#/dirs/employees">Изменить</a></div></div>
          <label class="field"><span>Исполнитель</span>
            <select name="executor_id" required><option value="">Выберите сотрудника</option>
              ${options(employees, '', (e) => e.full_name + (e.phone ? `, ${e.phone}` : ''))}</select></label>
          <div class="form-foot">
            <button type="submit" class="btn primary">Создать и сформировать Word</button>
            <p class="muted small" id="number-note"></p>
          </div>
        </form>
        <div class="sheet-wrap" id="preview"></div>
      </div>`;

    const form = document.getElementById('memo-form');
    const byId = (rows, id) => rows.find((r) => String(r.id) === String(id)) || {};
    let chosen = []; // id адресатов в порядке выбора

    function updatePicker() {
      for (const box of form.querySelectorAll('.pick input')) {
        const pos = chosen.indexOf(box.value);
        box.closest('.pick').querySelector('.pick-order').textContent = pos >= 0 ? pos + 1 : '';
      }
    }

    function updatePreview() {
      const d = Object.fromEntries(new FormData(form));
      const year = /^\d{4}/.test(d.memo_date) ? Number(d.memo_date.slice(0, 4)) : new Date().getFullYear();
      const reg = formatRegNumber(store.numberPrefix, nextNumber(year), year);
      const executor = byId(employees, d.executor_id);
      document.getElementById('preview').innerHTML = sheetHtml({
        reg_number: reg, memo_date: d.memo_date, topic_name: byId(topics, d.topic_id).name,
        body: d.body, attachment: d.attachment, org_name: store.orgName, city: store.city,
        signer_name: signer && signer.full_name, signer_position: signer && signer.position,
        executor_name: executor.full_name, executor_phone: executor.phone,
      }, chosen.map((id) => byId(addressees, id)), { preview: true });
      document.getElementById('number-note').textContent = `Номер ${reg} закрепится при сохранении.`;
    }

    form.addEventListener('change', (ev) => {
      if (ev.target.matches('.pick input')) {
        const id = ev.target.value;
        chosen = ev.target.checked ? [...chosen, id] : chosen.filter((x) => x !== id);
        updatePicker();
      }
    });
    const filter = document.getElementById('addr-filter');
    if (filter) {
      filter.addEventListener('input', () => {
        const q = filter.value.trim().toLocaleLowerCase('ru');
        for (const pick of form.querySelectorAll('.pick')) pick.hidden = q && !pick.dataset.text.includes(q);
      });
    }
    form.addEventListener('input', updatePreview);
    form.addEventListener('change', updatePreview);
    updatePreview();

    form.onsubmit = (ev) => {
      ev.preventDefault();
      const data = Object.fromEntries(new FormData(form));
      busy(form.querySelector('button[type=submit]'), () => createMemo(data, chosen));
    };
  }

  async function createMemo(data, addresseeIds) {
    const body = data.body.trim();
    if (!addresseeIds.length) throw new Error('Выберите хотя бы одного адресата.');
    if (!body) throw new Error('Заполните текст служебной записки.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data.memo_date)) throw new Error('Укажите дату.');
    const year = Number(data.memo_date.slice(0, 4));

    const memoId = await store.write((s) => {
      const topic = s.one('SELECT * FROM topics WHERE id = ?', [data.topic_id]);
      const executor = s.one('SELECT * FROM employees WHERE id = ?', [data.executor_id]);
      const signer = s.signer;
      const addrs = addresseeIds.map((id) => s.one('SELECT * FROM addressees WHERE id = ?', [id]));
      if (!signer) {
        throw new Error('Подписант не назначен. Отметьте его галочкой в справочнике «Сотрудники».');
      }
      if (!topic || !executor || addrs.some((a) => !a)) {
        throw new Error('Выбранную запись справочника кто-то удалил. Обновите страницу и выберите заново.');
      }
      const noDative = addrs.find((a) => !a.head_name_dative);
      if (noDative) {
        throw new Error(`У адресата ${noDative.head_name} не заполнена Фамилия И.О. в дательном падеже. Заполните её в справочнике «Адресаты».`);
      }
      if (!signer.position) {
        throw new Error(`У подписанта ${signer.full_name} не указана должность. Заполните её в справочнике «Сотрудники».`);
      }
      // Номер считается под блокировкой, поэтому у двух пользователей он не совпадёт
      const number = s.one('SELECT COALESCE(MAX(number), 0) + 1 AS n FROM memos WHERE year = ?', [year]).n;
      // Колонки addressee_* - сводка для поиска; сами адресаты - в memo_addressees
      const joined = (key) => addrs.map((a) => a[key]).join('; ');
      const id = s.run(
        `INSERT INTO memos (number, year, reg_number, memo_date, topic_id, topic_name,
           addressee_id, addressee_department, addressee_position, addressee_name,
           body, attachment, signer_id, signer_name, signer_position,
           executor_id, executor_name, executor_phone, city, org_name)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [number, year, formatRegNumber(s.numberPrefix, number, year), data.memo_date,
          topic.id, topic.name, addrs[0].id, joined('department'), joined('head_position'),
          joined('head_name'), body, data.attachment.trim(), signer.id, signer.full_name,
          signer.position, executor.id, executor.full_name, executor.phone, s.city, s.orgName]);
      addrs.forEach((a, i) => s.run(
        `INSERT INTO memo_addressees (memo_id, ord, addressee_id, department, head_position, head_name,
           head_name_dative) VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [id, i, a.id, a.department, a.head_position, a.head_name, a.head_name_dative]));
      return id;
    });

    const memo = store.one('SELECT * FROM memos WHERE id = ?', [memoId]);
    try {
      const path = await store.saveDocument(memo);
      await store.write((s) => s.run('UPDATE memos SET file_path = ? WHERE id = ?', [path, memoId]));
      flash(`Записка ${memo.reg_number} создана, документ Word сохранён в папку.`);
    } catch (e) {
      console.error(e);
      flash(`Записка ${memo.reg_number} сохранена в реестре, но документ Word создать не удалось: ${e.message}`, 'error');
    }
    location.hash = `#/memo/${memoId}`; // hashchange отрисует страницу
  }

  // ---------- Реестр ----------

  function addresseesCell(list) {
    const shown = list.slice(0, 2).map((a) =>
      `<div>${esc(a.department)}<div class="sub">${esc(shortName(a.head_name))}</div></div>`).join('');
    return shown + (list.length > 2 ? `<div class="sub">и ещё ${list.length - 2}</div>` : '');
  }

  function renderList() {
    const { q, topic, year } = listFilter;
    let sql = 'SELECT * FROM memos WHERE 1=1';
    const params = [];
    if (topic) { sql += ' AND topic_id = ?'; params.push(Number(topic)); }
    if (year) { sql += ' AND year = ?'; params.push(Number(year)); }
    let memos = store.all(`${sql} ORDER BY year DESC, number DESC`, params);
    if (q) {
      // LIKE в SQLite не учитывает регистр кириллицы, поэтому ищем в JS
      const needle = q.toLocaleLowerCase('ru');
      memos = memos.filter((m) => [m.reg_number, m.topic_name, m.addressee_department,
        m.addressee_name, m.executor_name, m.signer_name, m.body, m.attachment]
        .join(' ').toLocaleLowerCase('ru').includes(needle));
    }
    const addrMap = {};
    for (const a of store.all('SELECT * FROM memo_addressees ORDER BY memo_id, ord')) {
      (addrMap[a.memo_id] ||= []).push(a);
    }
    const topics = store.all('SELECT * FROM topics ORDER BY name');
    const years = store.all('SELECT DISTINCT year FROM memos ORDER BY year DESC').map((r) => r.year);
    const total = store.one('SELECT COUNT(*) AS n FROM memos').n;
    const filtered = q || topic || year;

    main.innerHTML = `
      ${flashHtml()}
      <header class="page-head">
        <h1>Реестр</h1>
        <a class="btn primary" href="#/new">Новая СЗ</a>
      </header>
      ${total ? `
      <form class="toolbar" id="filters" role="search">
        <input type="search" name="q" value="${esc(q)}" placeholder="Номер, тема, адресат, исполнитель или слово из текста" aria-label="Поиск">
        <select name="topic" aria-label="Тематика"><option value="">Все тематики</option>${options(topics, topic, (t) => t.name)}</select>
        <select name="year" aria-label="Год"><option value="">Все годы</option>
          ${years.map((y) => `<option ${String(y) === year ? 'selected' : ''}>${y}</option>`).join('')}</select>
        <button type="submit" class="btn">Найти</button>
        ${filtered ? '<button type="button" class="btn ghost" id="reset">Сбросить</button>' : ''}
      </form>` : ''}
      ${memos.length ? `
        <div class="panel table-wrap"><table class="table">
          <thead><tr><th>Номер</th><th>Дата</th><th>Тематика</th><th>Кому</th><th>Исполнитель</th><th><span class="sr">Действия</span></th></tr></thead>
          <tbody>${memos.map((m) => `
            <tr>
              <td class="num"><a href="#/memo/${m.id}">${esc(m.reg_number)}</a></td>
              <td class="nowrap">${formatDate(m.memo_date)}</td>
              <td>${esc(m.topic_name)}</td>
              <td>${addresseesCell(addrMap[m.id] || [])}</td>
              <td>${esc(m.executor_name)}</td>
              <td class="row-actions"><button class="btn small" data-download="${m.id}">Word</button></td>
            </tr>`).join('')}
          </tbody></table></div>
        <p class="muted small">${filtered ? `Найдено ${memos.length} из ${total}` : `Всего ${total}`}</p>`
        : total
          ? '<div class="empty"><p>По этим условиям ничего не нашлось.</p></div>'
          : `<div class="empty"><p>Здесь появятся созданные служебные записки.</p>
              <a class="btn primary" href="#/new">Создать первую</a></div>`}`;

    const form = document.getElementById('filters');
    if (form) {
      form.onsubmit = (ev) => {
        ev.preventDefault();
        const d = Object.fromEntries(new FormData(form));
        listFilter = { q: d.q.trim(), topic: d.topic, year: d.year };
        render();
      };
    }
    const reset = document.getElementById('reset');
    if (reset) reset.onclick = () => { listFilter = { q: '', topic: '', year: '' }; render(); };
    for (const b of main.querySelectorAll('[data-download]')) {
      b.onclick = () => downloadMemo(Number(b.dataset.download)).catch(showError);
    }
  }

  async function downloadMemo(id) {
    const memo = store.one('SELECT * FROM memos WHERE id = ?', [id]);
    const file = memo && await store.readDocument(memo.file_path);
    if (!file) {
      flash('Документ Word не найден в папке. Откройте записку и сформируйте его заново.', 'error');
      showFlashes();
      return;
    }
    download(file, memo.file_path.split('/').pop());
  }

  // ---------- Просмотр СЗ ----------

  async function renderMemo(id) {
    const memo = store.one('SELECT * FROM memos WHERE id = ?', [id]);
    if (!memo) {
      main.innerHTML = `${flashHtml()}<div class="empty"><p>Такой записки нет в реестре.</p><a class="btn" href="#/list">Открыть реестр</a></div>`;
      return;
    }
    const addressees = store.memoAddressees(id);

    main.innerHTML = `
      ${flashHtml()}
      <header class="page-head">
        <div>
          <a class="back" href="#/list">Реестр</a>
          <h1 class="reg">${esc(memo.reg_number)}</h1>
        </div>
      </header>
      <div class="compose">
        <aside class="panel memo-meta">
          <dl class="details">
            <dt>Дата</dt><dd>${formatDate(memo.memo_date)}</dd>
            <dt>Тематика</dt><dd>${esc(memo.topic_name)}</dd>
            <dt>Кому</dt><dd>${addressees.map((a) => `<div>${esc(a.department)}<div class="sub">${esc(a.head_position)} ${esc(addresseeName(a))}</div></div>`).join('')}</dd>
            ${memo.signer_name ? `<dt>Подписант</dt><dd>${esc(shortName(memo.signer_name))}<div class="sub pre">${esc(memo.signer_position)}</div></dd>` : ''}
            <dt>Исполнитель</dt><dd>${esc(memo.executor_name)}${memo.executor_phone ? `<div class="sub">${esc(memo.executor_phone)}</div>` : ''}</dd>
            <dt>Создана</dt><dd>${formatDate(memo.created_at.slice(0, 10))} в ${esc(memo.created_at.slice(11, 16))}</dd>
            <dt>Документ</dt><dd id="file-status" class="sub">проверяем…</dd>
          </dl>
          <div class="stack">
            <button class="btn primary" id="download" hidden>Скачать Word</button>
            <button class="btn" id="regenerate">Сформировать Word заново</button>
            <button class="btn danger" id="delete">Удалить из реестра</button>
          </div>
        </aside>
        <div class="sheet-wrap">${sheetHtml(memoForSheet(memo), addressees)}</div>
      </div>`;

    document.getElementById('download').onclick = () => downloadMemo(id).catch(showError);
    document.getElementById('regenerate').onclick = (ev) => busy(ev.target, async () => {
      const path = await store.saveDocument(memo, memo.file_path);
      await store.write((s) => s.run('UPDATE memos SET file_path = ? WHERE id = ?', [path, id]));
      flash('Документ Word сформирован заново.');
    });
    document.getElementById('delete').onclick = (ev) => {
      if (!confirm(`Удалить ${memo.reg_number} из реестра? Документ Word в папке останется.`)) return;
      busy(ev.target, async () => {
        await store.write((s) => {
          s.run('DELETE FROM memo_addressees WHERE memo_id = ?', [id]);
          s.run('DELETE FROM memos WHERE id = ?', [id]);
        });
        flash(`${memo.reg_number} удалена из реестра. Документ Word остался в папке.`);
        location.hash = '#/list';
      });
    };

    const file = await store.readDocument(memo.file_path).catch(() => null);
    const status = document.getElementById('file-status');
    if (!status) return; // пользователь уже ушёл со страницы
    if (file) {
      status.textContent = memo.file_path;
      document.getElementById('download').hidden = false;
    } else {
      status.innerHTML = `<span class="error-text">Не найден в папке.</span> Нажмите «Сформировать Word заново».`;
    }
  }

  // ---------- Статистика ----------

  function niceStep(max) {
    const raw = Math.max(1, max) / 4;
    const pow = 10 ** Math.floor(Math.log10(raw));
    const step = [1, 2, 5, 10].map((k) => k * pow).find((s) => s >= raw);
    return Math.max(1, step);
  }

  function columnChart(values, labels, tips, highlight) {
    const W = 720; const H = 240;
    const L = 32; const R = 8; const T = 22; const B = 28;
    const step = niceStep(Math.max(...values));
    const top = Math.max(step, Math.ceil(Math.max(...values) / step) * step);
    const plotW = W - L - R; const plotH = H - T - B;
    const band = plotW / values.length;
    const colW = Math.min(24, band * 0.56);
    const y = (v) => T + plotH - (v / top) * plotH;
    const parts = [];

    for (let v = 0; v <= top; v += step) {
      parts.push(`<line class="grid" x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}"/>`);
      parts.push(`<text class="tick" x="${L - 8}" y="${y(v) + 4}" text-anchor="end">${v}</text>`);
    }
    values.forEach((v, i) => {
      const cx = L + band * i + band / 2;
      const x0 = cx - colW / 2; const x1 = cx + colW / 2;
      const base = y(0);
      if (v > 0) {
        const h = base - y(v); const r = Math.min(4, h);
        const t = y(v);
        parts.push(`<path class="col" d="M${x0} ${base}V${t + r}Q${x0} ${t} ${x0 + r} ${t}H${x1 - r}Q${x1} ${t} ${x1} ${t + r}V${base}Z"/>`);
        parts.push(`<text class="cap" x="${cx}" y="${t - 6}" text-anchor="middle">${v}</text>`);
      }
      parts.push(`<text class="tick ${i === highlight ? 'tick-now' : ''}" x="${cx}" y="${H - 8}" text-anchor="middle">${labels[i]}</text>`);
      parts.push(`<rect class="hit" x="${L + band * i}" y="${T}" width="${band}" height="${plotH}" tabindex="0" data-tip="${esc(tips[i])}"/>`);
    });
    return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Количество служебных записок по месяцам">${parts.join('')}</svg>`;
  }

  function barList(rows, total) {
    if (!rows.length) return '<p class="muted small">Нет данных.</p>';
    const max = Math.max(...rows.map((r) => r.n));
    return `<ol class="bars">${rows.map((r) => {
      const share = Math.round((r.n / total) * 100);
      return `<li class="bar-row" tabindex="0" data-tip="${esc(`${r.label}: ${r.n} (${share}% за год)`)}">
        <span class="bar-label">${esc(r.label)}</span>
        <span class="bar-track"><span class="bar" style="width: calc((100% - 3rem) * ${r.n / max})"></span><span class="bar-value">${r.n}</span></span>
      </li>`;
    }).join('')}</ol>`;
  }

  // Адресатов у записки может быть несколько, поэтому они считаются по memo_addressees
  const GROUP_QUERIES = {
    topic: 'SELECT topic_name AS label, COUNT(*) AS n FROM memos WHERE year = ? GROUP BY label',
    executor: 'SELECT executor_name AS label, COUNT(*) AS n FROM memos WHERE year = ? GROUP BY label',
    addressee: `SELECT a.department AS label, COUNT(*) AS n FROM memo_addressees a
      JOIN memos m ON m.id = a.memo_id WHERE m.year = ? GROUP BY label`,
  };

  function groupCounts(kind, year, limit = 7) {
    const rows = store.all(`${GROUP_QUERIES[kind]} ORDER BY n DESC, label`, [year]);
    if (rows.length <= limit + 1) return rows;
    const rest = rows.slice(limit).reduce((s, r) => s + r.n, 0);
    return [...rows.slice(0, limit), { label: `Остальные (${rows.length - limit})`, n: rest }];
  }

  function renderStats() {
    const now = new Date();
    const currentYear = now.getFullYear();
    const years = store.all('SELECT DISTINCT year FROM memos ORDER BY year DESC').map((r) => r.year);
    if (!years.includes(currentYear)) years.unshift(currentYear);
    if (!years.includes(statsYear)) statsYear = currentYear;
    const year = statsYear;
    const isCurrent = year === currentYear;

    const total = store.one('SELECT COUNT(*) AS n FROM memos WHERE year = ?', [year]).n;
    const allTime = store.one('SELECT COUNT(*) AS n FROM memos').n;
    const byMonth = Array(12).fill(0);
    for (const r of store.all(`SELECT CAST(substr(memo_date, 6, 2) AS INTEGER) AS m, COUNT(*) AS n
        FROM memos WHERE year = ? GROUP BY m`, [year])) byMonth[r.m - 1] = r.n;
    const distinct = (col) => store.one(`SELECT COUNT(DISTINCT ${col}) AS n FROM memos WHERE year = ?`, [year]).n;

    const monthTile = isCurrent
      ? [`В ${MONTHS_IN[now.getMonth()]}`, byMonth[now.getMonth()]]
      : ['В среднем за месяц', (total / 12).toLocaleString('ru', { maximumFractionDigits: 1 })];

    main.innerHTML = `
      ${flashHtml()}
      <header class="page-head">
        <h1>Статистика</h1>
        <label class="inline-select"><span class="sr">Год</span>
          <select id="stats-year">${years.map((y) => `<option ${y === year ? 'selected' : ''}>${y}</option>`).join('')}</select>
        </label>
      </header>
      ${total ? `
      <section class="stats-head panel">
        <div class="hero">
          <div class="hero-value">${total}</div>
          <div class="hero-label">служебных записок за ${year} год</div>
        </div>
        <dl class="tiles">
          <div><dt>${monthTile[0]}</dt><dd>${monthTile[1]}</dd></div>
          <div><dt>Тематик</dt><dd>${distinct('topic_name')}</dd></div>
          <div><dt>Исполнителей</dt><dd>${distinct('executor_name')}</dd></div>
          <div><dt>Всего в реестре</dt><dd>${allTime}</dd></div>
        </dl>
      </section>

      <section class="panel chart-panel">
        <header class="chart-head">
          <h2>По месяцам</h2>
          <button class="btn small ghost" id="toggle-table">${statsAsTable ? 'Показать график' : 'Показать таблицей'}</button>
        </header>
        ${statsAsTable ? `
          <table class="table compact"><thead><tr><th>Месяц</th><th class="r">СЗ</th></tr></thead>
          <tbody>${byMonth.map((n, i) => `<tr><td>${MONTHS_FULL[i]}</td><td class="r">${n}</td></tr>`).join('')}</tbody></table>`
          : columnChart(byMonth, MONTHS, byMonth.map((n, i) => `${MONTHS_FULL[i]} ${year}: ${n} СЗ`), isCurrent ? now.getMonth() : -1)}
      </section>

      <div class="stats-grid">
        <section class="panel"><h2>По тематикам</h2>${barList(groupCounts('topic', year), total)}</section>
        <section class="panel"><h2>По адресатам</h2>${barList(groupCounts('addressee', year), total)}</section>
        <section class="panel"><h2>По исполнителям</h2>${barList(groupCounts('executor', year), total)}</section>
      </div>`
      : `<div class="empty"><p>За ${year} год служебных записок нет.</p><a class="btn primary" href="#/new">Создать СЗ</a></div>`}`;

    document.getElementById('stats-year').onchange = (e) => { statsYear = Number(e.target.value); render(); };
    const toggle = document.getElementById('toggle-table');
    if (toggle) toggle.onclick = () => { statsAsTable = !statsAsTable; render(); };
  }

  // ---------- Справочники ----------

  function dirTabs(kind) {
    const tabs = Object.entries(DIRECTORIES).map(([k, d]) => [k, d.title]);
    tabs.push(['settings', 'Настройки']);
    return `<nav class="tabs" aria-label="Справочники">${tabs.map(([k, t]) =>
      `<a href="#/dirs/${k}" class="${k === kind ? 'active' : ''}" ${k === kind ? 'aria-current="page"' : ''}>${t}</a>`).join('')}</nav>`;
  }

  // Подписант один: пока галочка стоит у одного сотрудника, у остальных она заблокирована
  function signerCell(row, current) {
    const locked = current && current.id !== row.id;
    const title = locked ? `Подписант уже назначен: ${shortName(current.full_name)}. Сначала снимите галочку у него.` : '';
    return `<td class="c"><input type="checkbox" class="signer-box" data-signer="${row.id}"
      ${row.is_signer ? 'checked' : ''} ${locked ? `disabled title="${esc(title)}"` : ''}
      aria-label="Подписант: ${esc(row.full_name)}"></td>`;
  }

  async function setSigner(id, on) {
    const hash = location.hash;
    try {
      await store.write((s) => {
        if (on) {
          const current = s.signer;
          if (current && current.id !== id) {
            throw new Error(`Подписант уже назначен: ${current.full_name}. Сначала снимите галочку у него.`);
          }
          const emp = s.one('SELECT * FROM employees WHERE id = ?', [id]);
          if (!emp) throw new Error('Сотрудник удалён. Обновите страницу.');
          if (!emp.position) {
            throw new Error(`У сотрудника ${emp.full_name} не указана должность: она печатается рядом с подписью. Сначала заполните её.`);
          }
        }
        s.run('UPDATE employees SET is_signer = ? WHERE id = ?', [on ? 1 : 0, id]);
      });
      flash(on ? 'Подписант назначен. Он будет подставляться в новые записки.' : 'Подписант снят.');
    } catch (e) {
      console.error(e);
      flash(e.message, 'error');
    }
    if (location.hash === hash) render();
  }

  function renderDirs(kind) {
    if (kind === 'settings') return renderSettings();
    const spec = DIRECTORIES[kind];
    const rows = store.all(`SELECT * FROM ${kind} ORDER BY ${spec.order}`);
    const editRow = editId != null ? rows.find((r) => r.id === editId) : null;
    if (!editRow) editId = null;

    main.innerHTML = `
      ${flashHtml()}
      <header class="page-head"><h1>Справочники</h1></header>
      ${dirTabs(kind)}
      <form class="panel form dir-form ${editRow ? 'is-editing' : ''}" id="dir-form">
        <h2>${editRow ? 'Изменить запись' : 'Добавить запись'}</h2>
        <div class="fields">
          ${spec.fields.map(([name, label, required, multiline]) => `
            <label class="field"><span>${label}${required ? '' : ' <i>необязательно</i>'}</span>
              ${multiline
                ? `<textarea name="${name}" rows="2" ${required ? 'required' : ''} placeholder="Enter - перенос строки">${esc(editRow ? editRow[name] : '')}</textarea>`
                : `<input type="text" name="${name}" value="${esc(editRow ? editRow[name] : '')}" ${required ? 'required' : ''}>`}</label>`).join('')}
          <div class="field-actions">
            <button type="submit" class="btn primary">${editRow ? 'Сохранить' : 'Добавить'}</button>
            ${editRow ? '<button type="button" class="btn ghost" id="cancel">Отмена</button>' : ''}
          </div>
        </div>
      </form>
      ${rows.length ? `
        <div class="panel table-wrap"><table class="table">
          <thead><tr>${spec.fields.map(([, label]) => `<th>${label}</th>`).join('')}${spec.signer ? '<th class="c">Подписант</th>' : ''}<th><span class="sr">Действия</span></th></tr></thead>
          <tbody>${rows.map((r) => `
            <tr class="${editRow && editRow.id === r.id ? 'editing' : ''}">
              ${spec.fields.map(([name, , , multiline]) => `<td class="${multiline ? 'pre' : ''}">${esc(r[name]) || '<span class="muted">—</span>'}</td>`).join('')}
              ${spec.signer ? signerCell(r, rows.find((x) => x.is_signer)) : ''}
              <td class="row-actions">
                <button class="btn small ghost" data-edit="${r.id}">Изменить</button>
                <button class="btn small ghost danger" data-delete="${r.id}">Удалить</button>
              </td>
            </tr>`).join('')}
          </tbody></table></div>
        <p class="muted small">Записей: ${rows.length}</p>`
        : `<div class="empty"><p>Справочник «${spec.title}» пока пуст. Добавьте первую запись в форме выше.</p></div>`}`;

    const form = document.getElementById('dir-form');
    if (editRow) form.querySelector('input').focus();
    form.onsubmit = (ev) => {
      ev.preventDefault();
      const values = {};
      for (const [name, label, required, multiline] of spec.fields) {
        const raw = form.elements[name].value || '';
        values[name] = multiline ? splitLines(raw).join('\n') : raw.trim();
        if (required && !values[name]) {
          flash(`Заполните поле «${label}».`, 'error');
          showFlashes();
          return;
        }
      }
      const cols = Object.keys(values);
      const id = editRow ? editRow.id : null;
      busy(form.querySelector('button[type=submit]'), async () => {
        try {
          await store.write((s) => {
            if (id != null) {
              s.run(`UPDATE ${kind} SET ${cols.map((c) => `${c} = ?`).join(', ')} WHERE id = ?`,
                [...cols.map((c) => values[c]), id]);
            } else {
              s.run(`INSERT INTO ${kind} (${cols.join(', ')}) VALUES (${cols.map(() => '?').join(', ')})`,
                cols.map((c) => values[c]));
            }
          });
          flash(id != null ? 'Изменения сохранены.' : 'Запись добавлена.');
          editId = null;
        } catch (e) {
          throw /UNIQUE/.test(e.message) ? new Error('Такая запись уже есть в справочнике.') : e;
        }
      });
    };
    const cancel = document.getElementById('cancel');
    if (cancel) cancel.onclick = () => { editId = null; render(); };
    for (const b of main.querySelectorAll('[data-edit]')) {
      b.onclick = () => { editId = Number(b.dataset.edit); render(); };
    }
    for (const box of main.querySelectorAll('[data-signer]')) {
      box.onchange = () => setSigner(Number(box.dataset.signer), box.checked);
    }
    for (const b of main.querySelectorAll('[data-delete]')) {
      b.onclick = () => {
        if (!confirm('Удалить запись? В уже созданных записках данные сохранятся.')) return;
        busy(b, async () => {
          await store.write((s) => s.run(`DELETE FROM ${kind} WHERE id = ?`, [Number(b.dataset.delete)]));
          flash('Запись удалена.');
        });
      };
    }
  }

  function renderSettings() {
    const year = new Date().getFullYear();
    main.innerHTML = `
      ${flashHtml()}
      <header class="page-head"><h1>Справочники</h1></header>
      ${dirTabs('settings')}
      <form class="panel form settings" id="settings-form">
        <label class="field"><span>Наименование организации</span>
          <textarea name="org_name" rows="3" placeholder="Каждая строка шапки с новой строки">${esc(store.orgName)}</textarea>
          <small>Печатается в шапке по центру полужирным. Переносы строк сохраняются как введены.</small></label>
        <label class="field field-short"><span>Город</span>
          <input type="text" name="city" value="${esc(store.city)}" placeholder="Например, Москва">
          <small>Печатается под заголовком «СЛУЖЕБНАЯ ЗАПИСКА».</small></label>
        <label class="field field-short"><span>Префикс номера</span>
          <input type="text" name="number_prefix" value="${esc(store.numberPrefix)}" required maxlength="20">
          <small id="prefix-example"></small></label>
        <label class="field"><span>Папка для документов Word</span>
          <input type="text" name="output_dir" value="${esc(store.outputDir)}" required>
          <small>Внутри общей папки с SZ.html, например «${esc(DEFAULT_OUTPUT_DIR)}» или «Документы/СЗ». Подпапки по годам создаются сами.</small></label>
        <div class="form-foot"><button type="submit" class="btn primary">Сохранить настройки</button></div>
      </form>`;

    const form = document.getElementById('settings-form');
    const example = () => {
      const prefix = form.elements.number_prefix.value.trim() || DEFAULT_NUMBER_PREFIX;
      document.getElementById('prefix-example').textContent =
        `Номер выглядит так: ${formatRegNumber(prefix, 1, year)}. Уже созданные записки сохраняют свои номера.`;
    };
    form.elements.number_prefix.addEventListener('input', example);
    example();

    form.onsubmit = (ev) => {
      ev.preventDefault();
      const orgName = form.elements.org_name.value.split(/\r?\n/).map((l) => l.trim()).filter(Boolean).join('\n');
      const prefix = form.elements.number_prefix.value.trim();
      const outputDir = form.elements.output_dir.value.trim().replace(/\\/g, '/');
      const badChars = /[\\/:*?"<>|]/;
      if (!prefix || badChars.test(prefix)) {
        flash('Префикс не может быть пустым и не может содержать символы \\ / : * ? " < > |', 'error');
        showFlashes();
        return;
      }
      if (!outputDir || outputDir.split('/').some((p) => p === '..' || /[:*?"<>|]/.test(p))) {
        flash('Укажите папку внутри общей папки, без «..» и символов : * ? " < > |', 'error');
        showFlashes();
        return;
      }
      busy(form.querySelector('button'), async () => {
        await store.write((s) => {
          s.setSetting('org_name', orgName);
          s.setSetting('city', form.elements.city.value.trim());
          s.setSetting('number_prefix', prefix);
          s.setSetting('output_dir', outputDir);
        });
        flash('Настройки сохранены.');
      });
    };
  }

  // ---------- Запуск ----------

  window.addEventListener('hashchange', async () => {
    editId = null;
    try {
      if (store) await store.refresh(false);
    } catch (e) {
      console.error(e);
    }
    render();
  });

  // Подтягиваем изменения других пользователей при возврате во вкладку
  window.addEventListener('focus', async () => {
    if (!store) return;
    try {
      const { page } = route();
      if (await store.refresh(false) && ['list', 'memo', 'stats'].includes(page)) render();
    } catch (e) {
      console.error(e);
    }
  });

  async function start() {
    if (!window.showDirectoryPicker) {
      main.innerHTML = `<section class="connect"><h1>Этот браузер не подходит</h1>
        <p>Откройте SZ.html в Google Chrome, Microsoft Edge или Яндекс Браузере.</p></section>`;
      return;
    }
    try {
      const bin = Uint8Array.from(atob(window.SQL_WASM_BASE64), (c) => c.charCodeAt(0));
      window.SZ_SQL = await window.initSqlJs({ wasmBinary: bin });
    } catch (e) {
      main.innerHTML = `<div class="flash flash-error">Не удалось загрузить модуль базы данных: ${esc(e.message)}</div>`;
      return;
    }
    let handle = await savedHandle();
    // SZ.html открыт из другой папки, чем запомненная: запомненную не используем
    if (handle && !isPageFolder(handle)) handle = null;
    if (handle) {
      try {
        if (await handle.queryPermission({ mode: 'readwrite' }) === 'granted') {
          await connect(handle);
          return;
        }
      } catch {
        // Сохранённая папка недоступна - спросим заново
      }
    }
    renderConnect(handle);
  }

  start();
})();
