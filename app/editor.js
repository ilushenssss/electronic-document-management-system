// Поле «Текст» с панелью оформления: полужирный, курсив, подчёркивание, размер шрифта,
// выравнивание, отступ первой строки и перенос строки внутри абзаца.
//
// Текст хранится как список абзацев (формат SZDocx.normalizeRich), поле на экране -
// его отображение. Кнопки, Enter, Shift+Enter и вставка меняют список и перерисовывают
// поле; набор и удаление символов делает браузер, после чего список читается из поля.
// Поэтому в записку не попадает чужая разметка, а отмена (Ctrl+Z) работает для всех
// действий, в том числе для размера шрифта, которого у браузера нет в пунктах.
(function (root) {
  'use strict';

  const { RICH_SIZES, normalizeRich, richPlainText } = root.SZDocx;

  const DEFAULT_SIZE = 14;
  const FLAGS = [
    ['bold', 'Полужирный (Ctrl+B)', 'KeyB', '<b>Ж</b>'],
    ['italic', 'Курсив (Ctrl+I)', 'KeyI', '<i>К</i>'],
    ['underline', 'Подчёркнутый (Ctrl+U)', 'KeyU', '<u>Ч</u>'],
  ];
  // Ctrl+L, Ctrl+E, Ctrl+R, Ctrl+J в браузере заняты (адресная строка, обновление страницы),
  // поэтому выравнивание - только кнопками
  const ALIGNS = [['left', 'По левому краю'], ['center', 'По центру'], ['right', 'По правому краю'],
    ['both', 'По ширине']];
  const HISTORY_LIMIT = 200;
  const TYPING_PAUSE_MS = 800; // набор без пауз отменяется одним шагом

  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

  // ---------- Значки ----------

  const svg = (body) => `<svg viewBox="0 0 18 18" aria-hidden="true">${body}</svg>`;
  const lines = (rows) => rows.map(([x1, x2], i) => `<path d="M${x1} ${4 + i * 3.4}H${x2}"/>`).join('');
  const ICONS = {
    left: svg(lines([[3, 15], [3, 11], [3, 15], [3, 9]])),
    center: svg(lines([[3, 15], [5, 13], [3, 15], [6, 12]])),
    right: svg(lines([[3, 15], [7, 15], [3, 15], [9, 15]])),
    both: svg(lines([[3, 15], [3, 15], [3, 15], [3, 15]])),
    indent: svg('<path d="M8 4H15M3 7.4H15M3 10.8H15M3 14.2H15M3 4H5.5"/>'),
    break: svg('<path d="M14 3.5V10H4.5M4.5 10L7.5 7M4.5 10L7.5 13"/>'),
    clear: svg('<text x="1.5" y="12" font-size="11" font-weight="700" stroke="none" fill="currentColor">A</text>' +
      '<path class="x" d="M10.5 10.5L16 16M16 10.5L10.5 16"/>'),
  };

  // ---------- Модель: абзацы и фрагменты ----------

  const FORMAT_KEYS = ['bold', 'italic', 'underline', 'size'];

  function formatOf(run) {
    const f = {};
    for (const k of FORMAT_KEYS) if (run[k]) f[k] = run[k];
    if (f.size === DEFAULT_SIZE) delete f.size;
    return f;
  }

  const sameFormat = (a, b) => FORMAT_KEYS.every((k) => (a[k] || 0) === (b[k] || 0));
  const paraLength = (p) => p.runs.reduce((n, r) => n + r.text.length, 0);
  const paraText = (p) => p.runs.map((r) => r.text).join('');
  const cmp = (a, b) => a.p - b.p || a.o - b.o;
  const clone = (model) => model.map((p) => ({ align: p.align, indent: p.indent, runs: p.runs.map((r) => ({ ...r })) }));

  function mergeRuns(runs) {
    const out = [];
    for (const r of runs) {
      if (!r.text) continue;
      const last = out[out.length - 1];
      if (last && sameFormat(last, r)) last.text += r.text;
      else out.push({ ...formatOf(r), text: r.text });
    }
    return out;
  }

  // Фрагменты абзаца между позициями a и b (копии)
  function sliceRuns(runs, a, b) {
    const out = [];
    let pos = 0;
    for (const r of runs) {
      const end = pos + r.text.length;
      const from = Math.max(a, pos);
      const to = Math.min(b, end);
      if (from < to) out.push({ ...r, text: r.text.slice(from - pos, to - pos) });
      pos = end;
    }
    return out;
  }

  // Оформление символа перед позицией (в начале абзаца - после неё)
  function formatAt(model, pos) {
    const runs = model[pos.p] ? model[pos.p].runs : [];
    let at = Math.max(0, pos.o - 1);
    for (const r of runs) {
      if (at < r.text.length) return formatOf(r);
      at -= r.text.length;
    }
    return runs.length ? formatOf(runs[runs.length - 1]) : {};
  }

  // Применяет fn к каждому фрагменту выделения
  function mapRange(model, start, end, fn) {
    const out = clone(model);
    for (let i = start.p; i <= end.p; i++) {
      const p = out[i];
      const a = i === start.p ? start.o : 0;
      const b = i === end.p ? end.o : paraLength(p);
      p.runs = mergeRuns([...sliceRuns(p.runs, 0, a), ...sliceRuns(p.runs, a, b).map(fn),
        ...sliceRuns(p.runs, b, Infinity)]);
    }
    return out;
  }

  function rangeRuns(model, start, end) {
    const runs = [];
    for (let i = start.p; i <= end.p; i++) {
      const p = model[i];
      runs.push(...sliceRuns(p.runs, i === start.p ? start.o : 0, i === end.p ? end.o : Infinity));
    }
    return runs;
  }

  function deleteRange(model, start, end) {
    const out = clone(model);
    const first = out[start.p];
    const last = out[end.p];
    const joined = { align: first.align, indent: first.indent,
      runs: mergeRuns([...sliceRuns(first.runs, 0, start.o), ...sliceRuns(last.runs, end.o, Infinity)]) };
    out.splice(start.p, end.p - start.p + 1, joined);
    return out;
  }

  // Вставляет абзацы в позицию: первый дописывается к текущему абзацу, последний -
  // перед его продолжением, остальные получают выравнивание и отступ текущего
  function insertParagraphs(model, pos, paragraphs) {
    const out = clone(model);
    const cur = out[pos.p];
    const head = sliceRuns(cur.runs, 0, pos.o);
    const tail = sliceRuns(cur.runs, pos.o, Infinity);
    const made = paragraphs.map((p) => ({ align: cur.align, indent: cur.indent, runs: p.runs }));
    made[0].runs = [...head, ...made[0].runs];
    const lastPara = made[made.length - 1];
    const caret = { p: pos.p + made.length - 1, o: lastPara.runs.reduce((n, r) => n + r.text.length, 0) };
    lastPara.runs = [...lastPara.runs, ...tail];
    for (const p of made) p.runs = mergeRuns(p.runs);
    out.splice(pos.p, 1, ...made);
    return { model: out, caret };
  }

  const emptyModel = () => [{ align: 'both', indent: true, runs: [] }];

  // ---------- Чтение абзацев из HTML ----------

  const BLOCK_TAGS = new Set(['P', 'DIV', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE',
    'TR', 'TABLE', 'UL', 'OL', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'DD', 'DT']);
  const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'TEMPLATE', 'HEAD', 'TITLE', 'META', 'LINK', 'IMG', 'svg']);

  function sizeFromStyle(el, basePx) {
    if (el.dataset && el.dataset.size) return Number(el.dataset.size);
    const fs = el.style && el.style.fontSize;
    if (!fs) return null;
    const v = parseFloat(fs);
    if (fs.endsWith('em')) return Math.round(v * DEFAULT_SIZE);
    // Размер относительно основного шрифта поля, который соответствует 14 pt документа
    if (fs.endsWith('pt')) return Math.round(((v * 4) / 3 / basePx) * DEFAULT_SIZE);
    if (fs.endsWith('px')) return Math.round((v / basePx) * DEFAULT_SIZE);
    return null;
  }

  function childFormat(el, fmt, opts) {
    const f = { ...fmt };
    const tag = el.tagName;
    if (tag === 'B' || tag === 'STRONG') f.bold = true;
    if (tag === 'I' || tag === 'EM') f.italic = true;
    if (tag === 'U' || tag === 'INS') f.underline = true;
    const st = el.style;
    if (st) {
      const w = st.fontWeight;
      if (w === 'bold' || w === 'bolder' || Number(w) >= 600) f.bold = true;
      else if (w === 'normal' || (w && Number(w) < 600)) delete f.bold;
      if (st.fontStyle === 'italic' || st.fontStyle === 'oblique') f.italic = true;
      else if (st.fontStyle === 'normal') delete f.italic;
      const deco = st.textDecorationLine || st.textDecoration || '';
      if (deco.includes('underline')) f.underline = true;
      else if (deco === 'none') delete f.underline;
    }
    if (opts.keepLayout) {
      const size = sizeFromStyle(el, opts.basePx);
      if (size) f.size = size;
    }
    return f;
  }

  function alignOf(el) {
    const a = el && el.dataset && el.dataset.align;
    if (a) return a;
    const ta = el && el.style ? el.style.textAlign : '';
    return { justify: 'both', center: 'center', right: 'right', left: 'left', start: 'left', end: 'right' }[ta] || 'both';
  }

  function indentOf(el) {
    if (el && el.dataset && el.dataset.indent) return el.dataset.indent === '1';
    return !(el && el.style && el.style.textIndent && parseFloat(el.style.textIndent) === 0);
  }

  // Абзацы из содержимого элемента. points - точки DOM {node, offset}, для которых нужно
  // узнать позицию {p, o} (так выделение переносится в модель).
  // keepLayout - брать размер шрифта; collapse - схлопывать пробелы, как в обычном HTML.
  function readParagraphs(rootEl, { keepLayout = true, collapse = false, basePx = 16, points = [] } = {}) {
    const paras = [];
    const found = points.map(() => null);
    const pendingStart = []; // точки «перед ещё не начатым абзацем»
    let cur = null;
    const opts = { keepLayout, basePx };

    function start(el) {
      cur = { align: alignOf(el), indent: indentOf(el), runs: [] };
      paras.push(cur);
      for (const i of pendingStart.splice(0)) found[i] = { p: paras.length - 1, o: 0 };
    }
    const curLength = () => paraLength(cur);
    function mark(node, offset) {
      points.forEach((pt, i) => {
        if (found[i] || pt.node !== node || pt.offset !== offset) return;
        if (cur) found[i] = { p: paras.length - 1, o: curLength() };
        else pendingStart.push(i);
      });
    }
    function push(text, fmt) {
      if (!cur) start(null);
      cur.runs.push({ ...fmt, text });
    }
    function endBlock() {
      // Последний <br> в абзаце браузер ставит, чтобы абзац не схлопнулся
      const last = cur && cur.runs[cur.runs.length - 1];
      if (last && last.text.endsWith('\n')) last.text = last.text.slice(0, -1);
      cur = null;
    }
    function walk(node, fmt) {
      const kids = node.childNodes;
      for (let k = 0; k < kids.length; k++) {
        mark(node, k);
        const child = kids[k];
        if (child.nodeType === 3) {
          let text = child.data.replace(/ /g, ' ');
          if (collapse) text = text.replace(/[\t\n\r ]+/g, ' ');
          points.forEach((pt, i) => {
            if (!found[i] && pt.node === child) {
              if (!cur) start(null);
              found[i] = { p: paras.length - 1, o: curLength() + Math.min(pt.offset, text.length) };
            }
          });
          if (text) push(text, fmt);
        } else if (child.nodeType === 1 && !SKIP_TAGS.has(child.tagName)) {
          if (child.tagName === 'BR') {
            push('\n', fmt);
            continue;
          }
          const f = childFormat(child, fmt, opts);
          if (BLOCK_TAGS.has(child.tagName)) {
            if (cur) endBlock();
            start(child);
            walk(child, f);
            mark(child, child.childNodes.length);
            endBlock();
          } else {
            walk(child, f);
            mark(child, child.childNodes.length);
          }
        }
      }
      mark(node, kids.length);
    }
    walk(rootEl, {});
    if (cur) endBlock();
    if (!paras.length) paras.push(...emptyModel());
    for (const p of paras) {
      p.runs = mergeRuns(p.runs);
      if (!['left', 'center', 'right', 'both'].includes(p.align)) p.align = 'both';
      for (const r of p.runs) if (r.size && !RICH_SIZES.includes(r.size)) delete r.size;
      p.runs = mergeRuns(p.runs);
    }
    // Точки за концом текста - в конец последнего абзаца
    const clampPos = (pos) => {
      if (!pos) {
        const p = paras.length - 1;
        return { p, o: paraLength(paras[p]) };
      }
      const p = Math.min(pos.p, paras.length - 1);
      return { p, o: Math.min(pos.o, paraLength(paras[p])) };
    };
    return { paragraphs: paras, points: found.map(clampPos) };
  }

  // ---------- Отображение модели в поле ----------

  function runHtml(r) {
    let h = esc(r.text).replace(/\n/g, '<br>');
    if (r.bold) h = `<b>${h}</b>`;
    if (r.italic) h = `<i>${h}</i>`;
    if (r.underline) h = `<u>${h}</u>`;
    if (r.size) h = `<span data-size="${r.size}" style="font-size:${(r.size / DEFAULT_SIZE).toFixed(4)}em">${h}</span>`;
    return h;
  }

  function paraHtml(p) {
    const text = paraText(p);
    // Пустой абзац и абзац с переносом в конце без <br> на конце схлопнулись бы
    const tail = !text || text.endsWith('\n') ? '<br>' : '';
    return `<p data-align="${p.align}" data-indent="${p.indent ? 1 : 0}">${p.runs.map(runHtml).join('')}${tail}</p>`;
  }

  // ---------- Поле ----------

  function create(container, { placeholder = '', onChange = () => {} } = {}) {
    container.innerHTML = `
      <div class="rich">
        <div class="rich-bar" role="toolbar" aria-label="Оформление текста">
          ${FLAGS.map(([name, title, , label]) =>
            `<button type="button" class="rich-btn" data-flag="${name}" title="${title}" aria-label="${title}" aria-pressed="false">${label}</button>`).join('')}
          <select class="rich-size" title="Размер шрифта" aria-label="Размер шрифта">
            ${RICH_SIZES.map((s) => `<option value="${s}" ${s === DEFAULT_SIZE ? 'selected' : ''}>${s}</option>`).join('')}
          </select>
          <span class="rich-sep"></span>
          ${ALIGNS.map(([name, title]) =>
            `<button type="button" class="rich-btn" data-align="${name}" title="${title}" aria-label="${title}" aria-pressed="false">${ICONS[name]}</button>`).join('')}
          <span class="rich-sep"></span>
          <button type="button" class="rich-btn" data-cmd="indent" title="Отступ первой строки" aria-label="Отступ первой строки" aria-pressed="false">${ICONS.indent}</button>
          <button type="button" class="rich-btn" data-cmd="break" title="Перенос строки без нового абзаца (Shift+Enter)" aria-label="Перенос строки">${ICONS.break}</button>
          <button type="button" class="rich-btn" data-cmd="clear" title="Убрать оформление: шрифт 14, по ширине, с отступом" aria-label="Убрать оформление">${ICONS.clear}</button>
        </div>
        <div class="rich-edit" contenteditable="true" role="textbox" aria-multiline="true" spellcheck="true"
          aria-label="${esc(placeholder || 'Текст')}" data-placeholder="${esc(placeholder)}"></div>
      </div>`;
    const bar = container.querySelector('.rich-bar');
    const edit = container.querySelector('.rich-edit');
    const sizeSelect = bar.querySelector('.rich-size');

    let model = emptyModel();
    let lastSel = { start: { p: 0, o: 0 }, end: { p: 0, o: 0 } };
    let pending = null; // оформление для следующего набранного текста (кнопка нажата без выделения)
    let pendingAt = null;
    const undo = [];
    const redo = [];
    let typing = false;
    let typingTimer = null;
    let ready = false; // onChange - только после создания поля

    const basePx = () => parseFloat(getComputedStyle(edit).fontSize) || 16;

    function selectionInside() {
      const s = document.getSelection();
      return s && s.rangeCount && edit.contains(s.anchorNode) && edit.contains(s.focusNode) ? s : null;
    }

    // Модель и выделение по текущему содержимому поля
    function read() {
      const s = selectionInside();
      const points = s ? [{ node: s.anchorNode, offset: s.anchorOffset }, { node: s.focusNode, offset: s.focusOffset }] : [];
      const res = readParagraphs(edit, { basePx: basePx(), points });
      model = res.paragraphs;
      if (s) {
        const [a, b] = res.points;
        lastSel = cmp(a, b) <= 0 ? { start: a, end: b } : { start: b, end: a };
      } else {
        const clamp = (pos) => {
          const p = Math.min(pos.p, model.length - 1);
          return { p, o: Math.min(pos.o, paraLength(model[p])) };
        };
        lastSel = { start: clamp(lastSel.start), end: clamp(lastSel.end) };
      }
      return { model, start: lastSel.start, end: lastSel.end };
    }

    function domPoint(pos) {
      const pEl = edit.children[pos.p];
      let rest = pos.o;
      const walker = document.createTreeWalker(pEl, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      let node;
      while ((node = walker.nextNode())) {
        if (node.nodeType === 3) {
          if (rest <= node.data.length) return [node, rest];
          rest -= node.data.length;
        } else if (node.tagName === 'BR') {
          if (rest === 0) return [node.parentNode, Array.prototype.indexOf.call(node.parentNode.childNodes, node)];
          rest -= 1;
        }
      }
      return [pEl, pEl.childNodes.length];
    }

    function render(sel) {
      edit.innerHTML = model.map(paraHtml).join('');
      edit.classList.toggle('is-empty', model.length === 1 && !paraLength(model[0]));
      if (sel) {
        lastSel = sel;
        if (document.activeElement === edit || edit.contains(document.activeElement)) placeSelection(sel);
      }
    }

    function placeSelection(sel) {
      const range = document.createRange();
      range.setStart(...domPoint(sel.start));
      range.setEnd(...domPoint(sel.end));
      const s = document.getSelection();
      s.removeAllRanges();
      s.addRange(range);
    }

    function changed() {
      edit.classList.toggle('is-empty', model.length === 1 && !paraLength(model[0]));
      sync();
      if (ready) onChange();
    }

    // ---------- История ----------

    const snapshot = () => ({ model: clone(model), sel: { start: { ...lastSel.start }, end: { ...lastSel.end } } });

    function remember() {
      undo.push(snapshot());
      if (undo.length > HISTORY_LIMIT) undo.shift();
      redo.length = 0;
    }

    function endTyping() {
      typing = false;
      clearTimeout(typingTimer);
    }

    function restore(state) {
      model = clone(state.model);
      edit.focus();
      render(state.sel);
      pending = null;
      changed();
    }

    function doUndo() {
      read();
      endTyping();
      if (!undo.length) return;
      redo.push(snapshot());
      restore(undo.pop());
    }

    function doRedo() {
      read();
      endTyping();
      if (!redo.length) return;
      undo.push(snapshot());
      restore(redo.pop());
    }

    // Действие над моделью: fn(model, start, end) -> { model, sel }
    function apply(fn) {
      const { start, end } = read();
      endTyping();
      remember();
      const res = fn(model, start, end);
      model = res.model;
      edit.focus();
      render(res.sel);
      changed();
    }

    // ---------- Оформление ----------

    function currentFormat() {
      const { start, end } = lastSel;
      if (pending && cmp(start, end) === 0) return pending;
      if (cmp(start, end) === 0) return formatAt(model, start);
      return formatAt(model, { p: start.p, o: start.o + 1 });
    }

    function toggleFlag(flag) {
      const { start, end } = read();
      if (cmp(start, end) === 0) {
        // Без выделения: оформление для следующего набранного текста
        const f = { ...currentFormat() };
        if (f[flag]) delete f[flag];
        else f[flag] = true;
        pending = f;
        pendingAt = start;
        edit.focus();
        sync();
        return;
      }
      const all = rangeRuns(model, start, end).every((r) => r[flag] || !r.text.trim());
      apply((m, s, e) => ({
        model: mapRange(m, s, e, (r) => {
          const n = { ...r };
          if (all) delete n[flag];
          else n[flag] = true;
          return n;
        }),
        sel: { start: s, end: e },
      }));
    }

    function setSize(size) {
      const { start, end } = read();
      if (cmp(start, end) === 0) {
        pending = { ...currentFormat(), size };
        if (size === DEFAULT_SIZE) delete pending.size;
        pendingAt = start;
        edit.focus(); // фокус был у списка размеров: возвращаем курсор на место
        placeSelection(lastSel);
        sync();
        return;
      }
      apply((m, s, e) => ({
        model: mapRange(m, s, e, (r) => {
          const n = { ...r, size };
          if (size === DEFAULT_SIZE) delete n.size;
          return n;
        }),
        sel: { start: s, end: e },
      }));
    }

    function setParagraphs(fn) {
      apply((m, s, e) => {
        const out = clone(m);
        for (let i = s.p; i <= e.p; i++) fn(out[i]);
        return { model: out, sel: { start: s, end: e } };
      });
    }

    function insertParas(paragraphs, fmt) {
      apply((m, s, e) => {
        const base = cmp(s, e) ? deleteRange(m, s, e) : m;
        const withFmt = paragraphs.map((p) => ({ runs: p.runs.map((r) => ({ ...fmt, ...r })) }));
        const res = insertParagraphs(base, s, withFmt);
        return { model: res.model, sel: { start: res.caret, end: res.caret } };
      });
      pending = null;
    }

    function clearFormat() {
      apply((m, s, e) => {
        const out = cmp(s, e) ? mapRange(m, s, e, (r) => ({ text: r.text })) : clone(m);
        for (let i = s.p; i <= e.p; i++) Object.assign(out[i], { align: 'both', indent: true });
        return { model: out, sel: { start: s, end: e } };
      });
      pending = null;
    }

    // Кнопки показывают оформление текста под курсором
    function sync() {
      const f = currentFormat();
      for (const b of bar.querySelectorAll('[data-flag]')) {
        b.setAttribute('aria-pressed', f[b.dataset.flag] ? 'true' : 'false');
      }
      sizeSelect.value = String(f.size || DEFAULT_SIZE);
      const p = model[lastSel.start.p] || model[0];
      for (const b of bar.querySelectorAll('[data-align]')) {
        b.setAttribute('aria-pressed', b.dataset.align === p.align ? 'true' : 'false');
      }
      bar.querySelector('[data-cmd=indent]').setAttribute('aria-pressed', p.indent ? 'true' : 'false');
    }

    // ---------- События ----------

    // Кнопки не забирают фокус и выделение у поля
    bar.addEventListener('mousedown', (e) => {
      if (e.target.closest('button')) e.preventDefault();
    });
    bar.addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.flag) toggleFlag(b.dataset.flag);
      else if (b.dataset.align) setParagraphs((p) => { p.align = b.dataset.align; });
      else if (b.dataset.cmd === 'indent') {
        const on = !(model[lastSel.start.p] || model[0]).indent;
        setParagraphs((p) => { p.indent = on; });
      } else if (b.dataset.cmd === 'break') insertParas([{ runs: [{ text: '\n' }] }], currentFormat());
      else if (b.dataset.cmd === 'clear') clearFormat();
    });
    sizeSelect.addEventListener('change', () => setSize(Number(sizeSelect.value)));

    edit.addEventListener('keydown', (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      const flag = FLAGS.find((f) => f[2] === e.code);
      if (flag && !e.shiftKey) {
        e.preventDefault();
        toggleFlag(flag[0]);
      } else if (e.code === 'KeyZ' && !e.shiftKey) {
        e.preventDefault();
        doUndo();
      } else if (e.code === 'KeyY' || (e.code === 'KeyZ' && e.shiftKey)) {
        e.preventDefault();
        doRedo();
      }
    });

    edit.addEventListener('beforeinput', (e) => {
      const t = e.inputType;
      if (t === 'historyUndo' || t === 'historyRedo') {
        e.preventDefault();
        if (t === 'historyUndo') doUndo();
        else doRedo();
        return;
      }
      if (t === 'insertParagraph') {
        e.preventDefault();
        insertParas([{ runs: [] }, { runs: [] }], {});
        return;
      }
      if (t === 'insertLineBreak') {
        e.preventDefault();
        insertParas([{ runs: [{ text: '\n' }] }], currentFormat());
        return;
      }
      if (t.startsWith('format')) { // форматирование средствами браузера не используем
        e.preventDefault();
        return;
      }
      if (pending && (t === 'insertText' || t === 'insertReplacementText') && e.data) {
        // Дальше набор продолжается в том же оформлении: оно теперь у символа перед курсором
        e.preventDefault();
        insertParas([{ runs: [{ text: e.data }] }], pending);
        return;
      }
      if (!typing) {
        read();
        remember();
        typing = true;
      }
      clearTimeout(typingTimer);
      typingTimer = setTimeout(endTyping, TYPING_PAUSE_MS);
    });

    // Переход курсора в другое место завершает шаг отмены
    edit.addEventListener('mousedown', endTyping);
    edit.addEventListener('keydown', (e) => {
      if (/^(Arrow|Home|End|Page)/.test(e.key)) endTyping();
    });

    edit.addEventListener('input', () => {
      read();
      pending = null;
      changed();
    });

    // Вставка: из других программ - только текст, полужирный, курсив и подчёркивание
    edit.addEventListener('paste', (e) => {
      e.preventDefault();
      const dt = e.clipboardData;
      const html = dt.getData('text/html');
      const text = dt.getData('text/plain');
      let paragraphs;
      if (html) {
        const doc = new DOMParser().parseFromString(html, 'text/html');
        // Скопированное из этого же поля вставляется с размером шрифта
        const own = !!doc.querySelector('[data-size], [data-indent]');
        paragraphs = readParagraphs(doc.body, { keepLayout: own, collapse: true, basePx: basePx() }).paragraphs
          .filter((p) => paraText(p).trim());
        if (paragraphs.length) {
          const first = paragraphs[0].runs[0];
          first.text = first.text.replace(/^ +/, '');
          const lastRuns = paragraphs[paragraphs.length - 1].runs;
          lastRuns[lastRuns.length - 1].text = lastRuns[lastRuns.length - 1].text.replace(/ +$/, '');
        }
      } else if (text) {
        paragraphs = text.replace(/\r\n?/g, '\n').split('\n').map((line) => ({ runs: [{ text: line }] }));
      }
      if (!paragraphs || !paragraphs.length) return;
      // Размер - как у текста в месте вставки
      const size = currentFormat().size;
      insertParas(paragraphs, size ? { size } : {});
    });

    // Перетаскивание могло бы принести чужую разметку
    edit.addEventListener('drop', (e) => e.preventDefault());

    document.addEventListener('selectionchange', () => {
      if (!selectionInside()) return;
      read();
      if (pending && pendingAt && cmp(pendingAt, lastSel.start) !== 0) pending = null;
      sync();
    });

    // ---------- Внешний интерфейс ----------

    function setModel(paragraphs) {
      model = paragraphs.length ? clone(paragraphs.map((p) => ({ align: p.align || 'both',
        indent: p.indent !== false, runs: p.runs || [] }))) : emptyModel();
      undo.length = 0;
      redo.length = 0;
      pending = null;
      render({ start: { p: 0, o: 0 }, end: { p: 0, o: 0 } });
      changed();
    }

    setModel([]);
    ready = true;

    return {
      // Абзацы для документа: без пустых абзацев и лишнего оформления
      paragraphs: () => normalizeRich(read().model),
      plainText: () => richPlainText(normalizeRich(read().model)),
      setParagraphs: setModel,
      setText: (text) => setModel(String(text).split(/\r?\n/).map((line) => ({ runs: [{ text: line }] }))),
      focus: () => edit.focus(),
      element: edit,
    };
  }

  root.SZEditor = { create, readParagraphs };
})(typeof window !== 'undefined' ? window : globalThis);
