/*!
 * Text to Excel Converter: user interface.
 * All processing happens in this browser tab. Nothing is sent to a server.
 */
"use strict";
(() => {
  const $ = (id) => document.getElementById(id);
  const el = {
    text: $('textIn'), chars: $('charCount'), lines: $('lineCount'), delimInfo: $('delimInfo'),
    status: $('status'), statusText: $('statusText'), alerts: $('alerts'),
    fileInput: $('fileInput'), drop: $('dropzone'), fileChip: $('fileChip'), fileName: $('fileName'),
    delimSel: $('delimSel'), customWrap: $('customWrap'), customDelim: $('customDelim'), optHeader: $('optHeader'), dateOrder: $('dateOrder'),
    stRows: $('stRows'), stCols: $('stCols'), stEmpty: $('stEmpty'), stDup: $('stDup'), stSize: $('stSize'), stEmptyBox: $('stEmptyBox'), stDupBox: $('stDupBox'),
    dims: $('dims'), preview: $('previewArea'), pager: $('pager'), pageInfo: $('pageInfo'), pageSize: $('pageSize'),
    colList: $('colList'), addColForm: $('addColForm'), newColName: $('newColName'), newColDefault: $('newColDefault'),
    sheetName: $('sheetName'), fileNameIn: $('fileNameIn'), tableHint: $('tableHint'),
    splitRowsWrap: $('splitRowsWrap'), splitRows: $('splitRows'), splitColWrap: $('splitColWrap'), splitCol: $('splitCol'), splitHint: $('splitHint'),
    progress: $('progress'), exportSummary: $('exportSummary'), convertBtn: $('convertBtn'), mobileConvert: $('mobileConvert'),
  };

  const CLEAN_IDS = { removeEmptyRows: 'clEmptyRows', removeEmptyCols: 'clEmptyCols', trim: 'clTrim', removeDupes: 'clDupes', normalize: 'clNormalize', numbers: 'clNumbers', dates: 'clDates' };
  const FMT_IDS = { bold: 'fxBold', freeze: 'fxFreeze', autoWidth: 'fxWidth', filter: 'fxFilter', table: 'fxTable', altRows: 'fxAlt' };
  const DEFAULTS = {
    delimMode: 'auto', customDelim: '', hasHeader: true, dateOrder: 'dmy',
    clean: { removeEmptyRows: true, removeEmptyCols: true, trim: true, removeDupes: false, normalize: true, numbers: true, dates: true },
    sheetName: 'Sheet1', fileName: 'converted-data.xlsx',
    fmt: { bold: true, freeze: true, autoWidth: true, filter: true, table: false, altRows: true },
    split: { mode: 'none', rows: 1000, col: '' },
  };
  const TYPE_LABEL = { auto: 'Auto detect', text: 'Text', number: 'Number', date: 'Date', boolean: 'Boolean' };

  const SAMPLE = [
    'Name,Email,Phone,City,Joining Date,Salary,Active',
    'Aarav Patel,aarav.patel@example.com,9876543210,Ahmedabad,2024-04-01,52000,yes',
    'Diya Shah,diya.shah@example.com,9824012345,Patan,15/07/2023,48500.50,yes',
    'Mohammed Ansari,m.ansari@example.com,9898011122,Siddhpur,03-01-2025,61000,no',
    '"Joshi, Kavya",kavya.joshi@example.com,09426011223,Mehsana,2022-11-20,57250,yes',
    'Rahul Mehta,rahul.mehta@example.com,9712345678,Ahmedabad,01/02/2021,"1,02,000",yes',
    '',
    'Neha Trivedi,neha.t@example.com,9909912345,Patan,2025-06-30,45000,no',
    '  Imran Sheikh  ,imran.sheikh@example.com,9638527410,Siddhpur,12/12/2020,73800,yes',
    'Rahul Mehta,rahul.mehta@example.com,9712345678,Ahmedabad,01/02/2021,"1,02,000",yes',
    'Pooja Raval,pooja.raval@example.com,9426543210,Mehsana,2023-09-18,50500,yes',
    'Karan Desai,karan.desai@example.com,9586012347,Ahmedabad,28-02-2024,66400,no',
  ].join('\n');

  const state = {
    fileName: '', result: null, cols: [], signature: '', addSeq: 0,
    page: 1, pageSize: 25, opts: structuredClone(DEFAULTS), busy: false,
  };

  /* ---------- helpers ---------- */
  const fmtInt = (n) => n.toLocaleString('en-IN');
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function fmtBytes(b) { if (b < 1024) return `${b} B`; if (b < 1048576) return `${(b / 1024).toFixed(b < 10240 ? 1 : 0)} KB`; return `${(b / 1048576).toFixed(1)} MB`; }
  const tick = () => new Promise(r => setTimeout(r, 30));
  function setStatus(s) {
    const map = { ready: 'Ready', processing: 'Processing', converted: 'Converted', error: 'Error' };
    el.status.dataset.state = s; el.statusText.textContent = map[s];
  }
  function alertMsg(kind, msg, sticky) {
    const div = document.createElement('div');
    div.className = `alert ${kind}`;
    div.setAttribute('role', kind === 'error' ? 'alert' : 'status');
    div.innerHTML = `<p></p><button type="button" aria-label="Dismiss">×</button>`;
    div.querySelector('p').textContent = msg;
    div.querySelector('button').onclick = () => div.remove();
    el.alerts.appendChild(div);
    if (!sticky && kind !== 'error') setTimeout(() => div.remove(), 6000);
    return div;
  }
  function clearAlerts(kind) { [...el.alerts.children].forEach(a => { if (!kind || a.classList.contains(kind)) a.remove(); }); }

  /* ---------- theme ---------- */
  const root = document.documentElement;
  function effectiveDark() { const t = root.getAttribute('data-theme'); return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches; }
  function paintThemeBtn() {
    const dark = effectiveDark();
    $('themeLabel').textContent = dark ? 'Light mode' : 'Dark mode';
    $('themeIcon').innerHTML = dark
      ? '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>'
      : '<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>';
  }
  try { const t = localStorage.getItem('t2x-theme'); if (t === 'dark' || t === 'light') root.setAttribute('data-theme', t); } catch (e) {}
  $('themeBtn').onclick = () => {
    const next = effectiveDark() ? 'light' : 'dark';
    root.setAttribute('data-theme', next);
    try { localStorage.setItem('t2x-theme', next); } catch (e) {}
    paintThemeBtn();
  };
  try { matchMedia('(prefers-color-scheme: dark)').addEventListener('change', paintThemeBtn); } catch (e) {}
  paintThemeBtn();

  /* ---------- options <-> UI ---------- */
  function readOpts() {
    const o = state.opts;
    o.delimMode = el.delimSel.value;
    o.customDelim = el.customDelim.value.replace(/\\t/g, '\t');
    o.hasHeader = el.optHeader.checked;
    o.dateOrder = el.dateOrder.value;
    for (const k in CLEAN_IDS) o.clean[k] = $(CLEAN_IDS[k]).checked;
    for (const k in FMT_IDS) o.fmt[k] = $(FMT_IDS[k]).checked;
    o.sheetName = el.sheetName.value;
    o.fileName = el.fileNameIn.value;
    o.split.mode = document.querySelector('input[name="splitMode"]:checked').value;
    o.split.rows = Math.max(1, parseInt(el.splitRows.value, 10) || 1000);
    o.split.col = el.splitCol.value;
  }
  function writeOpts() {
    const o = state.opts;
    el.delimSel.value = o.delimMode; el.customDelim.value = o.customDelim; el.optHeader.checked = o.hasHeader; el.dateOrder.value = o.dateOrder;
    for (const k in CLEAN_IDS) $(CLEAN_IDS[k]).checked = o.clean[k];
    for (const k in FMT_IDS) $(FMT_IDS[k]).checked = o.fmt[k];
    el.sheetName.value = o.sheetName; el.fileNameIn.value = o.fileName;
    document.querySelector(`input[name="splitMode"][value="${o.split.mode}"]`).checked = true;
    el.splitRows.value = o.split.rows;
    syncConditional();
  }
  function syncConditional() {
    el.customWrap.hidden = el.delimSel.value !== 'custom';
    el.tableHint.hidden = !$('fxTable').checked;
    const mode = document.querySelector('input[name="splitMode"]:checked').value;
    el.splitRowsWrap.hidden = mode !== 'rows';
    el.splitColWrap.hidden = mode !== 'column';
    updateSplitHint();
  }
  function updateSplitHint() {
    const mode = document.querySelector('input[name="splitMode"]:checked').value;
    const n = state.result ? state.result.rows.length : 0;
    if (mode === 'none') el.splitHint.textContent = 'All rows go into a single worksheet.';
    else if (mode === 'rows') {
      const per = Math.max(1, parseInt(el.splitRows.value, 10) || 1000);
      el.splitHint.textContent = n ? `Creates ${fmtInt(Math.max(1, Math.ceil(n / per)))} sheet(s), each with its own header row.` : 'Each sheet repeats the header row.';
    } else {
      const col = state.cols.find(c => c.key === el.splitCol.value);
      if (!col || !n) { el.splitHint.textContent = 'Pick a column such as City.'; return; }
      const vals = new Set();
      for (const r of state.result.rows) { vals.add(String(cellOf(r, col)).trim() || '(blank)'); if (vals.size > 500) break; }
      el.splitHint.textContent = vals.size > 500 ? 'More than 500 distinct values. Pick a column with fewer values.' : `Creates ${fmtInt(vals.size)} sheet(s), one per ${col.name} value.`;
    }
  }

  /* ---------- core pipeline ---------- */
  const cellOf = (row, col) => (col.src >= 0 ? (row[col.src] ?? '') : col.def);

  function effectiveType(col) { return col.type === 'auto' ? col.inferred : col.type; }

  function rebuildCols(result) {
    const sig = result.headers.join('\u0001') + '|' + result.headers.length;
    if (sig === state.signature && state.cols.length) {
      for (const c of state.cols) if (c.src >= 0) c.inferred = result.inferred[c.src] || 'text';
      return;
    }
    state.signature = sig;
    state.cols = result.headers.map((h, i) => ({ key: 's' + i, src: i, name: h, type: 'auto', inferred: result.inferred[i] }));
  }

  function runPipeline() {
    readOpts();
    const text = el.text.value;
    clearAlerts('error'); clearAlerts('warn');
    updateCounts(text);
    if (!text.trim()) {
      state.result = null; state.cols = []; state.signature = '';
      el.delimInfo.innerHTML = 'Detected delimiter: <b>—</b>';
      renderAll();
      setStatus('ready');
      return;
    }
    const o = state.opts;
    let result;
    try {
      result = Core.process(text, {
        delimMode: o.delimMode, customDelim: o.customDelim, hasHeader: o.hasHeader, dateOrder: o.dateOrder,
        trim: o.clean.trim, normalize: o.clean.normalize, removeEmptyRows: o.clean.removeEmptyRows,
        removeEmptyCols: o.clean.removeEmptyCols, removeDupes: o.clean.removeDupes, numbers: o.clean.numbers, dates: o.clean.dates,
      });
    } catch (e) {
      state.result = null; state.cols = []; state.signature = '';
      renderAll();
      setStatus('error');
      alertMsg('error', o.delimMode === 'custom' && !o.customDelim
        ? 'Enter a custom delimiter, or choose another delimiter option.'
        : 'Unable to parse the provided data. Please check the delimiter or input format.');
      return;
    }
    el.delimInfo.innerHTML = `${o.delimMode === 'auto' ? 'Detected delimiter' : 'Delimiter'}: <b>${esc(result.delimLabel)}</b>`;
    state.result = result;
    if (!result.headers.length) {
      state.cols = []; state.signature = '';
      renderAll(); setStatus('error');
      alertMsg('error', 'No valid columns were detected.');
      return;
    }
    if (result.unbalanced) alertMsg('warn', 'A quotation mark is never closed, so the rest of the text was read as one cell. Check the quotes in your data.', true);
    rebuildCols(result);
    const maxPage = Math.max(1, Math.ceil(result.rows.length / state.pageSize));
    if (state.page > maxPage) state.page = maxPage;
    renderAll();
    setStatus('ready');
  }

  let debounceT = 0;
  function schedule(delay) {
    clearTimeout(debounceT);
    const big = el.text.value.length > 1_000_000;
    if (big) setStatus('processing');
    debounceT = setTimeout(runPipeline, delay ?? (big ? 600 : 220));
  }

  function updateCounts(text) {
    el.chars.textContent = fmtInt(text.length);
    let lines = 0;
    if (text.length) { lines = 1; let i = -1; while ((i = text.indexOf('\n', i + 1)) !== -1) lines++; if (text.endsWith('\n')) lines--; }
    el.lines.textContent = fmtInt(lines);
  }

  /* ---------- rendering ---------- */
  function renderAll() { renderStats(); renderPreview(); renderCols(); renderSplitCols(); updateSplitHint(); updateSummary(); }

  function renderStats() {
    const r = state.result;
    const text = el.text.value;
    if (!r) {
      for (const k of ['stRows', 'stCols', 'stEmpty', 'stDup']) el[k].textContent = '—';
      el.stSize.textContent = text ? fmtBytes(new Blob([text]).size) : '—';
      el.stEmptyBox.classList.remove('flag'); el.stDupBox.classList.remove('flag');
      el.dims.textContent = '0 rows × 0 columns';
      return;
    }
    el.stRows.textContent = fmtInt(r.rows.length);
    el.stCols.textContent = fmtInt(state.cols.length);
    el.stEmpty.textContent = fmtInt(r.stats.emptyRows);
    el.stDup.textContent = fmtInt(r.stats.dupRows);
    el.stEmptyBox.classList.toggle('flag', r.stats.emptyRows > 0 && !state.opts.clean.removeEmptyRows);
    el.stDupBox.classList.toggle('flag', r.stats.dupRows > 0 && !state.opts.clean.removeDupes);
    el.stSize.textContent = fmtBytes(text.length < 5_000_000 ? new Blob([text]).size : text.length);
    el.dims.textContent = `${fmtInt(r.rows.length)} rows × ${fmtInt(state.cols.length)} columns`;
  }

  function renderPreview() {
    const r = state.result;
    if (!r || !state.cols.length) {
      el.preview.innerHTML = `<div class="empty-state"><strong>No data yet</strong>Paste text above, drop a file, or load the sample data to see a preview.</div>`;
      el.pager.hidden = true;
      return;
    }
    const total = r.rows.length;
    const pages = Math.max(1, Math.ceil(total / state.pageSize));
    const page = Math.min(state.page, pages);
    const start = (page - 1) * state.pageSize;
    const end = Math.min(total, start + state.pageSize);
    const cols = state.cols;
    const types = cols.map(effectiveType);
    let h = '<div class="table-wrap"><table class="data-grid"><thead><tr class="letters"><th class="rn"></th>';
    for (let c = 0; c < cols.length; c++) h += `<th>${Core.colLetter(c)}</th>`;
    h += '</tr><tr class="names"><th class="rn">1</th>';
    for (let c = 0; c < cols.length; c++) h += `<th title="${esc(cols[c].name)}">${esc(cols[c].name)}<span class="type-tag ${types[c]}">${types[c]}</span></th>`;
    h += '</tr></thead><tbody>';
    if (!total) h += `<tr><td class="rn"></td><td colspan="${cols.length}" class="muted-cell">The header row was found, but there are no data rows.</td></tr>`;
    for (let i = start; i < end; i++) {
      const row = r.rows[i];
      h += `<tr><td class="rn">${i + 2}</td>`;
      for (let c = 0; c < cols.length; c++) {
        const v = cellOf(row, cols[c]);
        const s = String(v);
        if (s.trim() === '') { h += '<td class="empty"></td>'; continue; }
        const num = types[c] === 'number' || types[c] === 'date';
        const disp = s.length > 200 ? s.slice(0, 200) + '…' : s;
        h += `<td${num ? ' class="num"' : ''}${s.length > 30 ? ` title="${esc(s.slice(0, 500))}"` : ''}>${esc(disp)}</td>`;
      }
      h += '</tr>';
    }
    h += '</tbody></table></div>';
    el.preview.innerHTML = h;
    el.pager.hidden = total <= 25 && state.pageSize === 25;
    el.pageInfo.textContent = total ? `Rows ${fmtInt(start + 1)}–${fmtInt(end)} of ${fmtInt(total)} · Page ${fmtInt(page)} of ${fmtInt(pages)}` : '';
    $('firstPage').disabled = $('prevPage').disabled = page <= 1;
    $('nextPage').disabled = $('lastPage').disabled = page >= pages;
  }

  function renderCols() {
    if (!state.cols.length) {
      el.colList.innerHTML = `<div class="empty-state compact">Columns appear here after your data is parsed.</div>`;
      el.addColForm.hidden = true;
      return;
    }
    el.addColForm.hidden = false;
    const n = state.cols.length;
    let h = '';
    state.cols.forEach((c, i) => {
      const opts = Object.keys(TYPE_LABEL).map(t => `<option value="${t}"${c.type === t ? ' selected' : ''}>${TYPE_LABEL[t]}${t === 'auto' ? ` (${c.inferred})` : ''}</option>`).join('');
      h += `<div class="col-row" data-i="${i}">
        <span class="col-letter">${Core.colLetter(i)}</span>
        <input class="input" data-act="rename" value="${esc(c.name)}" aria-label="Name of column ${Core.colLetter(i)}">
        <select class="select" data-act="type" aria-label="Data type of ${esc(c.name)}">${opts}</select>
        <div class="col-actions">
          <button class="btn ghost icon" type="button" data-act="up" aria-label="Move ${esc(c.name)} left"${i === 0 ? ' disabled' : ''}><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="m6 15 6-6 6 6"/></svg></button>
          <button class="btn ghost icon" type="button" data-act="down" aria-label="Move ${esc(c.name)} right"${i === n - 1 ? ' disabled' : ''}><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="m6 9 6 6 6-6"/></svg></button>
          <button class="btn ghost icon danger-text" type="button" data-act="del" aria-label="Delete ${esc(c.name)}"${n === 1 ? ' disabled' : ''}><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg></button>
        </div>
        ${c.src < 0 ? `<div class="inferred">Added column${c.def ? ` · every row: “${esc(c.def)}”` : ' · empty'}</div>` : ''}
      </div>`;
    });
    el.colList.innerHTML = h;
  }

  function renderSplitCols() {
    const cur = el.splitCol.value || state.opts.split.col;
    el.splitCol.innerHTML = state.cols.map(c => `<option value="${c.key}">${esc(c.name)}</option>`).join('') || '<option value="">No columns</option>';
    if (state.cols.some(c => c.key === cur)) el.splitCol.value = cur;
    else { const city = state.cols.find(c => /city|region|state|category|department|type/i.test(c.name)); if (city) el.splitCol.value = city.key; }
  }

  function updateSummary() {
    const r = state.result;
    if (!r || !state.cols.length) { el.exportSummary.textContent = 'Paste or upload data to begin.'; return; }
    el.exportSummary.textContent = `${fmtInt(r.rows.length)} rows × ${fmtInt(state.cols.length)} columns → ${Core.cleanFileName(el.fileNameIn.value)}`;
  }

  /* ---------- column manager events ---------- */
  el.colList.addEventListener('input', (e) => {
    const t = e.target; if (t.dataset.act !== 'rename') return;
    const i = +t.closest('.col-row').dataset.i;
    state.cols[i].name = t.value;
    renderPreview(); renderSplitCols();
  });
  el.colList.addEventListener('change', (e) => {
    const t = e.target;
    const i = +t.closest('.col-row')?.dataset.i;
    if (t.dataset.act === 'type') { state.cols[i].type = t.value; renderPreview(); }
    if (t.dataset.act === 'rename' && !t.value.trim()) { state.cols[i].name = `Column ${i + 1}`; t.value = state.cols[i].name; renderPreview(); renderSplitCols(); }
  });
  el.colList.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const i = +b.closest('.col-row').dataset.i;
    const cols = state.cols;
    if (b.dataset.act === 'up' && i > 0) [cols[i - 1], cols[i]] = [cols[i], cols[i - 1]];
    else if (b.dataset.act === 'down' && i < cols.length - 1) [cols[i + 1], cols[i]] = [cols[i], cols[i + 1]];
    else if (b.dataset.act === 'del' && cols.length > 1) { const [gone] = cols.splice(i, 1); alertMsg('info', `Removed column “${gone.name}”. Use “Restore original columns” to bring it back.`); }
    else return;
    renderAll();
    const sel = el.colList.querySelector(`.col-row[data-i="${b.dataset.act === 'up' ? i - 1 : b.dataset.act === 'down' ? i + 1 : Math.min(i, cols.length - 1)}"] button[data-act="${b.dataset.act}"]`);
    if (sel && !sel.disabled) sel.focus();
  });
  el.addColForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if (!state.result) return;
    const name = el.newColName.value.trim() || `Column ${state.cols.length + 1}`;
    const def = el.newColDefault.value;
    state.cols.push({ key: 'a' + (++state.addSeq), src: -1, name, def, type: 'auto', inferred: def.trim() ? Core.inferType([def], name, { numbers: state.opts.clean.numbers, dates: state.opts.clean.dates, dateOrder: state.opts.dateOrder }) : 'text' });
    el.newColName.value = ''; el.newColDefault.value = '';
    renderAll();
  });
  $('restoreCols').onclick = () => { if (!state.result) return; state.signature = ''; rebuildCols(state.result); renderAll(); };

  /* ---------- pagination ---------- */
  const pages = () => Math.max(1, Math.ceil((state.result ? state.result.rows.length : 0) / state.pageSize));
  $('firstPage').onclick = () => { state.page = 1; renderPreview(); };
  $('prevPage').onclick = () => { state.page = Math.max(1, state.page - 1); renderPreview(); };
  $('nextPage').onclick = () => { state.page = Math.min(pages(), state.page + 1); renderPreview(); };
  $('lastPage').onclick = () => { state.page = pages(); renderPreview(); };
  el.pageSize.onchange = () => { state.pageSize = +el.pageSize.value; state.page = 1; renderPreview(); };

  /* ---------- input events ---------- */
  el.text.addEventListener('input', () => { if (state.fileName) setFile(''); schedule(); });
  function setFile(name, size) {
    state.fileName = name;
    el.fileChip.hidden = !name;
    el.fileName.textContent = name ? `${name}${size != null ? ` · ${fmtBytes(size)}` : ''}` : '';
  }
  const OK_EXT = /\.(txt|csv|tsv|log)$/i;
  function loadFile(file) {
    if (!file) return;
    clearAlerts();
    if (!OK_EXT.test(file.name)) { setStatus('error'); alertMsg('error', `“${file.name}” is not a supported file. Upload a .txt, .csv, .tsv or .log file.`); return; }
    if (file.size === 0) { setStatus('error'); alertMsg('error', `“${file.name}” is empty. Please enter or upload some data first.`); return; }
    setStatus('processing');
    const reader = new FileReader();
    reader.onload = () => {
      let text = String(reader.result || '');
      if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
      el.text.value = text;
      setFile(file.name, file.size);
      if (/\.tsv$/i.test(file.name) && el.delimSel.value === 'auto') { /* auto detection handles tabs */ }
      if (!/\.(xlsx|xls)$/i.test(el.fileNameIn.value) || el.fileNameIn.value === DEFAULTS.fileName) el.fileNameIn.value = file.name.replace(/\.[^.]+$/, '') + '.xlsx';
      state.page = 1;
      runPipeline();
      if (state.result) alertMsg('success', `Loaded ${file.name}: ${fmtInt(state.result.rows.length)} rows × ${fmtInt(state.cols.length)} columns.`);
    };
    reader.onerror = () => { setStatus('error'); alertMsg('error', `Could not read “${file.name}”. Try saving it again as plain text.`); };
    reader.readAsText(file);
  }
  el.fileInput.addEventListener('change', () => { loadFile(el.fileInput.files[0]); el.fileInput.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => el.drop.addEventListener(ev, (e) => { e.preventDefault(); el.drop.classList.add('over'); }));
  ['dragleave', 'dragend'].forEach(ev => el.drop.addEventListener(ev, (e) => { if (!el.drop.contains(e.relatedTarget)) el.drop.classList.remove('over'); }));
  el.drop.addEventListener('drop', (e) => { e.preventDefault(); el.drop.classList.remove('over'); loadFile(e.dataTransfer.files[0]); });
  // Dropping a file onto the textarea should load it too, instead of the browser opening it.
  el.text.addEventListener('dragover', (e) => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
  el.text.addEventListener('drop', (e) => { if (e.dataTransfer.files.length) { e.preventDefault(); loadFile(e.dataTransfer.files[0]); } });

  $('pasteBtn').onclick = async () => {
    try {
      const t = await navigator.clipboard.readText();
      if (!t) throw new Error('empty');
      el.text.value = t; setFile(''); state.page = 1; runPipeline();
    } catch (e) {
      el.text.focus(); el.text.select();
      alertMsg('info', 'Press Ctrl+V (or ⌘V on Mac) to paste into the text box.');
    }
  };
  function loadSample() {
    el.text.value = SAMPLE; setFile('Sample data'); state.page = 1; state.signature = '';
    runPipeline();
  }
  $('sampleBtn').onclick = loadSample;
  $('qaSample').onclick = loadSample;
  function clearText() {
    el.text.value = ''; setFile(''); state.page = 1; state.signature = ''; state.cols = [];
    clearAlerts(); el.progress.hidden = true; resetProgress(); runPipeline(); el.text.focus();
  }
  $('clearBtn').onclick = clearText;
  $('qaClear').onclick = clearText;

  $('qaCopy').onclick = () => {
    if (!state.result || !state.cols.length) { alertMsg('error', 'Please enter or upload some data first.'); return; }
    const q = (s) => { s = String(s); return /[\t\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
    const lines = [state.cols.map(c => q(c.name)).join('\t')];
    for (const r of state.result.rows) lines.push(state.cols.map(c => q(cellOf(r, c))).join('\t'));
    const tsv = lines.join('\n');
    const fallback = () => {
      const prev = el.text.value;
      if (tsv.length > 2_000_000) { alertMsg('error', 'This data is too large to copy. Download the Excel file instead.'); return; }
      const ta = document.createElement('textarea');
      ta.value = tsv; ta.style.position = 'fixed'; ta.style.opacity = '0'; document.body.appendChild(ta); ta.select();
      let ok = false; try { ok = document.execCommand('copy'); } catch (e) {}
      ta.remove();
      alertMsg(ok ? 'success' : 'error', ok ? `Copied ${fmtInt(state.result.rows.length)} rows. Paste them straight into Excel or Google Sheets.` : 'Copying is blocked here. Download the Excel file instead.');
      void prev;
    };
    try {
      navigator.clipboard.writeText(tsv).then(
        () => alertMsg('success', `Copied ${fmtInt(state.result.rows.length)} rows as tab-separated text. Paste them straight into Excel or Google Sheets.`),
        fallback);
    } catch (e) { fallback(); }
  };

  $('qaReset').onclick = () => {
    state.opts = structuredClone(DEFAULTS);
    writeOpts();
    state.signature = '';
    state.pageSize = 25; el.pageSize.value = '25'; state.page = 1;
    runPipeline();
    alertMsg('success', 'Settings restored to their defaults.');
  };

  // Options that change parsing re-run the pipeline; Excel-only options just refresh the summary.
  ['delimSel', 'optHeader', 'dateOrder', ...Object.values(CLEAN_IDS)].forEach(id => $(id).addEventListener('change', () => { syncConditional(); state.page = 1; runPipeline(); }));
  el.customDelim.addEventListener('input', () => schedule(300));
  Object.values(FMT_IDS).forEach(id => $(id).addEventListener('change', () => { readOpts(); syncConditional(); }));
  document.querySelectorAll('input[name="splitMode"]').forEach(r => r.addEventListener('change', syncConditional));
  el.splitRows.addEventListener('input', updateSplitHint);
  el.splitCol.addEventListener('change', updateSplitHint);
  el.fileNameIn.addEventListener('input', updateSummary);
  el.fileNameIn.addEventListener('change', () => { el.fileNameIn.value = Core.cleanFileName(el.fileNameIn.value); updateSummary(); });

  /* ---------- export ---------- */
  // Hand the generated file to the browser's normal download flow. The blob never leaves the device.
  function saveBlob(blob, fileName) {
    if (window.navigator && typeof window.navigator.msSaveOrOpenBlob === 'function') { window.navigator.msSaveOrOpenBlob(blob, fileName); return; }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = fileName; a.rel = 'noopener'; a.hidden = true;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  function resetProgress() { el.progress.querySelectorAll('li').forEach(li => li.className = ''); }
  function step(i, cls) { const li = el.progress.querySelector(`li[data-step="${i}"]`); if (li) li.className = cls; }

  async function convert() {
    if (state.busy) return;
    clearTimeout(debounceT);
    clearAlerts();
    el.progress.hidden = false; resetProgress();
    let cur = 0;
    const fail = (msg) => { step(cur, 'fail'); setStatus('error'); alertMsg('error', msg); };
    state.busy = true; el.convertBtn.disabled = el.mobileConvert.disabled = $('qaDownload').disabled = true;
    try {
      setStatus('processing');
      step(0, 'active'); await tick();
      if (!el.text.value.trim()) return fail('Please enter or upload some data first.');
      if (typeof XLSX === 'undefined' || typeof JSZip === 'undefined') return fail('The Excel library did not load. Check your connection and reload the page.');
      step(0, 'done'); cur = 1; step(1, 'active'); await tick();
      runPipeline();
      if (!state.result) { return fail(state.opts.delimMode === 'custom' && !state.opts.customDelim ? 'Enter a custom delimiter, or choose another delimiter option.' : 'Unable to parse the provided data. Please check the delimiter or input format.'); }
      if (!state.cols.length) return fail('No valid columns were detected.');
      step(1, 'done'); cur = 2; step(2, 'active'); await tick();
      readOpts();
      const o = state.opts;
      const table = {
        headers: state.cols.map(c => c.name),
        types: state.cols.map(effectiveType),
        rows: state.result.rows.map(r => state.cols.map(c => cellOf(r, c))),
      };
      const splitIdx = o.split.mode === 'column' ? state.cols.findIndex(c => c.key === o.split.col) : -1;
      if (o.split.mode === 'column' && splitIdx < 0) return fail('Choose a column to split the sheets by.');
      const fileName = Core.cleanFileName(o.fileName);
      el.fileNameIn.value = fileName;
      const settings = { sheetName: o.sheetName.trim() || 'Sheet1', fileName, fmt: o.fmt, dateOrder: o.dateOrder, split: { mode: o.split.mode, rows: o.split.rows, colIndex: splitIdx } };
      step(2, 'done'); cur = 3; step(3, 'active'); await tick();
      let built;
      try { built = await Core.buildWorkbook(XLSX, JSZip, table, settings); }
      catch (e) {
        const m = String(e && e.message || e);
        if (m.startsWith('TOO_MANY_SHEETS')) return fail(`Splitting would create ${fmtInt(+m.split(':')[1])} sheets. Pick a column with fewer distinct values (500 at most).`);
        if (m === 'TOO_MANY_ROWS') return fail('A sheet would exceed Excel’s limit of 1,048,576 rows. Turn on “Multiple sheets → By rows”.');
        if (m === 'TOO_MANY_COLS') return fail('Excel supports at most 16,384 columns. Check the delimiter setting.');
        console.error(e);
        return fail('Something went wrong while creating the workbook. Try turning off “Format as Excel Table” and convert again.');
      }
      step(3, 'done'); cur = 4; step(4, 'active'); await tick();
      const blob = new Blob([built.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const sheetsNote = built.sheets.length > 1 ? ` across ${built.sheets.length} sheets` : '';
      saveBlob(blob, fileName);
      step(4, 'done'); setStatus('converted');
      alertMsg('success', `Download ready: ${fileName} (${fmtBytes(blob.size)}) with ${fmtInt(table.rows.length)} rows${sheetsNote}.`);
    } catch (e) {
      console.error(e);
      fail('Something went wrong during conversion. Check your data and try again.');
    } finally {
      state.busy = false; el.convertBtn.disabled = el.mobileConvert.disabled = $('qaDownload').disabled = false;
    }
  }
  el.convertBtn.onclick = convert;
  el.mobileConvert.onclick = convert;
  $('qaDownload').onclick = convert;

  /* ---------- start in a working state with the sample loaded ---------- */
  writeOpts();
  el.text.value = SAMPLE;
  setFile('Sample data');
  runPipeline();
})();
