/*!
 * Text to Excel Converter: parsing, cleaning, type detection and workbook building.
 * Pure functions; no DOM access. Runs in the browser and in Node (for tests).
 */
"use strict";
const Core = (() => {
  const DELIMS = {
    comma: { ch: ',', label: 'Comma (,)' },
    tab: { ch: '\t', label: 'Tab (\\t)' },
    semicolon: { ch: ';', label: 'Semicolon (;)' },
    pipe: { ch: '|', label: 'Pipe (|)' },
    colon: { ch: ':', label: 'Colon (:)' },
  };
  const DETECT_ORDER = ['tab', 'comma', 'semicolon', 'pipe'];

  function countOf(s, ch) { let n = 0, i = -1; while ((i = s.indexOf(ch, i + 1)) !== -1) n++; return n; }

  // Split a sample into logical records, blanking out quoted text so quoted delimiters and line breaks are ignored.
  function sampleRecords(text, max) {
    const out = [];
    let cur = '', inQ = false, atStart = true;
    const lim = Math.min(text.length, 200000);
    for (let i = 0; i < lim && out.length < max; i++) {
      const c = text[i];
      if (inQ) { if (c === '"') { if (text[i + 1] === '"') i++; else inQ = false; } continue; }
      if (c === '"' && atStart) { inQ = true; cur += 'Q'; atStart = false; continue; }
      if (c === '\n' || c === '\r') { if (cur.trim() !== '') out.push(cur); cur = ''; atStart = true; continue; }
      cur += c;
      atStart = c === ',' || c === '\t' || c === ';' || c === '|' || c === ':';
    }
    if (cur.trim() !== '' && out.length < max) out.push(cur);
    return out;
  }

  // Pick the delimiter whose per-line count is most consistent across a sample of lines.
  function detectDelimiter(text) {
    const lines = sampleRecords(text, 60);
    if (!lines.length) return null;
    const score = (key) => {
      const ch = DELIMS[key].ch;
      const counts = lines.map(l => countOf(l, ch));
      const freq = new Map();
      for (const c of counts) if (c > 0) freq.set(c, (freq.get(c) || 0) + 1);
      if (!freq.size) return null;
      let mode = 0, best = 0;
      for (const [c, f] of freq) if (f > best || (f === best && c > mode)) { best = f; mode = c; }
      return { key, mode, consistency: best / lines.length };
    };
    const cands = DETECT_ORDER.map(score).filter(c => c && c.consistency >= 0.8);
    if (cands.length) {
      cands.sort((a, b) => (b.consistency - a.consistency) || (b.mode - a.mode) || (DETECT_ORDER.indexOf(a.key) - DETECT_ORDER.indexOf(b.key)));
      return cands[0].key;
    }
    // Ragged data (rows with missing trailing fields): trust the delimiter the first line uses,
    // provided most lines contain it too.
    let ragged = null;
    for (const key of DETECT_ORDER) {
      const ch = DELIMS[key].ch;
      if (countOf(lines[0], ch) === 0) continue;
      const present = lines.filter(l => l.indexOf(ch) !== -1).length / lines.length;
      if (present >= 0.5 && (!ragged || present > ragged.present)) ragged = { key, present };
    }
    if (ragged) return ragged.key;
    // Colon only as a last resort, and not when the colons look like times or URLs.
    const looksTimeOrUrl = lines.some(l => /\d:\d\d|:\/\//.test(l));
    const colon = score('colon');
    if (colon && colon.consistency >= 0.9 && !looksTimeOrUrl) return 'colon';
    return null;
  }

  // RFC 4180 style parser: quoted fields, escaped quotes, multi-character delimiters, \n / \r\n / \r rows.
  function parse(text, delim) {
    const rows = [];
    const n = text.length;
    const dl = delim ? delim.length : 0;
    const d0 = dl ? delim.charCodeAt(0) : -1;
    let unbalanced = false;
    const isDelimAt = dl === 0 ? () => false
      : dl === 1 ? (k) => text.charCodeAt(k) === d0
      : (k) => text.charCodeAt(k) === d0 && text.startsWith(delim, k);
    let i = 0;
    let row = [];
    if (n === 0) return { rows, unbalanced };
    while (true) {
      let field;
      if (text.charCodeAt(i) === 34) {
        let val = '', j = i + 1;
        for (;;) {
          const q = text.indexOf('"', j);
          if (q === -1) { val += text.slice(j); i = n; unbalanced = true; break; }
          if (text.charCodeAt(q + 1) === 34) { val += text.slice(j, q + 1); j = q + 2; continue; }
          val += text.slice(j, q); i = q + 1; break;
        }
        let k = i;
        while (k < n) { const c = text.charCodeAt(k); if (c === 10 || c === 13 || isDelimAt(k)) break; k++; }
        if (k > i) val += text.slice(i, k);
        i = k; field = val;
      } else {
        let k = i;
        while (k < n) { const c = text.charCodeAt(k); if (c === 10 || c === 13 || isDelimAt(k)) break; k++; }
        field = text.slice(i, k); i = k;
      }
      row.push(field);
      if (i >= n) { rows.push(row); row = null; break; }
      if (isDelimAt(i)) {
        i += dl;
        if (i >= n) { row.push(''); rows.push(row); row = null; break; }
        continue;
      }
      if (text.charCodeAt(i) === 13 && text.charCodeAt(i + 1) === 10) i += 2; else i++;
      rows.push(row); row = [];
      if (i >= n) break;
    }
    if (row && row.length) rows.push(row);
    return { rows, unbalanced };
  }

  // Plain, thousands-grouped (1,250,000) and Indian-grouped (12,50,000) numbers.
  const NUM_RE = /^[-+]?(?:\d{1,3}(?:,\d{3})+|\d{1,2}(?:,\d{2})+,\d{3}|\d+)(?:\.\d+)?$|^[-+]?\.\d+$/;
  const BOOL_RE = /^(?:true|false|yes|no)$/i;
  const ISO_RE = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
  const DMY_RE = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
  const ID_HEADER_RE = /\b(?:phone|mobile|tel|telephone|contact|whatsapp|zip|zipcode|pin|pincode|postal|postcode|id|account|acct|a\/c|sku|ifsc|aadhaar|aadhar|pan|gstin|code)\b/i;

  function isIdLike(s) {
    const t = s.replace(/^[-+]/, '');
    if (/^0\d/.test(t)) return true;               // leading zero: 001245, 09426...
    const digits = t.replace(/[^\d]/g, '');
    if (/^\d+$/.test(t) && t.length >= 10) return true; // phone / account numbers
    return digits.length > 15;                      // beyond Excel's 15-digit precision
  }
  function toNumber(s) {
    if (!NUM_RE.test(s)) return null;
    const v = Number(s.replace(/,/g, ''));
    return Number.isFinite(v) ? v : null;
  }
  const EPOCH = Date.UTC(1899, 11, 30);
  function toDate(s, order) {
    let y, m, d, h = 0, mi = 0, se = 0, time = false, mt;
    if ((mt = ISO_RE.exec(s))) { y = +mt[1]; m = +mt[2]; d = +mt[3]; if (mt[4] != null) { time = true; h = +mt[4]; mi = +mt[5]; se = +(mt[6] || 0); } }
    else if ((mt = DMY_RE.exec(s))) {
      const a = +mt[1], b = +mt[2]; y = +mt[3];
      if (order === 'mdy') { m = a; d = b; } else { d = a; m = b; }
      if (mt[4] != null) { time = true; h = +mt[4]; mi = +mt[5]; se = +(mt[6] || 0); }
    } else return null;
    if (y < 1900 || y > 2200 || m < 1 || m > 12 || d < 1 || d > 31 || h > 23 || mi > 59 || se > 59) return null;
    const ms = Date.UTC(y, m - 1, d);
    const chk = new Date(ms);
    if (chk.getUTCMonth() !== m - 1 || chk.getUTCDate() !== d) return null;
    return { serial: (ms - EPOCH) / 86400000 + (h * 3600 + mi * 60 + se) / 86400, time };
  }
  function toBool(s) { if (!BOOL_RE.test(s)) return null; const l = s.toLowerCase(); return l === 'true' || l === 'yes'; }

  // Infer one type for a whole column so a column is never half numbers, half text.
  function inferType(values, header, opts) {
    let any = false, allNum = opts.numbers, allDate = opts.dates, allBool = true, idLike = false;
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      if (v == null) continue;
      const s = String(v).trim();
      if (s === '') continue;
      any = true;
      if (allBool && !BOOL_RE.test(s)) allBool = false;
      if (allNum) { if (toNumber(s) === null) allNum = false; else if (isIdLike(s)) idLike = true; }
      if (allDate && !toDate(s, opts.dateOrder)) allDate = false;
      if (!allNum && !allDate && !allBool) break;
    }
    if (!any) return 'text';
    if (allBool) return 'boolean';
    if (allDate) return 'date';
    if (allNum && !idLike && !(header && ID_HEADER_RE.test(header))) return 'number';
    return 'text';
  }

  const isBlank = (s) => s == null || String(s).trim() === '';

  function process(text, o) {
    let src = text;
    if (o.normalize) src = src.replace(/\r\n?/g, '\n');
    let delimKey = o.delimMode, delim = null;
    if (delimKey === 'auto') delimKey = detectDelimiter(src) || 'none';
    if (delimKey === 'custom') {
      if (!o.customDelim) throw new Error('PARSE');
      delim = o.customDelim;
    } else if (delimKey !== 'none') delim = DELIMS[delimKey].ch;
    const { rows: raw, unbalanced } = parse(src, delim);
    let width = 0;
    for (const r of raw) if (r.length > width) width = r.length;
    let rows = raw;
    for (const r of rows) {
      while (r.length < width) r.push('');
      if (o.trim) for (let c = 0; c < width; c++) r[c] = r[c].trim();
      if (o.normalize) for (let c = 0; c < width; c++) if (r[c].indexOf('\r') !== -1) r[c] = r[c].replace(/\r\n?/g, '\n');
    }
    const blankRow = (r) => { for (let c = 0; c < r.length; c++) if (!isBlank(r[c])) return false; return true; };
    let emptyRows = 0;
    for (const r of rows) if (blankRow(r)) emptyRows++;
    if (o.removeEmptyRows) rows = rows.filter(r => !blankRow(r));

    let headers;
    if (o.hasHeader && rows.length) {
      headers = rows[0].map((h, i) => (String(h).trim() || `Column ${i + 1}`));
      rows = rows.slice(1);
    } else {
      headers = Array.from({ length: width }, (_, i) => `Column ${i + 1}`);
    }

    // Duplicate counting ignores blank rows.
    const seen = new Set();
    let dupRows = 0;
    const deduped = [];
    for (const r of rows) {
      if (blankRow(r)) { deduped.push(r); continue; }
      const key = r.join('\u0001');
      if (seen.has(key)) { dupRows++; if (!o.removeDupes) deduped.push(r); }
      else { seen.add(key); deduped.push(r); }
    }
    rows = deduped;

    if (o.removeEmptyCols && width) {
      const keep = [];
      for (let c = 0; c < width; c++) {
        const headerBlank = !o.hasHeader || headers[c] === `Column ${c + 1}`;
        let has = false;
        for (let r = 0; r < rows.length; r++) if (!isBlank(rows[r][c])) { has = true; break; }
        if (has || !headerBlank) keep.push(c);
      }
      if (keep.length !== width) {
        headers = keep.map(c => headers[c]);
        rows = rows.map(r => keep.map(c => r[c]));
        if (!o.hasHeader) headers = headers.map((_, i) => `Column ${i + 1}`);
      }
    }

    const inferred = headers.map((h, c) => {
      const col = new Array(rows.length);
      for (let r = 0; r < rows.length; r++) col[r] = rows[r][c];
      return inferType(col, h, o);
    });

    return {
      headers, rows, inferred, unbalanced,
      delimKey, delimLabel: delimKey === 'custom' ? `Custom (${o.customDelim})` : delimKey === 'none' ? 'None (single column)' : DELIMS[delimKey].label,
      stats: { rows: rows.length, cols: headers.length, emptyRows, dupRows },
    };
  }

  // Convert one string into a SheetJS cell given the effective column type.
  function toCell(s, type, dateOrder) {
    if (s == null) return null;
    const t = String(s);
    if (t.trim() === '') return null;
    const v = t.trim();
    if (type === 'number') { const n = toNumber(v); if (n !== null) return { t: 'n', v: n }; }
    else if (type === 'date') { const d = toDate(v, dateOrder); if (d) return { t: 'n', v: d.serial, z: d.time ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd' }; }
    else if (type === 'boolean') { const b = toBool(v); if (b !== null) return { t: 'b', v: b }; }
    return { t: 's', v: t.length > 32767 ? t.slice(0, 32767) : t };
  }

  function colLetter(n) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  function cleanSheetName(name, used) {
    let base = String(name == null ? '' : name).replace(/[\[\]:*?\/\\]/g, '-').replace(/[\u0000-\u001f]/g, ' ').trim().replace(/^'+|'+$/g, '');
    if (!base) base = 'Sheet';
    if (base.toLowerCase() === 'history') base = 'History_';
    base = base.slice(0, 31);
    let out = base, k = 2;
    while (used.has(out.toLowerCase())) { const suf = ` (${k++})`; out = base.slice(0, 31 - suf.length) + suf; }
    used.add(out.toLowerCase());
    return out;
  }
  function uniqueHeaders(headers) {
    const used = new Set();
    return headers.map((h, i) => {
      let base = String(h).replace(/\s+/g, ' ').trim() || `Column ${i + 1}`;
      let out = base, k = 2;
      while (used.has(out.toLowerCase())) out = `${base} (${k++})`;
      used.add(out.toLowerCase());
      return out;
    });
  }
  function cleanFileName(name) {
    let n = String(name || '').trim().replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '-');
    n = n.replace(/\.xlsx$/i, '').replace(/\.+$/, '');
    if (!n) n = 'converted-data';
    return n + '.xlsx';
  }

  // table: {headers:[], types:[], rows:[[string]]}; split: {mode, rows, colIndex}
  function planSheets(table, sheetName, split) {
    const used = new Set();
    if (split.mode === 'rows' && split.rows > 0 && table.rows.length > split.rows) {
      const out = [];
      for (let i = 0, p = 1; i < table.rows.length; i += split.rows, p++) {
        out.push({ name: cleanSheetName(`${sheetName} ${p}`, used), rows: table.rows.slice(i, i + split.rows) });
      }
      return out;
    }
    if (split.mode === 'column' && split.colIndex >= 0) {
      const groups = new Map();
      for (const r of table.rows) {
        const raw = r[split.colIndex];
        const key = isBlank(raw) ? '(blank)' : String(raw).trim();
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(r);
      }
      return Array.from(groups, ([k, rows]) => ({ name: cleanSheetName(k, used), rows }));
    }
    return [{ name: cleanSheetName(sheetName, used), rows: table.rows }];
  }

  const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function buildWorksheet(XLSX, headers, types, rows, fmt, dateOrder) {
    const ws = {};
    const letters = headers.map((_, c) => colLetter(c));
    const nCols = headers.length;
    const useFill = !fmt.table;
    const border = { bottom: { style: 'thin', color: { rgb: 'B9C4BD' } } };
    const headStyle = {
      font: { bold: !!fmt.bold, color: { rgb: useFill ? 'FFFFFF' : '000000' } },
      alignment: { vertical: 'center' },
    };
    if (useFill) { headStyle.fill = { patternType: 'solid', fgColor: { rgb: '1D6B47' } }; headStyle.border = border; }
    for (let c = 0; c < nCols; c++) ws[letters[c] + '1'] = { t: 's', v: headers[c], s: headStyle };
    const altFill = { patternType: 'solid', fgColor: { rgb: 'EEF5F0' } };
    const altOn = fmt.altRows && useFill;
    const widths = headers.map(h => h.length);
    const widthSample = Math.min(rows.length, 3000);
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      const shade = altOn && (r % 2 === 1);
      const rn = r + 2;
      for (let c = 0; c < nCols; c++) {
        let cell = toCell(row[c], types[c], dateOrder);
        if (!cell) { if (shade) ws[letters[c] + rn] = { t: 's', v: '', s: { fill: altFill } }; continue; }
        if (cell.z || shade) {
          const s = {};
          if (cell.z) s.numFmt = cell.z;
          if (shade) s.fill = altFill;
          cell.s = s;
        }
        ws[letters[c] + rn] = cell;
        if (r < widthSample) {
          let len;
          if (cell.z) len = cell.z.length;
          else if (cell.t === 'b') len = 5;
          else { const str = String(row[c]); const nl = str.indexOf('\n'); len = nl === -1 ? str.length : nl; }
          if (len > widths[c]) widths[c] = len;
        }
      }
    }
    const lastRow = Math.max(1, rows.length + 1);
    ws['!ref'] = `A1:${letters[nCols - 1]}${lastRow}`;
    if (fmt.autoWidth) ws['!cols'] = widths.map(w => ({ wch: Math.max(8, Math.min(60, w + 3)) }));
    if (fmt.filter && !fmt.table) ws['!autofilter'] = { ref: ws['!ref'] };
    return { ws, ref: ws['!ref'], rowCount: rows.length };
  }

  async function postProcess(JSZip, bytes, metas, fmt, headers) {
    const needFreeze = fmt.freeze;
    const needTable = fmt.table && metas.some(m => m.rowCount > 0);
    if (!needFreeze && !needTable) return bytes;
    const zip = await JSZip.loadAsync(bytes);
    let ct = await zip.file('[Content_Types].xml').async('string');
    let tableNo = 0;
    for (let i = 0; i < metas.length; i++) {
      const m = metas[i];
      const path = `xl/worksheets/sheet${i + 1}.xml`;
      const f = zip.file(path);
      if (!f) continue;
      let xml = await f.async('string');
      if (needFreeze) {
        const pane = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
        if (/<sheetView\b[^>]*\/>/.test(xml)) xml = xml.replace(/<sheetView\b([^>]*)\/>/, `<sheetView$1>${pane}</sheetView>`);
        else if (/<sheetView\b[^>]*>/.test(xml)) xml = xml.replace(/(<sheetView\b[^>]*>)/, `$1${pane}`);
        else {
          const views = `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>`;
          if (/<dimension\b[^>]*\/>/.test(xml)) xml = xml.replace(/(<dimension\b[^>]*\/>)/, `$1${views}`);
          else xml = xml.replace(/(<worksheet\b[^>]*>)/, `$1${views}`);
        }
      }
      if (fmt.table && m.rowCount > 0) {
        tableNo++;
        const name = `Table${tableNo}`;
        const cols = headers.map((h, c) => `<tableColumn id="${c + 1}" name="${xmlEsc(h)}"/>`).join('');
        const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tableNo}" name="${name}" displayName="${name}" ref="${m.ref}" totalsRowShown="0"><autoFilter ref="${m.ref}"/><tableColumns count="${headers.length}">${cols}</tableColumns><tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0" showRowStripes="${fmt.altRows ? 1 : 0}" showColumnStripes="0"/></table>`;
        zip.file(`xl/tables/table${tableNo}.xml`, tableXml);
        ct = ct.replace('</Types>', `<Override PartName="/xl/tables/table${tableNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`);
        const relPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
        const rel = `<Relationship Id="rIdTbl${tableNo}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${tableNo}.xml"/>`;
        const rf = zip.file(relPath);
        if (rf) zip.file(relPath, (await rf.async('string')).replace('</Relationships>', rel + '</Relationships>'));
        else zip.file(relPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`);
        if (!/xmlns:r=/.test(xml)) xml = xml.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
        const part = `<tableParts count="1"><tablePart r:id="rIdTbl${tableNo}"/></tableParts>`;
        xml = /<extLst/.test(xml) ? xml.replace(/<extLst/, part + '<extLst') : xml.replace('</worksheet>', part + '</worksheet>');
      }
      zip.file(path, xml);
    }
    zip.file('[Content_Types].xml', ct);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  // Build the full .xlsx as bytes.
  async function buildWorkbook(XLSX, JSZip, table, settings) {
    const headers = uniqueHeaders(table.headers);
    const plans = planSheets(table, settings.sheetName || 'Sheet1', settings.split);
    if (plans.length > 500) throw new Error('TOO_MANY_SHEETS:' + plans.length);
    for (const p of plans) if (p.rows.length + 1 > 1048576) throw new Error('TOO_MANY_ROWS');
    if (headers.length > 16384) throw new Error('TOO_MANY_COLS');
    const wb = XLSX.utils.book_new();
    const metas = [];
    for (const p of plans) {
      const { ws, ref, rowCount } = buildWorksheet(XLSX, headers, table.types, p.rows, settings.fmt, settings.dateOrder);
      XLSX.utils.book_append_sheet(wb, ws, p.name);
      metas.push({ ref, rowCount, name: p.name });
    }
    wb.Props = { Title: settings.fileName.replace(/\.xlsx$/i, ''), Author: 'Text to Excel Converter' };
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
    const bytes = out instanceof Uint8Array ? out : new Uint8Array(out);
    const final = await postProcess(JSZip, bytes, metas, settings.fmt, headers);
    return { bytes: final, sheets: metas };
  }

  return { DELIMS, detectDelimiter, parse, process, inferType, toCell, toDate, toNumber, isIdLike, colLetter, cleanSheetName, cleanFileName, uniqueHeaders, planSheets, buildWorkbook };
})();

if (typeof module !== "undefined" && module.exports) module.exports = Core;
