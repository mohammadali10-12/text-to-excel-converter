/*
 * Test suite for the conversion core. Run with: npm test
 * Uses the same vendored SheetJS and JSZip builds the website loads.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const Core = require('../public/assets/js/core.js');
// Load SheetJS exactly as the browser does (as a plain script defining XLSX), not as a Node module.
const XLSX = new Function('module', 'require', 'define',
  fs.readFileSync(path.join(__dirname, '../public/assets/vendor/xlsx-js-style-1.2.0.min.js'), 'utf8') + '\nreturn XLSX;')();
const JSZip = require('../public/assets/vendor/jszip-3.10.1.min.js');

const OUT = path.join(__dirname, 'output');
fs.mkdirSync(OUT, { recursive: true });

const BASE = {
  delimMode: 'auto', customDelim: '', hasHeader: true, dateOrder: 'dmy',
  trim: true, normalize: true, removeEmptyRows: true, removeEmptyCols: true,
  removeDupes: false, numbers: true, dates: true,
};
const FMT = { bold: true, freeze: true, autoWidth: true, filter: true, table: false, altRows: true };

let passed = 0, failed = 0;
const tests = [];
const test = (name, fn) => tests.push({ name, fn });

function proc(text, extra) { return Core.process(text, Object.assign({}, BASE, extra || {})); }
async function build(table, extra) {
  const settings = Object.assign({ sheetName: 'Sheet1', fileName: 'test.xlsx', fmt: FMT, dateOrder: 'dmy', split: { mode: 'none' } }, extra || {});
  return Core.buildWorkbook(XLSX, JSZip, table, settings);
}
function readBack(bytes) { return XLSX.read(bytes, { type: 'array', cellDates: false }); }

/* ---------- delimiters ---------- */
const SIMPLE = { comma: 'Name,Age,City\nAli,25,Patan\nRahul,30,Ahmedabad', tab: 'Name\tAge\tCity\nAli\t25\tPatan\nRahul\t30\tAhmedabad', pipe: 'Name|Age|City\nAli|25|Patan\nRahul|30|Ahmedabad', semicolon: 'Name;Age;City\nAli;25;Patan\nRahul;30;Ahmedabad' };
for (const [key, text] of Object.entries(SIMPLE)) {
  test(`detects ${key} delimiter and parses rows`, () => {
    const r = proc(text);
    assert.strictEqual(r.delimKey, key);
    assert.deepStrictEqual(r.headers, ['Name', 'Age', 'City']);
    assert.deepStrictEqual(r.rows, [['Ali', '25', 'Patan'], ['Rahul', '30', 'Ahmedabad']]);
    assert.deepStrictEqual(r.inferred, ['text', 'number', 'text']);
  });
}
test('single-column text becomes one column', () => {
  const r = proc('Ali\nRahul\nAmit\nMohammed', { hasHeader: false });
  assert.strictEqual(r.delimKey, 'none');
  assert.deepStrictEqual(r.headers, ['Column 1']);
  assert.strictEqual(r.rows.length, 4);
});
test('colon delimiter is detected for key:value data', () => {
  assert.strictEqual(proc('key:value\nalpha:1\nbeta:2').delimKey, 'colon');
});
test('colons in timestamps are not treated as a delimiter', () => {
  assert.strictEqual(proc('2026-10-06 10:00:01 INFO start\n2026-10-06 10:00:02 WARN slow').delimKey, 'none');
});
test('custom multi-character delimiter', () => {
  const r = proc('a||b\n1||2', { delimMode: 'custom', customDelim: '||' });
  assert.deepStrictEqual(r.rows, [['1', '2']]);
});
test('empty custom delimiter throws a parse error', () => {
  assert.throws(() => proc('a,b', { delimMode: 'custom', customDelim: '' }), /PARSE/);
});

/* ---------- quoting & line breaks ---------- */
test('quoted fields with commas, escaped quotes and line breaks', () => {
  const r = proc('a,b\n"x, y","he said ""hi"""\n"multi\nline",z');
  assert.strictEqual(r.delimKey, 'comma');
  assert.deepStrictEqual(r.rows, [['x, y', 'he said "hi"'], ['multi\nline', 'z']]);
});
test('CRLF and CR line endings', () => {
  assert.strictEqual(proc('a,b\r\n1,2\r\n3,4\r\n').rows.length, 2);
  assert.strictEqual(proc('a,b\r1,2\r3,4').rows.length, 2);
});
test('unclosed quote is reported', () => {
  assert.strictEqual(proc('a,b\n"oops,1\n2,3').unbalanced, true);
});
test('ragged rows are padded', () => {
  assert.deepStrictEqual(proc('a,b,c\n1\n2,3').rows, [['1', '', ''], ['2', '3', '']]);
});

/* ---------- cleaning ---------- */
test('empty rows are counted and removed', () => {
  const r = proc('A,B\n1,x\n\n,\n2,y');
  assert.strictEqual(r.stats.emptyRows, 2);
  assert.strictEqual(r.rows.length, 2);
});
test('empty rows kept when the option is off', () => {
  assert.strictEqual(proc('A,B\n1,x\n\n2,y', { removeEmptyRows: false }).rows.length, 3);
});
test('duplicate rows are counted, and removed only when asked', () => {
  const t = 'A,B\n1,x\n1,x\n2,y';
  assert.strictEqual(proc(t).stats.dupRows, 1);
  assert.strictEqual(proc(t).rows.length, 3);
  assert.strictEqual(proc(t, { removeDupes: true }).rows.length, 2);
});
test('whitespace trimming is optional', () => {
  assert.strictEqual(proc('A\n  hi  ').rows[0][0], 'hi');
  assert.strictEqual(proc('A\n  hi  ', { trim: false }).rows[0][0], '  hi  ');
});
test('blank trailing columns are removed', () => {
  const r = proc('a,b,\n1,2,\n3,4,');
  assert.deepStrictEqual(r.headers, ['a', 'b']);
});
test('no header row gives Column 1..n', () => {
  assert.deepStrictEqual(proc('1,2\n3,4', { hasHeader: false }).headers, ['Column 1', 'Column 2']);
});

/* ---------- type detection ---------- */
test('numbers, dates and booleans are detected', () => {
  const r = proc('Amt,When,Flag\n99.50,2026-10-06,yes\n100,06/10/2026,no\n"1,250",06-10-2026,true');
  assert.deepStrictEqual(r.inferred, ['number', 'date', 'boolean']);
});
test('Indian digit grouping is a number', () => {
  assert.strictEqual(Core.toNumber('1,02,000'), 102000);
  assert.strictEqual(Core.toNumber('12,50,000.75'), 1250000.75);
});
test('phone numbers, leading zeros and ID columns stay text', () => {
  const r = proc('Phone,Zip,Qty,Customer ID\n9876543210,001245,3,101\n9876543211,382001,4,102');
  assert.deepStrictEqual(r.inferred, ['text', 'text', 'number', 'text']);
});
test('detection toggles disable number and date typing', () => {
  const r = proc('A,B\n1,2026-01-01', { numbers: false, dates: false });
  assert.deepStrictEqual(r.inferred, ['text', 'text']);
});
test('invalid calendar dates are rejected', () => {
  assert.strictEqual(Core.toDate('31/02/2026', 'dmy'), null);
  assert.ok(Core.toDate('29/02/2024', 'dmy'));
});
test('day/month order follows the setting', () => {
  const dmy = Core.toDate('06/10/2026', 'dmy').serial, mdy = Core.toDate('06/10/2026', 'mdy').serial;
  assert.strictEqual(dmy, Core.toDate('2026-10-06').serial);
  assert.strictEqual(mdy, Core.toDate('2026-06-10').serial);
});

/* ---------- names ---------- */
test('sheet names are cleaned and made unique', () => {
  const used = new Set();
  assert.strictEqual(Core.cleanSheetName('a/b:c', used), 'a-b-c');
  assert.strictEqual(Core.cleanSheetName('A-B-C', used), 'A-B-C (2)');
  assert.strictEqual(Core.cleanSheetName('x'.repeat(40), used).length, 31);
});
test('file names always end in .xlsx', () => {
  assert.strictEqual(Core.cleanFileName('report'), 'report.xlsx');
  assert.strictEqual(Core.cleanFileName('a/b.xlsx'), 'a-b.xlsx');
  assert.strictEqual(Core.cleanFileName(''), 'converted-data.xlsx');
});

/* ---------- workbook output ---------- */
test('workbook keeps types: numbers, dates, booleans, text phones', async () => {
  const r = proc('Phone,Amt,When,Flag\n09426011223,99.50,2026-10-06,yes\n9876543210,"1,02,000",06/10/2026,no');
  const out = await build({ headers: r.headers, types: r.inferred, rows: r.rows });
  fs.writeFileSync(path.join(OUT, 'types.xlsx'), out.bytes);
  const ws = readBack(out.bytes).Sheets.Sheet1;
  assert.strictEqual(ws.A2.t, 's'); assert.strictEqual(ws.A2.v, '09426011223');
  assert.strictEqual(ws.B3.t, 'n'); assert.strictEqual(ws.B3.v, 102000);
  assert.strictEqual(ws.C2.t, 'n'); assert.strictEqual(ws.C2.v, 46301); // 2026-10-06
  assert.strictEqual(ws.C3.v, 46301);
  assert.strictEqual(ws.D2.t, 'b'); assert.strictEqual(ws.D2.v, true);
});
test('freeze pane, autofilter and column widths are written', async () => {
  const out = await build({ headers: ['a', 'b'], types: ['text', 'number'], rows: [['x', '1'], ['y', '2']] });
  const zip = await JSZip.loadAsync(out.bytes);
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  assert.ok(/<pane ySplit="1"[^>]*state="frozen"/.test(xml), 'frozen pane');
  assert.ok(/<autoFilter ref="A1:B3"/.test(xml), 'autofilter');
  assert.ok(/<cols>/.test(xml), 'column widths');
});
test('Excel Table output is valid and replaces the sheet filter', async () => {
  const out = await build({ headers: ['Name', 'Name'], types: ['text', 'text'], rows: [['a', 'b']] }, { fmt: Object.assign({}, FMT, { table: true }) });
  fs.writeFileSync(path.join(OUT, 'table.xlsx'), out.bytes);
  const zip = await JSZip.loadAsync(out.bytes);
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const table = await zip.file('xl/tables/table1.xml').async('string');
  const ct = await zip.file('[Content_Types].xml').async('string');
  const rels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
  assert.ok(sheet.includes('<tableParts count="1">'));
  assert.ok(!/<autoFilter/.test(sheet), 'no sheet-level filter alongside a table');
  assert.ok(table.includes('name="Name (2)"'), 'duplicate headers made unique');
  assert.ok(ct.includes('/xl/tables/table1.xml'));
  assert.ok(rels.includes('../tables/table1.xml'));
});
test('split by column value creates one sheet per value', async () => {
  const out = await build({ headers: ['N', 'City'], types: ['text', 'text'], rows: [['a', 'Patan'], ['b', 'Siddhpur'], ['c', 'Patan'], ['d', '']] }, { split: { mode: 'column', colIndex: 1 } });
  const wb = readBack(out.bytes);
  assert.deepStrictEqual(wb.SheetNames, ['Patan', 'Siddhpur', '(blank)']);
  assert.strictEqual(XLSX.utils.sheet_to_json(wb.Sheets.Patan).length, 2);
});
test('split by row count', async () => {
  const rows = Array.from({ length: 25 }, (_, i) => [String(i)]);
  const out = await build({ headers: ['n'], types: ['text'], rows }, { sheetName: 'Data', split: { mode: 'rows', rows: 10 } });
  assert.deepStrictEqual(readBack(out.bytes).SheetNames, ['Data 1', 'Data 2', 'Data 3']);
});
test('header-only data still produces a valid workbook', async () => {
  const out = await build({ headers: ['a'], types: ['text'], rows: [] }, { fmt: Object.assign({}, FMT, { table: true }) });
  assert.deepStrictEqual(readBack(out.bytes).SheetNames, ['Sheet1']);
});
test('fixture files parse', () => {
  const csv = proc(fs.readFileSync(path.join(__dirname, 'fixtures/employees.csv'), 'utf8'));
  assert.strictEqual(csv.delimKey, 'comma'); assert.ok(csv.rows.length >= 10);
  const txt = proc(fs.readFileSync(path.join(__dirname, 'fixtures/inventory.txt'), 'utf8'));
  assert.strictEqual(txt.delimKey, 'pipe'); assert.ok(txt.rows.length >= 5);
});
test('large dataset (60,000 rows) parses and builds', async () => {
  let big = 'Order No,Customer,Amount,Date,City\n';
  for (let i = 0; i < 60000; i++) big += `${100000 + i},Customer ${i},${(i * 1.5).toFixed(2)},2025-01-${String((i % 28) + 1).padStart(2, '0')},City${i % 40}\n`;
  let t = Date.now();
  const r = proc(big);
  const parseMs = Date.now() - t;
  assert.strictEqual(r.rows.length, 60000);
  t = Date.now();
  const out = await build({ headers: r.headers, types: r.inferred, rows: r.rows });
  console.log(`      parse ${parseMs} ms, build ${Date.now() - t} ms, ${(out.bytes.length / 1048576).toFixed(1)} MB`);
  fs.writeFileSync(path.join(OUT, 'large.xlsx'), out.bytes);
});

(async () => {
  for (const { name, fn } of tests) {
    try { await fn(); passed++; console.log(`  ✓ ${name}`); }
    catch (e) { failed++; console.log(`  ✗ ${name}\n      ${e && e.message}`); }
  }
  console.log(`\n${passed} passed, ${failed} failed`);
  process.exit(failed ? 1 : 0);
})();
