/*
 * Browser end-to-end check (optional; needs Playwright + Chromium):
 *   PLAYWRIGHT_BROWSERS_PATH=... node tests/e2e.mjs [baseUrl]
 * Without a baseUrl it serves ./public locally with the same headers as vercel.json.
 */
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(process.env.PW_REQUIRE_FROM || import.meta.url);
const { chromium } = require('playwright');

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const pub = path.join(root, 'public');
const outDir = path.join(root, 'tests', 'output');
fs.mkdirSync(outDir, { recursive: true });

let server, base = process.argv[2];
if (!base) {
  const vercel = JSON.parse(fs.readFileSync(path.join(root, 'vercel.json'), 'utf8'));
  const globalHeaders = vercel.headers.find(h => h.source === '/(.*)').headers;
  const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.woff2': 'font/woff2' };
  server = http.createServer((req, res) => {
    let p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (p === '/') p = '/index.html';
    const f = path.join(pub, p);
    if (!f.startsWith(pub) || !fs.existsSync(f)) { res.writeHead(404); return res.end('not found'); }
    for (const h of globalHeaders) if (h.key !== 'Strict-Transport-Security') res.setHeader(h.key, h.value.replace('; upgrade-insecure-requests', ''));
    res.setHeader('Content-Type', types[path.extname(f)] || 'application/octet-stream');
    res.end(fs.readFileSync(f));
  });
  await new Promise(r => server.listen(0, r));
  base = `http://127.0.0.1:${server.address().port}`;
}

const results = [];
const check = (name, ok, detail = '') => { results.push({ name, ok }); console.log(`  ${ok ? '✓' : '✗'} ${name}${detail ? ` (${detail})` : ''}`); };

const browser = await chromium.launch(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {});
try {
  const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1366, height: 900 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('pageerror', e => errors.push(String(e)));
  await page.goto(base + '/', { waitUntil: 'networkidle' });

  check('page title', (await page.title()) === 'Text to Excel Converter');
  check('libraries loaded', await page.evaluate(() => typeof XLSX === 'object' && typeof JSZip === 'function' && typeof Core === 'object'));
  check('sample data parsed on load', (await page.textContent('#stRows')) === '10', await page.textContent('#dims'));
  check('delimiter shown', (await page.textContent('#delimInfo')).includes('Comma'));
  check('fonts loaded', await page.evaluate(async () => { await document.fonts.ready; return document.fonts.check('14px "IBM Plex Sans"'); }));
  await page.screenshot({ path: path.join(outDir, 'desktop-light.png'), fullPage: true });

  // Convert & download the sample
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#convertBtn')]);
  const saved = path.join(outDir, 'e2e-' + dl.suggestedFilename());
  await dl.saveAs(saved);
  check('download fires with .xlsx name', dl.suggestedFilename() === 'converted-data.xlsx');
  check('download is a zip/xlsx', fs.readFileSync(saved).subarray(0, 2).toString() === 'PK', `${fs.statSync(saved).size} bytes`);
  check('status shows Converted', (await page.textContent('#statusText')) === 'Converted');

  // Empty input
  await page.click('#qaClear');
  await page.click('#convertBtn');
  const emptyErr = await page.waitForFunction(() => document.querySelector('#alerts').textContent.includes('Please enter or upload some data first.'), null, { timeout: 5000 }).then(() => true, () => false);
  check('empty input error', emptyErr);

  // Upload .csv and .txt
  await page.setInputFiles('#fileInput', path.join(root, 'tests/fixtures/employees.csv'));
  await page.waitForFunction(() => document.querySelector('#stRows').textContent === '10');
  check('CSV upload parsed', true, await page.textContent('#fileName'));
  await page.setInputFiles('#fileInput', path.join(root, 'tests/fixtures/inventory.txt'));
  await page.waitForFunction(() => document.querySelector('#delimInfo').textContent.includes('Pipe'));
  check('TXT upload detects pipe', (await page.textContent('#stRows')) === '5');
  check('SKU with leading zeros stays text', (await page.locator('.data-grid thead th', { hasText: 'SKU' }).textContent()).includes('text'));

  // Tab-separated paste + dedupe toggle
  await page.fill('#textIn', 'Name\tAge\tCity\nAli\t25\tPatan\nAli\t25\tPatan\nRahul\t30\tAhmedabad');
  await page.waitForFunction(() => document.querySelector('#delimInfo').textContent.includes('Tab'));
  check('TSV duplicate counted', (await page.textContent('#stDup')) === '1');
  await page.check('#clDupes');
  check('dedupe removes row', (await page.textContent('#stRows')) === '2');

  // Column management
  await page.fill('#newColName', 'Country');
  await page.fill('#newColDefault', 'India');
  await page.click('#addColForm button[type=submit]');
  check('add column', (await page.locator('.col-row').count()) === 4);
  await page.locator('.col-row').nth(0).locator('button[data-act="down"]').click();
  check('reorder column', (await page.locator('.col-row input').nth(1).inputValue()) === 'Name');
  await page.locator('.col-row').nth(0).locator('button[data-act="del"]').click();
  check('delete column', (await page.locator('.col-row').count()) === 3);

  // Split by column, Excel Table on, download and inspect
  await page.click('#splitPanel summary');
  await page.click('label:has(input[name="splitMode"][value="column"])');
  await page.selectOption('#splitCol', { label: 'City' });
  await page.check('#fxTable');
  const [dl2] = await Promise.all([page.waitForEvent('download'), page.click('#qaDownload')]);
  const saved2 = path.join(outDir, 'e2e-split.xlsx');
  await dl2.saveAs(saved2);
  const sheetNames = await page.evaluate(async (b64) => {
    const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    return XLSX.read(bytes, { type: 'array' }).SheetNames;
  }, fs.readFileSync(saved2).toString('base64'));
  check('split workbook sheets', JSON.stringify(sheetNames) === JSON.stringify(['Patan', 'Ahmedabad']), sheetNames.join(', '));

  // Large paste stays responsive (pagination)
  const big = ['ID,Name,Amount'].concat(Array.from({ length: 30000 }, (_, i) => `${i},Person ${i},${i * 3}`)).join('\n');
  await page.evaluate(t => { const ta = document.querySelector('#textIn'); ta.value = t; ta.dispatchEvent(new Event('input')); }, big);
  await page.waitForFunction(() => document.querySelector('#stRows').textContent === '30,000', null, { timeout: 15000 });
  check('large data paginated', (await page.locator('.data-grid tbody tr').count()) === 25, await page.textContent('#pageInfo'));

  // Reset settings
  await page.click('#qaReset');
  check('reset restores defaults', !(await page.isChecked('#fxTable')) && (await page.inputValue('#sheetName')) === 'Sheet1');

  // Dark mode toggle
  await page.click('#qaSample');
  await page.click('#themeBtn');
  const darkBg = await page.evaluate(() => getComputedStyle(document.body).backgroundColor);
  check('dark mode applies', darkBg === 'rgb(14, 20, 17)', darkBg);
  await page.screenshot({ path: path.join(outDir, 'desktop-dark.png'), fullPage: false });
  check('no console errors', errors.length === 0, errors.slice(0, 3).join(' | '));

  // Mobile layout
  const m = await browser.newContext({ viewport: { width: 390, height: 844 }, isMobile: true, hasTouch: true, deviceScaleFactor: 2 });
  const mp = await m.newPage();
  await mp.goto(base + '/', { waitUntil: 'networkidle' });
  const overflow = await mp.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
  check('mobile: no horizontal page scroll', overflow <= 0, `overflow ${overflow}px`);
  check('mobile: sticky convert bar visible', await mp.isVisible('#mobileConvert'));
  await mp.screenshot({ path: path.join(outDir, 'mobile-light.png'), fullPage: false });
  await m.close();
} finally {
  await browser.close();
  if (server) server.close();
}
const failed = results.filter(r => !r.ok).length;
console.log(`\n${results.length - failed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
