# Text to Excel Converter

Paste or upload text data and download a clean, formatted Excel (`.xlsx`) workbook. Everything runs in the browser: your data never leaves your device.

## Features

- **Input**: paste text, drag and drop a file, or choose a `.txt`, `.csv`, `.tsv` or `.log` file. Shows character and line counts.
- **Delimiter detection**: comma, tab, semicolon, pipe and colon are detected automatically. You can also pick one manually or type a custom delimiter (multi-character delimiters such as `||` work). Quoted fields that contain delimiters, escaped quotes or line breaks are handled.
- **Headers**: use the first row as column headers, or generate `Column 1`, `Column 2`, and so on.
- **Cleaning**: each of these is optional:
  - remove empty rows and empty columns
  - trim whitespace
  - remove duplicate rows
  - normalize line breaks
- **Type detection**: numbers (including `1,250,000` and Indian `12,50,000` grouping), dates (`2026-10-06`, `06/10/2026`, `06-10-2026`, with a day/month order setting) and booleans (`true/false/yes/no`) are detected. Phone numbers, PIN/ZIP codes, IDs and values with leading zeros stay as text.
- **Column management**: rename, reorder, delete and add columns, and override a column's type (Text, Number, Date, Boolean, Auto).
- **Preview**: paginated preview with spreadsheet-style column letters and row numbers. Handles tens of thousands of rows.
- **Statistics**: counts of rows, columns, empty rows, duplicate rows, and the data size.
- **Excel output**:
  - custom sheet and file names
  - bold header, frozen header row, auto column widths, filters
  - optional native Excel Table and alternate row shading
  - real number, date and boolean cells
- **Multiple sheets**: split by a fixed number of rows, or one sheet per value of a column (for example, one sheet per City).
- **Extras**: dark mode, copy data as tab-separated text, reset settings, and a mobile-friendly layout.

## Privacy

The site is fully static. It makes no API calls and has no backend, database or analytics. Fonts and libraries are self-hosted. A strict Content-Security-Policy (`connect-src 'self'`) is set in `vercel.json`. Uploaded files are read with `FileReader` and are never stored.

## Tech stack

- HTML5, CSS3, vanilla JavaScript (no framework)
- [Tailwind CSS](https://tailwindcss.com) 3, compiled to a static stylesheet
- [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style) 1.2.0: SheetJS (XLSX 0.18.5) with cell styling support
- [JSZip](https://stuk.github.io/jszip/) 3.10.1: adds frozen panes and Excel Table parts to the SheetJS output
- IBM Plex Sans / Mono (self-hosted woff2)

## Project structure

```text
public/                  Deployed site (Vercel output directory)
  index.html
  favicon.svg
  assets/css/styles.css  Built by Tailwind (committed, so the site also works without a build)
  assets/js/core.js      Parsing, cleaning, type detection, workbook building (no DOM)
  assets/js/app.js       User interface
  assets/js/theme.js     Applies the saved theme before first paint
  assets/vendor/         xlsx-js-style and JSZip
  assets/fonts/          IBM Plex woff2 files
src/input.css            Tailwind source: design tokens and components
tests/run-tests.js       Unit tests for core.js (Node, no dependencies)
tests/e2e.mjs            Optional browser test (Playwright + Chromium)
tests/fixtures/          Sample .csv and .txt files
vercel.json              Build settings and security headers
```

## Development

```bash
npm install          # installs Tailwind CSS (the only dev dependency)
npm run build        # compiles src/input.css -> public/assets/css/styles.css
npm run watch:css    # rebuild CSS on change
npm run serve        # serve ./public at http://localhost:3000
npm test             # unit tests
```

You can also open `public/index.html` directly in a browser. Everything works from the file system, except that some browsers block the web fonts there.

Browser test (requires Playwright and a Chromium build):

```bash
npm i -D playwright && npx playwright install chromium
node tests/e2e.mjs                       # against a local server with production headers
node tests/e2e.mjs https://your-site.vercel.app
```

## Deployment

The repository is connected to Vercel. Every push to `main` deploys to production. Vercel runs `npm run build` and serves the `public/` directory. See `vercel.json` for the build settings and headers.

## License

MIT
