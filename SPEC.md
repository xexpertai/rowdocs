# rowdocs: SPEC

## 1. Problem
Clients often ask freelancers for the same job: "I have a spreadsheet of N people/orders and a Word template. Make one letter, certificate, invoice or contract per row, in Word and PDF." Most people do this by hand or with fragile Word mail-merge. rowdocs does it reliably from the command line or a small local web page. It checks the data before it generates anything.

**Different from the existing portfolio:**
- sheet-tidy cleans data.
- sheet-showcase publishes data to the web.
- postmeta-audit is a WordPress SEO plugin.
- rowdocs generates **documents** (Word/PDF) from data. That fills the "Word" and "PDF" gaps and adds to "Data Entry" and "Data Processing".

## 2. Scope (finishable in about 4 hours)
- Language: Python 3.11+.
- Dependencies: `openpyxl`, `docxtpl` (python-docx + Jinja2), `pypdf`, `flask`.
- PDF conversion uses LibreOffice headless (`soffice`) **if it is installed**. It is optional and detected at runtime.
- No network calls, no external APIs, no telemetry.
- Licence: MIT. Author: Harsha Shinde. The README has an "AI-assisted development" disclosure line.
- Sample data: fictional, and labelled "Demo data (fictional)" in the README and in the sample files.

## 3. Features

### 3.1 Inputs
- Data: `.xlsx` (first sheet, or `--sheet NAME`) or `.csv` (UTF-8 / UTF-8-BOM, auto-detects `,` `;` `\t`).
- Template: `.docx` with Jinja placeholders such as `{{ name }}`, `{{ amount }}` and `{% if vip %}…{% endif %}`. Placeholders can be in the body, in tables, and in headers/footers.
- Header normalisation: column headers are turned into safe keys (trim, lowercase, spaces/punctuation → `_`). Example: "Full Name" → `full_name`. The original → key mapping is shown in validation output.

### 3.2 Filters (usable in templates)
- `date(fmt)`: formats Excel dates or ISO strings, e.g. `{{ issued|date('%d %b %Y') }}`.
- `money(symbol, decimals=2)`: thousands separators, e.g. `{{ amount|money('₹') }}`.
- `upper`, `title`, and `default('…')` (built into Jinja).

### 3.3 Pre-flight validation (`rowdocs check`)
- Lists every placeholder variable in the template, including those in headers/footers.
- Reports:
  - (a) **ERROR**: template variables with no matching column;
  - (b) WARN: columns the template never uses;
  - (c) WARN, listed per row: rows where a used variable is empty;
  - (d) WARN: duplicate output file names that the naming pattern would produce.
- Exit code 0 if there are no errors, 2 if there are errors. `--strict` also turns warnings into exit code 2.
- Output: a human-readable table on stdout, and `--json` for machine-readable output.

### 3.4 Generation (`rowdocs build`)
- `rowdocs build DATA TEMPLATE -o OUTDIR [--name "{{ invoice_no }}_{{ full_name }}"] [--pdf] [--merge-pdf] [--zip] [--rows 2-10,15] [--where "status=paid"]`
- One `.docx` per selected row. File names come from the `--name` Jinja pattern (default `row_{{ _row }}`). Names are sanitised: no path separators, max 120 chars. Collisions get a suffix `_2`, `_3`.
- `--pdf`: converts each DOCX to PDF via `soffice --headless --convert-to pdf`, batched in one soffice call. If soffice is missing, it prints a clear message and exits with code 3. DOCX output is kept.
- `--merge-pdf`: also writes `ALL.pdf` with every PDF in row order (pypdf).
- `--zip`: writes `OUTDIR.zip` with all outputs plus `manifest.csv`.
- `manifest.csv` is always written. Columns: `row`, `output_docx`, `output_pdf`, `status` (ok/skipped/error), `message`.
- If one row fails, generation continues with the other rows. The error goes into the manifest and the exit code is 1.
- Validation runs automatically before the build and aborts on ERROR, unless `--force` is passed.

### 3.5 Local web UI (`rowdocs serve`)
- Binds to 127.0.0.1 only. The default port is 8765.
- A single page: plain HTML/CSS/vanilla JS, responsive, no CDN.
  1. Upload the data file and the template.
  2. Show the validation result: errors in red, warnings in amber, and the column↔placeholder mapping.
  3. Preview row 1 rendered as text (paragraph text extracted from the generated DOCX).
  4. Options: naming pattern, PDF, merged PDF.
  5. "Generate" downloads a ZIP.
- Upload limit: 20 MB. Only .xlsx, .csv and .docx are accepted. Each request uses its own temp directory, which is deleted after the response.

### 3.6 Samples
`examples/` contains:
- `certificates.xlsx` (12 fictional rows) and `certificate.docx`;
- `invoices.csv` (8 rows, including one with a missing amount, to show warnings) and `invoice.docx`, which uses a table, the `money` filter and an `{% if %}` block.

## 4. Acceptance criteria
1. `pip install -e .` provides a `rowdocs` command with `check`, `build` and `serve` subcommands. `rowdocs --help` lists them.
2. `rowdocs check examples/invoices.csv examples/invoice.docx`:
   - exits with 0;
   - prints at least one WARN naming the row with the missing amount.
3. A template with an `{{ unknown_field }}` placeholder makes `check` print an ERROR naming `unknown_field` and exit with 2. `build` on the same files aborts with no output, unless `--force` is passed.
4. `rowdocs build examples/certificates.xlsx examples/certificate.docx -o out --name "{{ full_name }}"`:
   - produces exactly 12 `.docx` files plus `manifest.csv`;
   - each file opens in python-docx and contains that row's name;
   - no `{{` or `}}` is left in any output.
5. Placeholders inside a table cell and inside the page header are replaced (covered by a test).
6. `money('₹')` renders 1234567.5 as `₹1,234,567.50`. `date('%d %b %Y')` renders an Excel date cell and the string `2026-03-05` as `05 Mar 2026`.
7. A naming pattern that yields duplicates produces `X.docx` and `X_2.docx`, and `check` warns about it. A name containing `../` cannot write outside OUTDIR (test).
8. `--rows 2-4` generates 3 files. `--where "status=paid"` generates only the matching rows.
9. With soffice available, `--pdf --merge-pdf` produces one PDF per row plus `ALL.pdf`, whose page count equals the sum of the parts. Without soffice, `--pdf` exits with 3 and a message saying how to install LibreOffice. PDF tests are skipped automatically when soffice is absent.
10. `--zip` produces a ZIP containing every generated file and `manifest.csv`.
11. If one row fails to render (e.g. a filter gets bad input), the other rows still generate, the manifest marks that row `error` with a message, and the exit code is 1.
12. Web UI:
    - `rowdocs serve` listens only on 127.0.0.1;
    - uploading the sample invoice files shows the warnings and the preview;
    - "Generate" downloads a valid ZIP;
    - a `.exe` upload is rejected with HTTP 400 (Flask test client);
    - temp directories are gone after the request (test).
13. `pytest` passes with ≥ 20 tests and no network access. `ruff check` is clean.
14. README:
    - one-paragraph pitch, install, a quick start with the sample commands, a template syntax cheat sheet, the filter list, and an FAQ ("Why are my placeholders not replaced?" covers Word splitting text runs; fix: retype the placeholder in one go);
    - a screenshot of the web UI;
    - labels: "Demo data (fictional)" and "AI-assisted development".
15. The repo contains no secrets, no personal data beyond the author name "Harsha Shinde", and no third-party trademarks in its sample documents.

## 5. Out of scope
Emailing documents, cloud storage, e-signatures, `.doc` (legacy) templates, image placeholders.