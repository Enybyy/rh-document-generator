<div align="center">

# Generar RH

Batch Word documents from tabular data and a template, with row review before file generation.

<a href="https://enybyy.github.io/rh-document-generator/"><img src="docs/media/demo.svg" width="360" alt="Open demo"></a>

<p><a href="https://github.com/Enybyy"><img src="docs/media/github.svg" width="112" alt="Eliud Rojas Mendoza on GitHub"></a>
<a href="https://www.linkedin.com/in/eliud-rojas-mendoza-414652212/"><img src="docs/media/linkedin.svg" width="112" alt="Eliud Rojas Mendoza on LinkedIn"></a>
<a href="https://www.upwork.com/freelancers/~01471ca462b236e8e5"><img src="docs/media/upwork.svg" width="112" alt="Eliud Rojas Mendoza on Upwork"></a></p>

[![Generar RH in use](assets/screenshots/rh-desktop.png)](https://enybyy.github.io/rh-document-generator/)

*Actual Generar RH screenshot with fictional sample data.*

[About](#about-the-project) · [Workflow](#everyday-workflow) · [Technology](#built-with) · [Run locally](#local-use)

</div>


## About the project

Preparing personnel documents often repeats the same structure for different records: change names, dates and amounts, then check that formatting remains consistent. Generar RH brings that preparation into a batch, connecting table data to fields in a Word template.

Review makes incomplete rows visible before generation. Once records are selected, each person receives an editable file, accompanied by a batch summary. The demo follows the workflow with a sample template; the Python application accepts Excel files and custom templates.

## Everyday workflow

| Inside the project | Detail |
| --- | --- |
| Data and template | Match Excel columns to tags in a Word document. |
| Batch review | Validate each row and select records before generation. |
| Editable documents | Replace tags while preserving supported DOCX formatting. |
| Batch delivery | ZIP download, individual documents and generation summary. |
| Two ways to explore | Browser demo with a sample template; local app with custom templates. |

## Explore the workflow

1. Download the sample Excel file and template from the application, or prepare your own files for the local version.
2. Use the first XLSX sheet with headers in its first row. Add tags such as `[NOMBRE]`, `[DNI]` or `{{ FECHA_INICIO }}` to Word.
3. Upload XLSX and DOCX files. An optional employee database with unique DNIs can complete or replace fields for the matching DNI.
4. Review row errors and select valid records.
5. Download a ZIP with one Word file per record, `revision.xlsx` and `resumen.json`.

Tags match normalized headers: uppercase, without accents, with spaces replaced by `_`. Existing aliases remain `NUMERO DE DOCUMENTO → DNI`, `APELLIDOS Y NOMBRES → NOMBRE`, `SALARIO → PAGO` and `FECHA DE INICIO → FECHA_INICIO`. Other columns and matching tags can also be used. DNI checks validate an eight-digit format, not real identity. Keep identifiers as text in Excel; an integer numeric DNI is padded to eight digits.

The engine replaces tags even when Word splits them across formatting runs. It retains the first run's formatting for each tag and preserves surrounding text. Supported content includes paragraphs, nested tables, headers and footers, including first-page and even-page variants. The included sample has two navigable pages. After generation, browse contracts and download individual Word files or the full ZIP. Custom templates show a field summary; open the generated DOCX in Word to view its original page layout.

## Public demo and local application

| Capability | GitHub Pages demo | Python application |
| --- | --- | --- |
| Edit sample data and select rows | Yes | Yes |
| Generate downloadable Word files | Yes, fixed sample template | Yes, custom template |
| Import data | CSV | XLSX |
| Review report | CSV inside the ZIP | XLSX inside the ZIP |
| Join by DNI | No | Optional employee database |
| Processing | Browser, no external calls | Request memory, uploads not saved |

RH refers to human resources. The demo uses fictional data and generates actual files; it does not query identity records or issue tax receipts. The sample contract needs company details, RUC, representative, scope and payment terms completed before use.

![Pages of generated contracts](assets/screenshots/rh-generated.png)

## Optional PDF receipt extraction

The function `receipts.extract_receipt(bytes)` and `POST /api/receipt` retain local extraction of series, date and total from the earlier workflow. The multipart field is named `pdf`. It handles text-containing PDFs and reports missing fields for manual review. It does not authenticate tax documents, download from Drive, require credentials or perform OCR. Different layouts require adjusted patterns.

## Limits and errors

- 500 rows, 100 columns, 10 MB per file, 25 MB per request and 64 MB of generated documents per batch.
- Rejects formulas, empty or duplicate headers, corrupt documents, macros and excessively large unpacked archives.
- Duplicate DNIs in the employee database block ambiguous joins. Unmatched rows remain pending.
- Invalid dates, an end date before the start, invalid amounts and missing tag values exclude a row. Amounts use a decimal point, up to two decimals and no thousands separators.
- Filenames are sanitized and indexed to prevent overwriting. Spreadsheet reports neutralize formula prefixes.
- Replacement is limited to supported Word content; it does not interpret native mail-merge fields or update Word indexes or formulas.
- Primary output is DOCX. Export PDF from your editor if needed; Word is not required to generate files.

## Built with

| Area | Technology |
| --- | --- |
| Local application | Python, Flask and Waitress |
| Documents and spreadsheets | python-docx and openpyxl |
| Demo | HTML, CSS and JavaScript; browser-based DOCX generation |
| Verification | pytest and Playwright |

## Local use

<details>
<summary><strong>Run on your computer</strong></summary>

Python 3.12+ is required. From the project folder:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5081`. On Linux/macOS, activate with `source .venv/bin/activate`. Waitress listens on the local computer only. Deployments handling real personnel files need authentication, HTTPS and appropriate concurrency limits.

</details>

<details>
<summary><strong>Verification</strong></summary>

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Tests cover replacement across runs, formatting, nested tables, headers/footers, selection, joins, dates, amounts, duplicate filenames, file errors and API behavior. `scripts/browser-test.cjs` checks the interface and captures it with Playwright, using the local server on 5081 and static server on 5085.

</details>

<details>
<summary><strong>Origin and consolidation</strong></summary>

Rebuilt from the Excel-to-Word generator in `business-automation-suite`, with review and selection added. Another variant extracted PDFs from Drive; its local extraction remains in `receipts.py`, while the specific integration was retired. `contract-automation-system` had template/source uploads but no complete generation engine and was retired in favor of this project. See [the comparison](docs/CONSOLIDATION.md).

The DNI and keyboard projects have independent repositories: [DNI Identity Validator](https://github.com/Enybyy/dni-identity-validator) and [Keyboard Event Lab](https://github.com/Enybyy/keyboard-event-lab).

</details>

---

<div align="center">

**Eliud Rojas Mendoza · Enybyy**

<p><a href="https://github.com/Enybyy"><img src="docs/media/github.svg" width="112" alt="Eliud Rojas Mendoza on GitHub"></a>
<a href="https://www.linkedin.com/in/eliud-rojas-mendoza-414652212/"><img src="docs/media/linkedin.svg" width="112" alt="Eliud Rojas Mendoza on LinkedIn"></a>
<a href="https://www.upwork.com/freelancers/~01471ca462b236e8e5"><img src="docs/media/upwork.svg" width="112" alt="Eliud Rojas Mendoza on Upwork"></a></p>

</div>
