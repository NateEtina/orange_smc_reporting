# CSV → Excel Converter + Weekly Report (Hebdo)

A web app that converts CSV files into richly formatted Excel workbooks — built with
**FastAPI** + **openpyxl**, with a single-page UI served from the same process.

It also ships **Feature 1 — Excel & PowerPoint report generation for tickets and
incidents (P1/P2 MTTR)**: point it at the two weekly sources (`brise_smc_weekXX`
and `GNOC_weekXX`) and it computes the metrics and fills the `hebdo_SXX` template
(`hebdo.py`, see below).

Everything runs locally: files never leave your machine.

## Quick start

```bash
pip install -r requirements.txt
python main.py
```

Open <http://127.0.0.1:8000>, drop your CSV files, pick options, convert, download.

Windows: double-click `run.bat`.

## What "all details possible" means here

### Reading the CSV faithfully
| Detail | Behaviour |
|---|---|
| Encoding | BOM (UTF-8/16/32) honoured, otherwise charset-normalizer detection, then safe fallbacks (UTF-8 → CP1252 → Latin-1). Reported with a confidence score. |
| Delimiter | `csv.Sniffer` + frequency analysis for `,` `;` tab `\|`, or force one (incl. custom). |
| Quoting | Standard CSV quoting, `""` escapes; double- or single-quote character selectable. |
| Ragged rows | Short rows are padded, over-long rows kept — nothing is dropped. Reported. |
| Empty rows | Skipped (toggleable); empty cells stay genuinely empty. |
| Header | Sanitised: trimmed, de-duplicated (`name`, `name (2)`), blanks → `Column N`. |
| Whitespace | Preserved by default; optional trimming. |

### Type inference (per column, 90% confidence over a 5,000-row sample)
- **Integers / decimals** with automatic separator style detection:
  `1,234.56` (US), `1.234,56` (EU), `1 234,56` (space grouping), scientific notation.
- **Currency**: `$ € £ ¥ ₹ ₩ ₽ …` and ISO codes (`USD 1,200.00`), prefix or suffix →
  numeric cell with a `"€"#,##0.00`-style format.
- **Percent**: `45%` → `0.45` with a `0.00%` format.
- **Booleans**: `true/false`, `yes/no`, `on/off`.
- **Dates & times**: ISO 8601 (`2024-01-15`, `2024-01-15T09:30:00Z`), `YYYY-MM-DD`,
  `DD/MM/YYYY` vs `MM/DD/YYYY` (auto-detected from the data, overridable), `15 Jan 2024`,
  `Jan 15, 2024`, `YYYYMMDD`, plus time-only columns (`09:30`, `1:30 PM`).
  Timezone offsets are dropped but the wall-clock time is kept.
- **Leading zeros** (`007`, `01234`) stay text — never silently turned into 7 / 1234.
- **16+-digit IDs** stay text (Excel only keeps 15 significant digits).
- Mixed columns that don't fit any type stay **text**, exactly as written.

### Excel output
- Per-column number formats (`#,##0.00`, `0.00%`, currency, `yyyy-mm-dd hh:mm:ss`).
- Styled header (bold + fill/text colours), frozen top row, auto-filter buttons,
  auto-sized columns (wide characters count double), optional borders, zebra stripes, wrap text.
- Sheet names sanitised (invalid characters, 31-char limit, collisions).
- Rows beyond **1,048,576** automatically continue on `name (2)`, `name (3)`, …
- Cells beyond **32,767** characters truncated (and reported).
- Formula injection (`=`, `+`, `-`, `@` payloads) neutralised as plain text by default.
- Optional **Details** sheet: file info, encoding, delimiter, and per-column type,
  number format, non-empty/empty counts, unique count, min, max, examples.

### Output modes
- **One workbook** — every CSV becomes its own sheet (default).
- **Separate workbooks** — delivered as a `.zip` (plus `conversion-report.json`).

### Monitoring profile (key=value tags)
A single switch (options → *Monitoring profile*) that understands structured
monitoring/test-result exports and produces a ready-made French report:

- **Wrapped-line repair** — files whose lines are each one quoted field with
  doubled quotes (`"a,""b"",1"`) are detected and unwrapped back into columns.
- **`tags` parsing** — a column named `tags` holding `key1=value1,key2=value2`
  is exploded into one independent column per key (escaped commas `c=x\,y` stay
  inside the value; escape characters such as `\ ` are removed → space). A key
  colliding with an existing column becomes `<key> (2)`; the raw `tags` column
  is dropped.
- **`time` conversion** — a column named `time` holding a UNIX timestamp
  (nanoseconds; microseconds/milliseconds/seconds auto-detected by magnitude)
  becomes `Horodatage (UTC)` (`YYYY-MM-DD HH:MM:SS`, UTC) plus a derived
  `Date` column (`YYYY-MM-DD`). Unparsable values stay unchanged and are
  reported.
- **Renaming to the French report headers** — sources (column names *and* tag
  keys, case-insensitive) are mapped: `zone → Zone Géographique`,
  `status → Statut Global`, `monitoring_name → Nom du Service (Monitoring)`,
  `test_name → Nom du Test`, `failure_reason → Raison de l'Échec`,
  `count_duration_seconds → Durée (sec)`, `cellular_ussd_status → Statut USSD`,
  `cellular_sms_reception_status → Statut SMS`,
  `failed_action_kind → Type d'Action en Échec`, `monitoring_id → ID Monitoring`,
  `test_id → ID Test`, `product_name → Nom du Produit`, `product_id → ID Produit`,
  `platform_id → ID Plateforme`, `name → Type Enregistrement`, plus English
  synonyms (`Geographic Zone`, `Overall Status`, `Service Name`…). Unknown names
  are kept as-is.
- **Standard column order** — `Date, Horodatage (UTC), Zone Géographique,
  Statut Global, Nom du Service (Monitoring), Nom du Test, Raison de l'Échec,
  Durée (sec), Statut USSD, Statut SMS, Type d'Action en Échec, ID Monitoring,
  ID Test, Nom du Produit, ID Produit, ID Plateforme, Type, Type Enregistrement`,
  then any remaining columns.
- **Descriptive labels** — empty cells become `Non Définie` (Zone Géographique),
  `Aucune (Succès)` (Raison de l'Échec) and `N/A` (Statut SMS,
  Type d'Action en Échec).
- **Mandated Excel look** — worksheet named **`Detailed Raw Data`**, midnight-blue
  header `#1F4E78` (bold, white, Segoe UI 11), zebra stripes `#F9FAFB`, status
  cells colored and bold (Segoe UI 9: `success` `#E2EFDA`, `failed` `#FCE4D6`,
  `error` `#FFF2CC` — matches `Statut Global`), auto column widths, header
  auto-filter, first row frozen.

Works for one or many uploads: one workbook (each CSV a `Detailed Raw Data`
sheet, numbered when several share a workbook) or one workbook per CSV (zip).

## Feature 1 · Weekly Excel & PowerPoint report (Hebdo)

UI at <http://127.0.0.1:8000/hebdo> (nav link from the main page). The engine lives
in `hebdo.py`; the base template ships as `templates/hebdo_S38.xlsx`.

Both pages ship a **light/dark theme switcher** (top bar) — the choice is persisted
in `localStorage` and restored before the first paint.

### Business rules implemented
| Rule | Behaviour |
|---|---|
| **Dynamic week** | Explicit `S39` / `39` / `Semaine 39` (API), or any date in the target week (ISO week resolution), or today by default. Output is `hebdo_S39.xlsx` / `hebdo_S39.pptx`, and document title, subject and print header update to the target week. In the **UI the week field is mandatory**: marked with `*`, digits only (`38`, `39` — the `S` prefix and any non-numeric character are stripped on input) and the buttons stay disabled until it holds 1–2 digits. |
| **Week filter** | Source column `semaine creation` matched against the target week — accepts `39`, `S39`, `2026-S39`, `Semaine 39`, plain dates (`2026-09-24`, `24/09/2026`) and native date cells. |
| **brise_smc → `IT` row** | Keeps only active EDS units `505010`, `555402`, `610136`; injected under the `IT` row of all 3 template sheets. |
| **GNOC → sub-domain rows** | Grouped by `valeur champ 2` into `Core CS`, `Core IP`, `Core PS`, `IN/TRANS/VAS`. **`IN/TRANS/VAS` is a composite domain**: the distinct cell values `IN`, `VAS` and `TRANS` (plus combined spellings such as `IN/TRANS/VAS`) all map to that one row — token-based matching, so `INCIDENT` cannot be mistaken for `IN`. Unknown sub-domains are skipped and reported. |
| **Tickets sheet** | `TT` = total; `P1–P4` = SLA compliance rate **per priority** = `GOOD ÷ tickets of that priority` (stored as a fraction with `0.00%` format). **Null vs zero**: if none of that priority's tickets carries an evaluated verdict (`GOOD`/`BAD`) the cell is empty (`-`) — never `0.00%` — while a real evaluated zero stays numeric `0.00%`. `In SLA` = GOOD, `Hors SLA` = BAD, `En cours` = `En Cours` (brise) or blank SLA (GNOC). |
| **MTTR sheets** | Per priority (P1, P2) two columns driven by the **SLA status verdict** — the `SLA` column in brise_smc, spelled **`Statut SLA Retablissement`** in GNOC exports (same role, auto-detected either way). Matching is case-insensitive and trimmed: **brise `Ouvert`** = status `En cours`; **GNOC `Ouvert`** = blank status. **`MTTR`** applies the final formula **Σ `Duree` ÷ count** over the subset of tickets that are (a) of that priority, (b) of that domain, (c) **resolved** (brise: status ≠ `En cours`; GNOC: status non-blank) and (d) **`Duree > 0`** (strict) — a resolved ticket whose `Duree` is blank, unparsable or `≤ 0` is excluded from both the sum and the count (raised as a warning), so an empty subset shows `-`. **When exactly one closed ticket qualifies for a domain, that ticket's `Duree` *is* the MTTR** (sum 1 ÷ count 1). Plain numbers are hours (`36.5`), and `HH:MM` / `HH:MM:SS` strings (lenient single-digit minutes/seconds as exported, e.g. `3:26:0`), compact `5h30` values, trailing-`h` values, Excel time cells and duration cells are auto-detected. The result is expressed in **total hours and displayed as `HH:MM`** — `36.5 h` → `36:30`, values beyond 24 h included; `-` when no ticket qualifies. |
| **MTTR SLA thresholds** | Strictly above the target is **BAD**: P1 `> 4.0 h`, P2 `> 8.0 h` → red fill (Excel & PowerPoint); at or below the target → green. PowerPoint additionally renders flagged cells in **bold red/green text**. |
| **Display convention** | **A calculated zero is shown** — a numeric `0` for counts (a row with tickets but an empty tally) and `0.00%` for a real evaluated zero rate. **No underlying data shows `-`** — an empty domain row, a priority with no ticket at all, an un-evaluated rate or an MTTR with no qualifying ticket. |
| **Sources** | `.csv` (encoding + delimiter auto-detected, reusing the converter's readers) and `.xlsx`; French/accented headers matched accent-insensitively. Column mapping is auto-detected and can be overridden per role (`Duree` is required — it supplies the MTTR values; `SLA` / `Statut SLA Retablissement` supplies the tallies, rates and the open/resolved verdict). |

### Outputs
- **Excel** — the `hebdo_S38` template filled in place (styles preserved), renamed dynamically.
- **PowerPoint** — generated from scratch with python-pptx: title slide + Tickets table +
  MTTR P1 / MTTR P2 tables (same values and colour flags as Excel).
- **Both** — delivered as `hebdo_S39.zip` (xlsx + pptx + `S39-metrics.json`).

Try it with the bundled samples: `sample-data/brise_smc_week39.csv` and
`sample-data/GNOC_week39.csv`.

## API

Interactive docs: <http://127.0.0.1:8000/docs>

| Endpoint | Method | Purpose |
|---|---|---|
| `/` | GET | The converter UI |
| `/hebdo` | GET | The weekly-report UI |
| `/api/preview` | POST | `file` + `options` → detected encoding/delimiter, inferred types, first rows |
| `/api/convert` | POST | `files[]` + `options` → `.xlsx` (or `.zip`); report in the `X-Report` header |
| `/api/hebdo/preview` | POST | `brise` + `gnoc` + `week`/`date` + `columns` → computed metrics JSON (row-by-row, detected columns, warnings) |
| `/api/hebdo/generate` | POST | Same inputs + `output` (`xlsx`\|`pptx`\|`both`) → report file (zip when both) |
| `/api/options` | GET | Default options + available encodings |
| `/api/health` | GET | Liveness check |

Options are a JSON object matching the `Options` dataclass in `converter.py`.
Hebdo column overrides are a JSON object mapping roles (`week`, `eds`, `priority`,
`sla`, `domain`, `duration`) to exact source column names.

## Tests

```bash
python test_converter.py     # 142 end-to-end engine checks (converter)
python test_hebdo.py         # 237 end-to-end checks (weekly report engine)
python check_ui.py           # static UI checks for both pages (JS syntax, wiring, endpoints,
                             #   theme switcher, mandatory digits-only week)
python api_check_hebdo.py    # 31 HTTP checks against a running server (python main.py)
                             #   uses port 8000; override with e.g.
                             #   HEBDO_API_BASE=http://127.0.0.1:8010 when 8000 is busy
```

`test_converter.py` covers encodings, delimiters, every inferred type, styling,
sheet splitting, multi-file modes, formula escaping and edge cases.

`test_hebdo.py` covers week resolution/matching, `Duree` parsing (numbers,
`HH:MM` incl. lenient `3:26:0`, compact `5h30`, Excel time/duration cells), the composite `IN`/`VAS`/`TRANS` domain
mapping, column detection (incl. accents, the GNOC `Statut SLA Retablissement`
header and overrides), both aggregations (filters, SLA tallies, null-vs-zero
rates, the file-specific open/resolved status rules feeding `Ouvert` and the
final MTTR formula (Σ `Duree` ÷ count over resolved tickets with `Duree > 0`,
a single closed ticket taking its own `Duree`,
exclusion + warning of resolved tickets without a valid `Duree`) — `HH:MM` display
beyond 24 h, 4 h/8 h thresholds), warnings, Excel cell values/formats/colour
flags (zero vs `-`), PowerPoint structure incl. the red/green text highlight,
dynamic week renaming and the bundled samples.

## Performance

Roughly **150,000 rows × 10 columns convert in about a minute** (≈11 MB CSV →
≈8.5 MB XLSX, measured end-to-end over HTTP). Type inference and row writing are
both linear — no quadratic openpyxl calls (`max_row`/`max_column` are avoided
inside the row loop), and date parsing is gated by a cheap regex so only
date-looking columns reach `strptime`.

## Limits

- 300 MB per file, 50 files per request.
- Type inference samples the first 5,000 rows (conversion then runs over every row).
- Indian-style grouping (`1,23,400`) is treated as text; force `decimal_separator`
  or pre-clean such columns if needed.

## Project layout

```
converter.py           # converter engine: parsing, inference, Excel writing
hebdo.py               # Feature 1 engine: week logic, aggregations, MTTR, template fill, pptx
main.py                # FastAPI app (converter + hebdo endpoints)
static/index.html      # converter UI
static/hebdo.html      # weekly-report UI
templates/hebdo_S38.xlsx  # base report template (filled + renamed per target week)
test_converter.py      # converter end-to-end tests (142 checks)
test_hebdo.py          # weekly-report end-to-end tests (237 checks)
check_ui.py            # static UI checks for both pages
api_check_hebdo.py     # HTTP checks for the hebdo endpoints (server must be running)
sample-data/           # employees.csv + brise_smc_week39.csv + GNOC_week39.csv
requirements.txt
```
