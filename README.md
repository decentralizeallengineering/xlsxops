# xlsxops

Reliable Excel (`.xlsx`) operations for an LLM tool layer.

Each operation is a plain Python function that an agent calls as a tool. It takes JSON-style
arguments, validates them against a JSON Schema, does one job on a workbook, saves it safely and
returns a JSON-serializable report. It never raises for user errors:

```python
{"ok": True, "changed": True, "warnings": ["..."], "details": {...}}
{"ok": False, "changed": False, "warnings": [], "details": {"error": "column 'Qyt' not found; headers: [...]", "error_type": "column_not_found"}}
```

- Python 3.12+, Windows and Linux. Runtime dependencies: `openpyxl`, `jsonschema`. No network,
  telemetry, subprocesses or Office automation.
- `xlsxops.SCHEMAS[name]` holds each operation's argument schema (each ≤ 600 characters).
  `xlsxops.tool_definitions()` returns `{name, description, input_schema}` for every
  operation, and `xlsxops.call(name, args)` dispatches a tool call.
- Apache-2.0.

```bash
pip install .            # or: pip install -e ".[test]" && pytest
```

```python
import xlsxops as X

X.read_summary("sales.xlsx")
X.calculated_column("sales.xlsx", sheet="Sales", header="Total", formula="=[@Qty]*[@Price]")
X.total_row("sales.xlsx", sheet="Sales", columns=["Qty", "Total"])
X.chart("sales.xlsx", sheet="Sales", type="bar", x="Rep", y=["Total"], title="Revenue by rep")
X.verify_formulas("sales.xlsx")
```

## Operations

Common arguments: `path` (the workbook), `sheet` (default: the active sheet), `header_row`
(default: the first non-empty row), `out_path` (write a copy instead of overwriting `path`),
`force` (see *What is preserved*). Columns can be named by header text (case-insensitive) or
column letter.

| Operation | What it does |
|---|---|
| `read_summary` | Sheets, used ranges, headers, data row counts, 5 sample rows (formulas shown with their computed value), formula/chart/table/validation counts, defined names, content that a write would lose. |
| `write_cells` | `cells: {"B2": 5, "C2": "=B2*2"}` and/or `start` + 2-D `values`. `"=..."` is a formula (syntax-checked first). `"'=..."` is literal text. `"2024-05-01"` becomes a date. Reports overwritten values and the computed values of new formulas. |
| `add_sheet` | Adds a sheet at `index` with optional bold `headers`. Creates the workbook if `path` does not exist. |
| `add_column` | Appends a column after the table, from `values` (one per row) or a `fill` value/formula (relative refs shift per row). Copies the neighbouring header style and extends Excel Tables and the autofilter. |
| `calculated_column` | One formula per data row. Write it for the first data row (`=D2*E2`) or with `[@Header]` placeholders (`=[@Qty]*[@Price]`). Re-running on an existing header refills that column. |
| `set_formula` | One cell, or a range filled from its top-left formula with relative refs shifted. Returns the computed values. |
| `total_row` | SUM/AVERAGE/COUNT (also COUNTA/MIN/MAX) under the table, per column via `funcs`. Re-running replaces the existing total row. |
| `sort` | Multi-key sort, header stays put. Excel ordering: numbers < text < logicals, case-insensitive, blanks last. Formulas, styles, comments and hyperlinks move with their rows. A total row is left in place. |
| `filter_rows` | Copies the header plus rows matching `conditions` (`=, !=, >, >=, <, <=, contains, not_contains, starts_with, ends_with, in, not_in, blank, not_blank, regex`; `match: all/any`) to a new sheet, as values with styles. |
| `dedupe` | Removes duplicate rows by key columns (case-insensitive and trimmed by default, like Excel's Remove Duplicates). Keeps the first or last row. Reports every removed row and the row it duplicated. References that ended at the old last row are shrunk: formulas, names, chart series, validations, conditional formats, tables. |
| `split_column` | By `delimiter`, or `regex` (capture groups → one column each; otherwise split). New columns go after the table. Numbers are inferred, leading zeros stay text. |
| `lookup` | Exact-match join from another sheet. `mode: formula` (`style: xlookup` / `vlookup` / `index_match`; VLOOKUP falls back to INDEX/MATCH when the return column is left of the key) or `mode: value`. Reports matched/unmatched keys, duplicate source keys, number-vs-text key mismatches, and matched-but-blank cells (Excel shows 0 for those). |
| `pivot_summary` | Group by one or more columns with sum/count/avg/min/max into a new sheet as values, with a Total row. `formulas: true` adds a live SUMIFS/COUNTIFS/AVERAGEIFS version on `"<new_sheet> (formulas)"` and checks that it matches the values version. |
| `chart` | bar/line/pie/scatter from `x` + `y` columns or a `range` (first row = series names, first column = categories). Title and axis titles. Placed right of the data and below existing charts/images without overlapping. |
| `conditional_format` | `color_scale`, `data_bar`, `cell_is` (`operator` + `values`), or `outliers`: a live z-score rule `ABS(x-AVERAGE)>z*STDEV`, with the flagged cells reported. |
| `data_validation` | `list` (inline `values` or a `source` range), `whole`, `decimal`, `date` with `min`/`max`. Reports existing cells that break the new rule and overlaps with existing validations. |
| `format_range` | `number_format` (presets `integer, decimal, currency, percent, date, datetime, text`, or any Excel format code), `bold_header`, `freeze` (`header`, a cell, or `none`), `autofit` (width estimate from displayed text). Other style attributes are untouched. |
| `from_csv` | CSV → new sheet (new workbook if `path` is missing). Sniffs the delimiter and decodes UTF-8 with cp1252/latin-1 fallbacks. Per-column type inference: numbers (thousands separators, currency, %, accounting negatives), ISO/US/EU dates (`dayfirst` resolves ambiguity; ambiguous columns are flagged), booleans. Values with leading zeros or more than 15 digits stay text. A field starting with `=` is stored as text, never as a live formula. |
| `multisheet_report` | One sheet per distinct value of a column (rows as values, header style, widths, freeze, autofilter), plus an index sheet with hyperlinks and counts. |
| `verify_formulas` | Evaluates every formula. Reports errors (`#REF!`, `#VALUE!` type errors, `#DIV/0!`, `#N/A`, `#NAME?`, and whether each error originates in that cell or is inherited), circular references and cells that depend on them, and formulas it cannot evaluate, with the reason. `write_cache: true` stores the computed values. |

## Formulas and cached values

openpyxl writes formulas without results, so a file it saves shows empty cells in any reader that
does not recalculate: pandas, openpyxl `data_only`, previewers, and many converters. xlsxops ships
a small pure-Python evaluator (`xlsxops.formula`). After every write it:

1. evaluates every formula in the workbook (`verify_formulas` runs this same pass);
2. writes the results into the file as cached values, with the right types (number, string,
   boolean, error);
3. sets `fullCalcOnLoad` so Excel still recalculates on open;
4. adds a `formula_check` summary to the report, and warnings for any errors, circular references
   or unevaluated formulas.

If a formula is not supported and its text is unchanged since the file was loaded, the cached value
the file already had is kept.

Supported: arithmetic (`+ - * / ^ %`, Excel precedence including `-2^2 = 4`), comparison, `&`,
cell/range/whole-column references across sheets, defined names, array constants, and
`SUM AVERAGE COUNT COUNTA COUNTBLANK COUNTIF(S) SUMIF(S) AVERAGEIF(S) MINIFS MAXIFS MIN MAX IF
IFERROR IFNA AND OR NOT ROUND ROUNDUP ROUNDDOWN ABS INT MOD SQRT SUMPRODUCT SUBTOTAL VLOOKUP HLOOKUP
XLOOKUP INDEX MATCH TEXT CONCAT(ENATE) LEN LEFT RIGHT MID UPPER LOWER TRIM ISBLANK ISNUMBER ISTEXT
ISERROR ISNA DATE YEAR MONTH DAY`. Criteria support operators, wildcards (`* ? ~`) and Excel's
case-insensitive text matching. Anything else (other functions, structured table references,
external links, spilling array results) is listed as "not evaluated"; it is never guessed.

Functions newer than Excel 2007 (XLOOKUP, IFNA, MAXIFS, ...) are stored with the `_xlfn.` prefix
Excel requires. Without it, Excel shows `#NAME?`.

## What is preserved

Every write keeps existing cell styles, number formats, rich text, column widths, row heights,
other sheets, defined names, charts, images, Excel tables, data validations, conditional formats,
comments, hyperlinks, merged cells, freeze panes and VBA (`.xlsm`). The test suite checks this for
every operation. Images survive even without Pillow: xlsxops carries their bytes through
unchanged.

Some content openpyxl cannot carry through a load/save: shapes and text boxes, slicers,
timelines, form/ActiveX controls, sparklines, Excel 2010+ conditional-format/validation
extensions, new chart types (waterfall, treemap…), Power Pivot models, query connections,
threaded comments. If a file contains any of these, write operations refuse with
`error_type: "would_lose_content"` and list what would go. `force: true` proceeds anyway; use
`out_path` to keep the original. `read_summary` lists such content up front. Chart style/color
side-parts and custom XML are dropped with a warning only.

Saves are atomic: the file is written to a temp file next to the target and swapped in. A failure,
or a target file locked by Excel, leaves the original untouched and is reported as an error.

Deliberate limits:
- Columns are added after the table, never inserted inside it. Inserting would require rewriting
  every reference in the workbook.
- `filter_rows`, `multisheet_report` and the values version of `pivot_summary` copy formula cells
  as their computed values.
- `sort` behaves like Excel: relative references move with their row. Formulas that point at other
  rows (running totals) are flagged in warnings.

## Tests

```bash
pip install -e ".[test]"
pytest
```

Tests build their own fixtures. Coverage: one test per operation, a round-trip test per write
operation proving styles/charts/sheets/names survive, cached values present after writes, images
kept without Pillow, the lossy-content guard, evaluator cases for every supported function, and
`verify_formulas` on a workbook with broken references, type errors, a circular chain and
unsupported functions.
