# Excel Criteria Data Match

A small CLI tool for a common bookkeeping/reporting problem: you have a
**detailed working sheet** (e.g. a system-exported trial balance or P&L)
and a **fixed reporting layout** (e.g. a reference balance sheet / P&L
format), and you need the right amount from the detail sheet lined up
against the right label in the reporting layout — every month, without
manually re-matching rows by eye.

This tool does that matching for you and writes a ready-to-paste result.

## What it does

1. Scans the current folder for Excel files (`.xlsx`, `.xlsm`, `.xls`).
2. You pick a **REFERENCE** source (file → sheet → the column holding
   the row labels you want to keep, e.g. `TRADE DEBTORS`, `TOTAL FIXED
   ASSETS`) and a **DATA** source (file → sheet → the column with the
   same labels + the column with the amounts).
3. It matches every reference row to a data row by label text, and
   writes a new workbook that mirrors the reference sheet's exact row
   order and spacing (shifted down by one row for a header), with the
   matched amount(s) ready to copy straight into your real workbook.
   You can pull more than one data column in a single run — a range
   (`F:H`), a comma list (`F,G,B`), or a mix (`F:H,B`) — instead of
   re-running the tool once per column.

## Why not just VLOOKUP?

Two reasons this isn't a one-line formula:

- **Repeated labels.** Financial statements often reuse the same
  wording twice — a section header and a later total line (e.g.
  `LONG TERM LIABILITIES` appearing once as a header and again as the
  total further down). A plain lookup grabs the first match for both.
  This tool matches by **occurrence**: the 1st time a label appears in
  the reference is matched to the 1st time it appears in the data, the
  2nd to the 2nd, and so on, top to bottom — so a header and a total
  sharing identical text never collide.
- **Silent data loss.** If the data sheet has *more* occurrences of a
  label than the reference has rows for, a lookup just never surfaces
  those extra amounts — real money quietly disappears from the output.
  This tool treats that as an error: unmatched data rows are flagged
  on screen and written to a `Data Overflow` sheet in the output file
  (criteria text, amount, source cell references, and source file), so
  nothing gets lost without you knowing.

## Usage

```
pip install openpyxl xlrd
python excel_criteria_data_match.py
```

Or run the pre-built Windows `.exe` (no Python required) — see
[Releases](../../releases) if one has been published.

Full step-by-step walkthrough, output format, and troubleshooting:
[`excel_criteria_data_match Instruction.txt`](excel_criteria_data_match%20Instruction.txt).

## Requirements

- Python 3.9+ (only if running from source)
- `openpyxl` (reads `.xlsx` / `.xlsm`), `xlrd` (reads legacy `.xls`)

## License

MIT — see [LICENSE](LICENSE). Free to use, modify, and share.
