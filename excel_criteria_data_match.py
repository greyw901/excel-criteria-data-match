"""
Match a criteria column from one Excel sheet (DATA) onto the row layout of
another Excel sheet (REFERENCE), by matching repeated label text in order of
occurrence, and write the result to a new Excel file ready to copy-paste.

Author: Yeoh Huey Sun
Email:  greyw901@gmail.com
GitHub: https://github.com/greyw901
"""

import os
import re
import sys
import time
import traceback
import warnings
from collections import defaultdict
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import openpyxl
import xlrd
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string

warnings.filterwarnings("ignore", message="Cannot parse header or footer")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, "frozen", False) else __file__
))
os.chdir(SCRIPT_DIR)

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")
CURRENCY_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
LOG_PREFIX = "error log"


class UserCancelled(Exception):
    pass


# ---------------------------------------------------------------------------
# File / workbook access
# ---------------------------------------------------------------------------

def list_excel_files() -> List[str]:
    names = []
    for name in os.listdir(SCRIPT_DIR):
        if name.startswith("~$"):
            continue
        if os.path.splitext(name)[1].lower() in EXCEL_EXTENSIONS:
            names.append(name)
    return sorted(names)


def get_sheet_names(filename: str) -> List[str]:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".xls":
        wb = xlrd.open_workbook(filename, on_demand=True)
        names = list(wb.sheet_names())
        wb.release_resources()
        return names
    wb = openpyxl.load_workbook(filename, read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def read_columns(filename: str, sheet: str, col_letters: List[str]) -> Dict[int, List]:
    """Return {row_number (1-based): [value_per_col_letter]} for the whole sheet."""
    col_idxs = [column_index_from_string(c.upper()) for c in col_letters]
    ext = os.path.splitext(filename)[1].lower()
    rows: Dict[int, List] = {}

    if ext == ".xls":
        wb = xlrd.open_workbook(filename)
        ws = wb.sheet_by_name(sheet)
        for r in range(ws.nrows):
            vals = []
            for ci in col_idxs:
                v = ws.cell_value(r, ci - 1) if ci - 1 < ws.ncols else None
                vals.append(None if v == "" else v)
            rows[r + 1] = vals
    else:
        wb = openpyxl.load_workbook(filename, data_only=True)
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            vals = [ws.cell(row=r, column=ci).value for ci in col_idxs]
            rows[r] = vals
        wb.close()

    return rows


def last_nonempty_row(rows: Dict[int, List], col_positions: List[int]) -> int:
    last = 0
    for r, vals in rows.items():
        for pos in col_positions:
            v = vals[pos]
            if v is not None and str(v).strip() != "":
                last = max(last, r)
                break
    return last


# ---------------------------------------------------------------------------
# Console prompts
# ---------------------------------------------------------------------------

def _print_heading(text: str) -> None:
    print()
    print(text)
    print()


def _confirm_exit_or_back() -> Optional[bool]:
    """Blank [Enter] at a selection prompt asks for confirmation before exiting.

    Returns True to exit (UserCancelled raised by caller), False to go back
    to the file list, or None to re-show the original prompt.
    """
    print()
    print("[Enter] again to exit")
    print("[0] to go back")
    choice = input("> ").strip()
    if choice == "":
        return True
    if choice == "0":
        return False
    return None


def prompt_select_file(files: List[str], label: str) -> str:
    while True:
        _print_heading(f"Select {label} source:")
        for i, name in enumerate(files, start=1):
            print(f'   [{i}] "{name}"')
        print()
        print("[#] to select")
        print("[0] to exit")
        choice = input("> ").strip()
        if choice == "0":
            raise UserCancelled()
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice) - 1]
        print()
        print(f"Invalid Input: {choice!r} is not a valid selection")


def prompt_select_sheet_cols(
    sheet_names: List[str], label: str, num_cols: int
) -> Optional[Tuple[int, List[str]]]:
    if num_cols == 1:
        pattern = re.compile(r"^(\d+)_([A-Za-z]+)$")
        example = "4_F"
    else:
        pattern = re.compile(r"^(\d+)_([A-Za-z]+)_([A-Za-z]+)$")
        example = "1_D_I"

    while True:
        if num_cols == 1:
            _print_heading(f"Select {label} [Criteria Sheet]_[Criteria Col]")
        else:
            _print_heading(f"Select {label} [Criteria Sheet]_[Criteria Col]_[Data Col]")
        print(f"Eg.: {example} + [Enter]")
        print()
        for i, name in enumerate(sheet_names, start=1):
            print(f'   [{i}] "{name}"')
        print()
        print("[#]_[Col] to select")
        print("[0] to back")
        print("[Enter] to exit")
        raw = input("> ").strip()

        if raw == "":
            result = _confirm_exit_or_back()
            if result is True:
                raise UserCancelled()
            if result is False:
                return None
            continue

        if raw == "0":
            return None

        m = pattern.match(raw)
        if not m:
            print()
            print(f"Invalid Input: {raw!r} does not match the pattern {example}")
            continue

        sheet_num = int(m.group(1))
        cols = [m.group(i) for i in range(2, 2 + num_cols)]
        if not (1 <= sheet_num <= len(sheet_names)):
            print()
            print(f"Invalid Input: sheet number {sheet_num} is out of range")
            continue
        try:
            for c in cols:
                column_index_from_string(c.upper())
        except ValueError:
            print()
            print(f"Invalid Input: {raw!r} has an invalid column letter")
            continue

        return sheet_num, [c.upper() for c in cols]


def prompt_confirm_source(
    label: str, filename: str, sheet_name: str, cols: List[str]
) -> bool:
    while True:
        _print_heading(f"Confirm {label} source:")
        print(f'   "{filename}"')
        if len(cols) == 1:
            print(f'   Sheet: [{sheet_name}], Col: [{cols[0]}]')
        else:
            print(f'   Sheet: [{sheet_name}], Criteria Col: [{cols[0]}], Data Col: [{cols[1]}]')
        print()
        print("[Enter] to Confirm")
        print("[0] to reselect")
        choice = input("> ").strip()
        if choice == "":
            return True
        if choice == "0":
            return False
        print()
        print(f"Invalid Input: {choice!r}")


def select_source(files: List[str], label: str, num_cols: int) -> Tuple[str, str, List[str]]:
    while True:
        filename = prompt_select_file(files, label)
        sheet_names = get_sheet_names(filename)
        while True:
            picked = prompt_select_sheet_cols(sheet_names, label, num_cols)
            if picked is None:
                break
            sheet_num, cols = picked
            sheet_name = sheet_names[sheet_num - 1]
            if prompt_confirm_source(label, filename, sheet_name, cols):
                return filename, sheet_name, cols


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

def build_data_lookup(
    data_rows: Dict[int, List]
) -> Tuple[Dict[Tuple[str, int], object], Dict[Tuple[str, int], int]]:
    lookup: Dict[Tuple[str, int], object] = {}
    key_row: Dict[Tuple[str, int], int] = {}
    counts: Dict[str, int] = defaultdict(int)
    for r in sorted(data_rows):
        text_val, data_val = data_rows[r]
        if isinstance(text_val, str) and text_val.strip():
            key_text = text_val.strip()
            counts[key_text] += 1
            key = (key_text, counts[key_text])
            lookup[key] = data_val
            key_row[key] = r
    return lookup, key_row


def build_output_rows(
    ref_rows: Dict[int, List], lookup: Dict[Tuple[str, int], object], last_row: int
) -> Tuple[List[Tuple[int, object, object]], List[Tuple[int, str]], set]:
    counts: Dict[str, int] = defaultdict(int)
    output: List[Tuple[int, object, object]] = []
    unmatched: List[Tuple[int, str]] = []
    used_keys: set = set()

    for r in range(1, last_row + 1):
        raw = ref_rows.get(r, [None])[0]
        if isinstance(raw, str) and raw.strip():
            key_text = raw.strip()
            counts[key_text] += 1
            key = (key_text, counts[key_text])
            if key in lookup:
                output.append((r, raw, lookup[key]))
                used_keys.add(key)
            else:
                output.append((r, raw, None))
                unmatched.append((r, key_text))
        else:
            output.append((r, raw, None))

    return output, unmatched, used_keys


def find_overflow(
    lookup: Dict[Tuple[str, int], object],
    key_row: Dict[Tuple[str, int], int],
    used_keys: set,
) -> List[Tuple[int, str, object]]:
    """DATA rows whose criteria text/occurrence has no matching REFERENCE row.

    This means the DATA source has more occurrences of that criteria than the
    REFERENCE layout does (or REFERENCE has none at all) -- those amounts
    would silently be left out of the output if not flagged.
    """
    overflow = []
    for key in lookup:
        if key not in used_keys:
            overflow.append((key_row[key], key[0], lookup[key]))
    overflow.sort(key=lambda item: item[0])
    return overflow


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def unique_path(base_name: str, ext: str) -> str:
    candidate = os.path.join(SCRIPT_DIR, base_name + ext)
    if not os.path.exists(candidate):
        return candidate
    n = 1
    while True:
        candidate = os.path.join(SCRIPT_DIR, f"{base_name}_{n}{ext}")
        if not os.path.exists(candidate):
            return candidate
        n += 1


def build_output_filename(data_file: str, ref_sheet: str) -> str:
    m = re.search(r"(20\d{2})(0[1-9]|1[0-2])", data_file)
    if m:
        yymm = m.group(1)[2:] + m.group(2)
        base = f"{ref_sheet} matched {yymm}"
    else:
        base = f"Match Result_{ref_sheet}_{datetime.now():%Y%m%d}"
    return unique_path(base, ".xlsx")


def write_output(
    path: str,
    output_rows: List[Tuple[int, object, object]],
    overflow: List[Tuple[int, str, object]],
    data_file: str,
    data_crit_col: str,
    data_data_col: str,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Matched"

    max_len = 10
    for _, raw, _ in output_rows:
        if raw is not None:
            max_len = max(max_len, len(str(raw)))
    ws.column_dimensions["A"].width = min(60, max_len + 2)
    ws.column_dimensions["B"].width = 16

    for row, raw, val in output_rows:
        if raw is not None:
            ws.cell(row=row, column=1, value=raw)
        if val is not None:
            cell = ws.cell(row=row, column=2, value=val)
            if isinstance(val, (int, float)):
                cell.number_format = CURRENCY_FORMAT

    if overflow:
        write_overflow_sheet(wb, overflow, data_file, data_crit_col, data_data_col)

    wb.save(path)


def write_overflow_sheet(
    wb: Workbook,
    overflow: List[Tuple[int, str, object]],
    data_file: str,
    data_crit_col: str,
    data_data_col: str,
) -> None:
    ws = wb.create_sheet("Data Overflow")
    headers = ["Data Criteria", "Data", "Criteria Cell", "Data Cell", "Source"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = openpyxl.styles.Font(bold=True)

    source_name = os.path.splitext(os.path.basename(data_file))[0]
    for i, (row, text, val) in enumerate(overflow, start=2):
        ws.cell(row=i, column=1, value=text)
        data_cell = ws.cell(row=i, column=2, value=val)
        if isinstance(val, (int, float)):
            data_cell.number_format = CURRENCY_FORMAT
        ws.cell(row=i, column=3, value=f"{data_crit_col}{row}")
        ws.cell(row=i, column=4, value=f"{data_data_col}{row}")
        ws.cell(row=i, column=5, value=source_name)

    max_len = max((len(str(text)) for _, text, _ in overflow), default=10)
    ws.column_dimensions["A"].width = min(60, max_len + 2)
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = max(14, len(source_name) + 2)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run() -> None:
    print("=" * 56)
    print(" Excel Criteria Matcher")
    print("=" * 56)

    files = list_excel_files()
    if not files:
        print()
        print("ERROR: no Excel files found in this folder")
        return

    ref_file, ref_sheet, ref_cols = select_source(files, "REFERENCE", 1)
    data_file, data_sheet, data_cols = select_source(files, "DATA", 2)

    start = time.time()
    print()
    print("-" * 56)
    print(f"Input & Output folder: \"{SCRIPT_DIR}\"")
    print()
    print("Inputs: 2")
    print(f"   1. \"{ref_file}\"!{ref_sheet}  Col [{ref_cols[0]}]")
    print(f"   2. \"{data_file}\"!{data_sheet}  Criteria Col [{data_cols[0]}], Data Col [{data_cols[1]}]")

    print()
    print("Processing")
    print("Reading reference ...")
    ref_rows = read_columns(ref_file, ref_sheet, ref_cols)
    ref_last = last_nonempty_row(ref_rows, [0])
    print(f"Reference rows : {ref_last}")

    print("Reading data ...")
    data_rows = read_columns(data_file, data_sheet, data_cols)
    data_last = last_nonempty_row(data_rows, [0, 1])
    print(f"Data rows : {data_last}")

    print("Matching ...")
    lookup, key_row = build_data_lookup(data_rows)
    output_rows, unmatched, used_keys = build_output_rows(ref_rows, lookup, ref_last)
    overflow = find_overflow(lookup, key_row, used_keys)
    matched_count = sum(1 for _, raw, val in output_rows if val is not None)
    print(f"Matched : {matched_count}")

    print("Writing output ...")
    out_path = build_output_filename(data_file, ref_sheet)
    write_output(out_path, output_rows, overflow, data_file, data_cols[0], data_cols[1])

    elapsed = time.time() - start
    print()
    print(f"Outputs: 1")
    print(f"   1. \"{os.path.basename(out_path)}\"")

    if unmatched:
        print()
        print(f"Unmatched reference labels: {len(unmatched)}")
        for i, (row, text) in enumerate(unmatched, start=1):
            print(f"   {i}. \"{text}\" (row {row})")

    if overflow:
        print()
        print(f"ERROR - Data overflow: {len(overflow)}")
        print("   DATA has more occurrences of this criteria than REFERENCE has rows for it.")
        print("   These amounts are NOT in the output and would be missed:")
        for i, (row, text, val) in enumerate(overflow, start=1):
            print(f"   {i}. \"{text}\" (data row {row}) = {val!r}")

    print()
    print(f"Elapsed: {elapsed:.1f}s")
    print("-" * 56)

    if unmatched or overflow:
        print("[Enter] to close...")
        print("[0] save error log and close")
        choice = input("> ").strip()
        if choice == "0":
            save_report_log(unmatched, overflow, ref_file, ref_sheet, data_file, data_sheet)
    else:
        input("[Enter] to close...")


def save_report_log(
    unmatched: List[Tuple[int, str]],
    overflow: List[Tuple[int, str, object]],
    ref_file: str,
    ref_sheet: str,
    data_file: str,
    data_sheet: str,
) -> None:
    path = unique_path(f"{LOG_PREFIX}_{datetime.now():%Y%m%d}", ".txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"Reference: \"{ref_file}\"!{ref_sheet}\n")
        f.write(f"Data: \"{data_file}\"!{data_sheet}\n\n")
        f.write(f"Unmatched reference labels: {len(unmatched)}\n")
        for row, text in unmatched:
            f.write(f"  row {row}: \"{text}\"\n")
        f.write(f"\nData overflow (not carried into output): {len(overflow)}\n")
        for row, text, val in overflow:
            f.write(f"  data row {row}: \"{text}\" = {val!r}\n")
    print(f"Saved: \"{os.path.basename(path)}\"")


def main() -> None:
    try:
        run()
    except UserCancelled:
        print()
        print("Cancelled.")
        input("[Enter] to close...")
    except Exception as e:
        print()
        print(f"ERROR: {type(e).__name__} - {e}")
        print(traceback.format_exc())
        input("[Enter] to close...")


if __name__ == "__main__":
    main()
