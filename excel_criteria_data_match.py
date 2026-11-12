"""
Match a criteria column from one Excel sheet (DATA) onto the row layout of
another Excel sheet (REFERENCE), by matching repeated label text in order of
occurrence, and write the result to a new Excel file ready to copy-paste.
DATA may contribute more than one value column per matched row (a range like
F:H or a list like F,G,B), in which case the output carries one column per
selected data column, in the order selected.

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
from typing import Dict, List, Optional, Tuple

import openpyxl
import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import column_index_from_string, get_column_letter

warnings.filterwarnings("ignore", message="Cannot parse header or footer")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, "frozen", False) else __file__
))
os.chdir(SCRIPT_DIR)

TOOL_NAME = "Excel Criteria Matcher"
VERSION = "v1.2_20261113"  # not shown on screen; goes into the exe version file
EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")
CURRENCY_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
DIVIDER_WIDTH = 56
MATCHED_SHEET = "Matched"
UNMATCHED_SHEET = "Unmatched Reference"
OVERFLOW_SHEET = "Data Overflow"
WIDTH_MARGIN = 2
MAX_TEXT_WIDTH = 60
FIRST_CANCEL_STEP = 2  # [00] shows from the 3rd prompt screen on


class GoBack(Exception):
    """[0] Back - return to the previous screen."""


class UserExit(Exception):
    """[0] Exit at the first screen - close immediately, no screen."""


class UserCancelled(Exception):
    """[00] Cancel & exit - show Cancelled. and pause before closing."""


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
# Column selector parsing (DATA source accepts multiple data columns)
# ---------------------------------------------------------------------------

def parse_data_columns(spec: str) -> List[str]:
    """Expand a Data Col(s) spec into an ordered list of column letters.

    Accepts a single letter ("F"), a range ("F:H" -> F, G, H, ascending
    only), a comma list ("F,G,B" -> exactly that order), or a mix
    ("F:H,B"). Raises ValueError with a user-facing message on anything
    invalid.
    """
    cols: List[str] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            raise ValueError(f"empty column in {spec!r}")
        if ":" in part:
            start, _, end = part.partition(":")
            try:
                start_idx = column_index_from_string(start.upper())
                end_idx = column_index_from_string(end.upper())
            except ValueError:
                raise ValueError(f"{part!r} is not a valid column range")
            if end_idx < start_idx:
                raise ValueError(
                    f"range {part!r} must be ascending (e.g. F:H, not H:F) "
                    f"- use a comma list for custom order"
                )
            for idx in range(start_idx, end_idx + 1):
                cols.append(get_column_letter(idx))
        else:
            try:
                column_index_from_string(part.upper())
            except ValueError:
                raise ValueError(f"{part!r} is not a valid column letter")
            cols.append(part.upper())
    return cols


# ---------------------------------------------------------------------------
# Console screens
# ---------------------------------------------------------------------------

def clear_screen() -> None:
    os.system("cls")


def print_banner() -> None:
    print("=" * DIVIDER_WIDTH)
    print(TOOL_NAME)
    print("=" * DIVIDER_WIDTH)
    print()


def open_screen(choices: List[str], error: Optional[str]) -> None:
    """Clears the console and prints the screen's leading lines: banner, the
    choices made so far, and the invalid-input notice from the previous try
    (error None = no notice, "" = bare notice, text = notice with reason)."""
    clear_screen()
    print_banner()
    if choices:
        for line in choices:
            print(line)
        print()
    if error is not None:
        print(f"Error: Invalid Input - {error}" if error else "Error: Invalid Input")
        print()


def print_options(options: List[str]) -> None:
    print()
    print("-" * DIVIDER_WIDTH)
    for line in options:
        print(line)
    print()


def back_options(first_screen: bool, allow_cancel: bool) -> List[str]:
    opts = ["[0] Exit" if first_screen else "[0] Back"]
    if allow_cancel:
        opts.append("[00] Cancel & exit")
    return opts


def read_choice(prompt: str, first_screen: bool, allow_cancel: bool) -> str:
    """Reads one input line and handles the navigation keys: [00] is matched
    on the whole input before any selection parsing; [0] backs out."""
    raw = input(prompt).strip()
    if allow_cancel and raw == "00":
        raise UserCancelled()
    if raw == "0":
        raise UserExit() if first_screen else GoBack()
    return raw


def value_lines(pairs: List[Tuple[str, str]]) -> List[str]:
    width = max(len(label) for label, _ in pairs)
    return [f"{label.ljust(width)} : {value}" for label, value in pairs]


def describe_reference(filename: str, sheet: Optional[str] = None,
                       cols: Optional[List[str]] = None) -> str:
    text = f'"{filename}"'
    if sheet is not None:
        text += f"!{sheet}, col {cols[0]}"
    return text


def describe_data(filename: str, sheet: Optional[str] = None,
                  cols: Optional[List[str]] = None) -> str:
    text = f'"{filename}"'
    if sheet is not None:
        text += f"!{sheet}, criteria col {cols[0]}, data cols {', '.join(cols[1:])}"
    return text


def prompt_select_file(files: List[str], label: str, choices: List[str],
                       first_screen: bool, allow_cancel: bool) -> str:
    error: Optional[str] = None
    while True:
        open_screen(choices, error)
        print(f"Select {label} source:")
        print()
        for i, name in enumerate(files, start=1):
            print(f'   [{i}] "{name}"')
        print_options(["[#] Select"] + back_options(first_screen, allow_cancel))
        raw = read_choice(f"{label.title()}: ", first_screen, allow_cancel)
        if raw.isdigit() and 1 <= int(raw) <= len(files):
            return files[int(raw) - 1]
        error = ""


def prompt_select_sheet_col(sheet_names: List[str], choices: List[str],
                            allow_cancel: bool) -> Tuple[str, List[str]]:
    """REFERENCE: sheet + exactly one label column."""
    pattern = re.compile(r"^(\d+)_([A-Za-z]+)$")
    example = "4_F"
    error: Optional[str] = None

    while True:
        open_screen(choices, error)
        print("Select REFERENCE sheet & col:")
        print()
        print(f"e.g. {example}")
        print()
        for i, name in enumerate(sheet_names, start=1):
            print(f'   [{i}] "{name}"')
        print_options(["[#]_[Col] Select"] + back_options(False, allow_cancel))
        raw = read_choice("Sheet & col: ", False, allow_cancel)

        if raw == "":
            error = ""
            continue
        m = pattern.match(raw)
        if not m:
            error = f"{raw!r} does not match the pattern {example}"
            continue
        sheet_num = int(m.group(1))
        col = m.group(2).upper()
        if not (1 <= sheet_num <= len(sheet_names)):
            error = f"sheet number {sheet_num} is out of range"
            continue
        try:
            column_index_from_string(col)
        except ValueError:
            error = f"{raw!r} has an invalid column letter"
            continue

        return sheet_names[sheet_num - 1], [col]


def prompt_select_sheet_cols_data(sheet_names: List[str], choices: List[str],
                                  allow_cancel: bool) -> Tuple[str, List[str]]:
    """DATA: sheet + one criteria column + one or more data columns."""
    pattern = re.compile(r"^(\d+)_([A-Za-z]+)_([A-Za-z:,]+)$")
    error: Optional[str] = None

    while True:
        open_screen(choices, error)
        print("Select DATA sheet & cols:")
        print()
        print("e.g. 1_D_I        one data col")
        print("     1_D_F:H      range: F, G, H")
        print("     1_D_F,G,B    list: F, G, B")
        print()
        for i, name in enumerate(sheet_names, start=1):
            print(f'   [{i}] "{name}"')
        print_options(["[#]_[Col]_[Col(s)] Select"] + back_options(False, allow_cancel))
        raw = read_choice("Sheet & cols: ", False, allow_cancel)

        if raw == "":
            error = ""
            continue
        m = pattern.match(raw)
        if not m:
            error = f"{raw!r} does not match the pattern 1_D_F:H"
            continue
        sheet_num = int(m.group(1))
        crit_col = m.group(2).upper()
        if not (1 <= sheet_num <= len(sheet_names)):
            error = f"sheet number {sheet_num} is out of range"
            continue
        try:
            column_index_from_string(crit_col)
        except ValueError:
            error = f"{crit_col!r} is not a valid column letter"
            continue
        try:
            data_cols = parse_data_columns(m.group(3))
        except ValueError as e:
            error = str(e)
            continue
        if crit_col in data_cols:
            error = f"criteria col {crit_col} cannot also be a data col"
            continue
        if len(set(data_cols)) != len(data_cols):
            dupes = sorted({c for c in data_cols if data_cols.count(c) > 1})
            error = f"duplicate data col(s) {', '.join(dupes)}"
            continue

        return sheet_names[sheet_num - 1], [crit_col] + data_cols


def prompt_confirm_source(label: str, details: List[str], choices: List[str]) -> None:
    """Returns on [Enter]; raises GoBack on [0], UserCancelled on [00]."""
    error: Optional[str] = None
    while True:
        open_screen(choices, error)
        print(f"Confirm {label} source:")
        print()
        for line in details:
            print(f"   {line}")
        print_options(["[Enter] Continue", "[0] Back", "[00] Cancel & exit"])
        raw = read_choice("Ans: ", False, True)
        if raw == "":
            return
        error = ""


def select_sources(files: List[str]) -> Tuple[str, str, List[str], str, str, List[str]]:
    """Runs the six prompt screens as a step machine. GoBack is caught here,
    and only here: it moves one step back. A confirm screen's choices block
    leaves out the source being confirmed - its details are the screen body."""
    s: Dict[str, object] = {}
    step = 0
    while step < 6:
        allow_cancel = step >= FIRST_CANCEL_STEP
        try:
            if step == 0:
                s["ref_file"] = prompt_select_file(files, "REFERENCE", [], True, allow_cancel)
                s["ref_sheets"] = get_sheet_names(s["ref_file"])
            elif step == 1:
                choices = [f"Reference : {describe_reference(s['ref_file'])}"]
                s["ref_sheet"], s["ref_cols"] = prompt_select_sheet_col(
                    s["ref_sheets"], choices, allow_cancel)
            elif step == 2:
                details = value_lines([
                    ("File", f'"{s["ref_file"]}"'),
                    ("Sheet", f'"{s["ref_sheet"]}"'),
                    ("Col", s["ref_cols"][0]),
                ])
                prompt_confirm_source("REFERENCE", details, [])
            elif step == 3:
                choices = [f"Reference : {describe_reference(s['ref_file'], s['ref_sheet'], s['ref_cols'])}"]
                s["data_file"] = prompt_select_file(files, "DATA", choices, False, allow_cancel)
                s["data_sheets"] = get_sheet_names(s["data_file"])
            elif step == 4:
                choices = value_lines([
                    ("Reference", describe_reference(s["ref_file"], s["ref_sheet"], s["ref_cols"])),
                    ("Data", describe_data(s["data_file"])),
                ])
                s["data_sheet"], s["data_cols"] = prompt_select_sheet_cols_data(
                    s["data_sheets"], choices, allow_cancel)
            elif step == 5:
                choices = [f"Reference : {describe_reference(s['ref_file'], s['ref_sheet'], s['ref_cols'])}"]
                details = value_lines([
                    ("File", f'"{s["data_file"]}"'),
                    ("Sheet", f'"{s["data_sheet"]}"'),
                    ("Criteria col", s["data_cols"][0]),
                    ("Data cols", ", ".join(s["data_cols"][1:])),
                ])
                prompt_confirm_source("DATA", details, choices)
            step += 1
        except GoBack:
            step -= 1
    return (s["ref_file"], s["ref_sheet"], s["ref_cols"],
            s["data_file"], s["data_sheet"], s["data_cols"])


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------

def build_data_lookup(
    data_rows: Dict[int, List], n_data_cols: int
) -> Tuple[Dict[Tuple[str, int], Tuple], Dict[Tuple[str, int], int]]:
    lookup: Dict[Tuple[str, int], Tuple] = {}
    key_row: Dict[Tuple[str, int], int] = {}
    counts: Dict[str, int] = defaultdict(int)
    for r in sorted(data_rows):
        row_vals = data_rows[r]
        text_val = row_vals[0]
        data_vals = tuple(row_vals[1:1 + n_data_cols])
        if isinstance(text_val, str) and text_val.strip():
            key_text = text_val.strip()
            counts[key_text] += 1
            key = (key_text, counts[key_text])
            lookup[key] = data_vals
            key_row[key] = r
    return lookup, key_row


def build_output_rows(
    ref_rows: Dict[int, List], lookup: Dict[Tuple[str, int], Tuple], last_row: int
) -> Tuple[List[Tuple[int, object, Optional[Tuple]]], List[Tuple[int, str]], set]:
    counts: Dict[str, int] = defaultdict(int)
    output: List[Tuple[int, object, Optional[Tuple]]] = []
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
    lookup: Dict[Tuple[str, int], Tuple],
    key_row: Dict[Tuple[str, int], int],
    used_keys: set,
) -> List[Tuple[int, str, Tuple]]:
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
        base = f"Match Result_{ref_sheet}_{time.strftime('%Y%m%d')}"
    return unique_path(base, ".xlsx")


def display_len(value: object) -> int:
    """Rendered length of a cell: numbers as their accounting-format string
    (thousands separators, 2 decimals, padding/parentheses either side)."""
    if value is None:
        return 0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return len(f"{abs(value):,.2f}") + 2
    return len(str(value))


def fit_width(values: List[object], cap: Optional[int] = None) -> int:
    width = max((display_len(v) for v in values), default=0) + WIDTH_MARGIN
    return min(width, cap) if cap else width


def write_header_row(ws, headers: List[str]) -> None:
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")


def write_value(ws, row: int, col: int, value: object) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.number_format = CURRENCY_FORMAT


def set_group_width(ws, first_col: int, columns: List[List[object]]) -> None:
    """Same-quantity columns share one width: the max across all of them."""
    width = max((fit_width(c) for c in columns), default=WIDTH_MARGIN)
    for j in range(len(columns)):
        ws.column_dimensions[get_column_letter(first_col + j)].width = width


def write_output(
    path: str,
    output_rows: List[Tuple[int, object, Optional[Tuple]]],
    unmatched: List[Tuple[int, str]],
    overflow: List[Tuple[int, str, Tuple]],
    ref_file: str,
    data_file: str,
    ref_col: str,
    data_crit_col: str,
    data_value_cols: List[str],
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = MATCHED_SHEET

    headers = [f"Reference_{ref_col}"] + [f"Data_{c}" for c in data_value_cols]
    write_header_row(ws, headers)

    label_col: List[object] = [headers[0]]
    value_cols: List[List[object]] = [[h] for h in headers[1:]]
    for row, raw, vals in output_rows:
        out_row = row + 1
        if raw is not None:
            ws.cell(row=out_row, column=1, value=raw)
            label_col.append(raw)
        if vals is not None:
            for j, v in enumerate(vals):
                if v is None:
                    continue
                write_value(ws, out_row, 2 + j, v)
                value_cols[j].append(v)

    ws.column_dimensions["A"].width = fit_width(label_col, MAX_TEXT_WIDTH)
    set_group_width(ws, 2, value_cols)

    if unmatched:
        write_unmatched_sheet(wb, unmatched, ref_file, ref_col)
    if overflow:
        write_overflow_sheet(wb, overflow, data_file, data_crit_col, data_value_cols)

    wb.save(path)


def write_unmatched_sheet(
    wb: Workbook,
    unmatched: List[Tuple[int, str]],
    ref_file: str,
    ref_col: str,
) -> None:
    ws = wb.create_sheet(UNMATCHED_SHEET)
    headers = ["Reference Cell", "Reference Label", "Source"]
    write_header_row(ws, headers)

    source_name = os.path.splitext(os.path.basename(ref_file))[0]
    cols: List[List[object]] = [[h] for h in headers]
    for i, (row, text) in enumerate(unmatched, start=2):
        values = [f"{ref_col}{row}", text, source_name]
        for j, v in enumerate(values):
            ws.cell(row=i, column=1 + j, value=v)
            cols[j].append(v)

    for j, values in enumerate(cols):
        ws.column_dimensions[get_column_letter(1 + j)].width = fit_width(values, MAX_TEXT_WIDTH)


def write_overflow_sheet(
    wb: Workbook,
    overflow: List[Tuple[int, str, Tuple]],
    data_file: str,
    data_crit_col: str,
    data_value_cols: List[str],
) -> None:
    ws = wb.create_sheet(OVERFLOW_SHEET)
    headers = (["Criteria Cell", "Data Criteria"]
               + [f"Data Col {c}" for c in data_value_cols] + ["Source"])
    write_header_row(ws, headers)

    source_name = os.path.splitext(os.path.basename(data_file))[0]
    n_val_cols = len(data_value_cols)
    cell_col: List[object] = [headers[0]]
    text_col: List[object] = [headers[1]]
    value_cols: List[List[object]] = [[h] for h in headers[2:2 + n_val_cols]]
    source_col: List[object] = [headers[-1]]
    for i, (row, text, vals) in enumerate(overflow, start=2):
        ws.cell(row=i, column=1, value=f"{data_crit_col}{row}")
        ws.cell(row=i, column=2, value=text)
        cell_col.append(f"{data_crit_col}{row}")
        text_col.append(text)
        for j, v in enumerate(vals):
            write_value(ws, i, 3 + j, v)
            value_cols[j].append(v)
        ws.cell(row=i, column=3 + n_val_cols, value=source_name)
        source_col.append(source_name)

    ws.column_dimensions["A"].width = fit_width(cell_col)
    ws.column_dimensions["B"].width = fit_width(text_col, MAX_TEXT_WIDTH)
    set_group_width(ws, 3, value_cols)
    ws.column_dimensions[get_column_letter(3 + n_val_cols)].width = fit_width(source_col)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def close_pause() -> None:
    print()
    print("-" * DIVIDER_WIDTH)
    input("[Enter] to close...")


def run() -> None:
    files = list_excel_files()
    if not files:
        clear_screen()
        print_banner()
        print("ERROR: no Excel files found in this folder")
        close_pause()
        return

    ref_file, ref_sheet, ref_cols, data_file, data_sheet, data_cols = select_sources(files)
    data_crit_col = data_cols[0]
    data_value_cols = data_cols[1:]
    n_data_cols = len(data_value_cols)

    start = time.time()
    clear_screen()
    print_banner()
    # Inputs lists files only (same file for both sources -> Inputs: 1); the
    # sheet/col choices follow as value lines.
    input_files = list(dict.fromkeys([ref_file, data_file]))
    print(f"Inputs: {len(input_files)}")
    for i, name in enumerate(input_files, start=1):
        print(f'   {i}. "{name}"')
    print()
    for line in value_lines([
        ("Reference sheet & col", f'"{ref_sheet}", {ref_cols[0]}'),
        ("Data sheet & cols",
         f'"{data_sheet}", criteria {data_crit_col}, data {", ".join(data_value_cols)}'),
    ]):
        print(line)

    print()
    print("Processing")
    print("Reading reference ...")
    ref_rows = read_columns(ref_file, ref_sheet, ref_cols)
    ref_last = last_nonempty_row(ref_rows, [0])
    print(f"Reference rows : {ref_last}")

    print("Reading data ...")
    data_rows = read_columns(data_file, data_sheet, data_cols)
    data_last = last_nonempty_row(data_rows, list(range(0, n_data_cols + 1)))
    print(f"Data rows : {data_last}")

    print("Matching ...")
    lookup, key_row = build_data_lookup(data_rows, n_data_cols)
    output_rows, unmatched, used_keys = build_output_rows(ref_rows, lookup, ref_last)
    overflow = find_overflow(lookup, key_row, used_keys)
    matched_count = sum(1 for _, _, vals in output_rows if vals is not None)
    print(f"Matched : {matched_count}")
    if unmatched:
        print(f'Unmatched reference : {len(unmatched)}   (sheet "{UNMATCHED_SHEET}")')
    if overflow:
        print(f'Data overflow : {len(overflow)}   (sheet "{OVERFLOW_SHEET}")')

    print("Writing output ...")
    out_path = build_output_filename(data_file, ref_sheet)
    write_output(out_path, output_rows, unmatched, overflow, ref_file, data_file,
                 ref_cols[0], data_crit_col, data_value_cols)

    elapsed = time.time() - start
    print()
    print("Outputs: 1")
    print(f"   1. \"{os.path.basename(out_path)}\"")

    print()
    print(f"Elapsed: {elapsed:.1f}s")
    close_pause()


def main() -> None:
    try:
        run()
    except UserExit:
        return
    except UserCancelled:
        clear_screen()
        print_banner()
        print("Cancelled.")
        close_pause()
    except Exception as e:
        print()
        print(f"ERROR: {type(e).__name__} - {e}")
        print(traceback.format_exc())
        close_pause()


if __name__ == "__main__":
    main()
