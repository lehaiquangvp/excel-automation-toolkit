"""
Excel Automation Toolkit - core functions.

Three tools:
  1. merge_files        - combine many Excel/CSV files into one workbook
  2. split_sheets       - save every sheet of a workbook as its own file
     split_by_column    - split one sheet into many by the values of a column
  3. remove_duplicates  - remove duplicate rows from every sheet

Every function returns a short text report so the GUI / CLI can show it.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}
_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_INVALID_FILE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# ---------------------------------------------------------------- helpers
def _read_all_sheets(path: Path) -> dict[str, pd.DataFrame]:
    """Return {sheet_name: DataFrame} for an Excel file, or one entry for CSV."""
    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {path.name}")
    if ext == ".csv":
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(path, encoding="latin-1")
        return {path.stem: df}
    return pd.read_excel(path, sheet_name=None)


def _safe_sheet_name(name: str, used: set[str]) -> str:
    """Excel sheet names: max 31 chars, no []:*?/\\ , unique (case-insensitive)."""
    base = _INVALID_SHEET_CHARS.sub("_", str(name)).strip("'").strip() or "Sheet"
    base = base[:31]
    candidate, i = base, 2
    while candidate.lower() in used:
        suffix = f" ({i})"
        candidate = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate.lower())
    return candidate


def _safe_file_name(name: str) -> str:
    cleaned = _INVALID_FILE_CHARS.sub("_", str(name)).strip(" .")
    return cleaned[:150] or "untitled"


def _unique_path(path: Path) -> Path:
    """Never overwrite an existing file: add (2), (3)... if needed."""
    if not path.exists():
        return path
    i = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def _write_workbook(sheets: dict[str, pd.DataFrame], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=_safe_sheet_name(name, used), index=False)
        if not sheets:  # Excel needs at least one sheet
            pd.DataFrame().to_excel(writer, sheet_name="Sheet1", index=False)


def collect_files(paths: list[str | Path]) -> list[Path]:
    """Accept files and/or folders; return supported files, sorted, no duplicates."""
    found: list[Path] = []
    for p in map(Path, paths):
        if p.is_dir():
            found += [f for f in sorted(p.iterdir())
                      if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS]
        elif p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
            found.append(p)
    # skip Excel lock files like "~$Report.xlsx"
    found = [f for f in found if not f.name.startswith("~$")]
    seen, result = set(), []
    for f in found:
        key = f.resolve()
        if key not in seen:
            seen.add(key)
            result.append(f)
    return result


# ---------------------------------------------------------------- 1. merge
def merge_files(paths: list[str | Path], output: str | Path,
                mode: str = "stack", add_source_column: bool = True) -> str:
    """
    mode="stack"  -> all rows from all sheets of all files in ONE sheet.
                     Columns are matched by header name.
    mode="sheets" -> every sheet of every file copied as a separate sheet.
    """
    files = collect_files(paths)
    if not files:
        raise ValueError("No Excel or CSV files found.")
    output = _unique_path(Path(output))

    total_rows, frames, sheets_out = 0, [], {}
    for f in files:
        for sheet, df in _read_all_sheets(f).items():
            total_rows += len(df)
            if mode == "stack":
                if add_source_column:
                    df = df.copy()
                    df.insert(0, "Source File", f.name)
                    df.insert(1, "Source Sheet", sheet)
                frames.append(df)
            else:
                label = f.stem if f.suffix.lower() == ".csv" else f"{f.stem}-{sheet}"
                sheets_out[label] = df

    if mode == "stack":
        merged = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
        _write_workbook({"Merged": merged}, output)
    elif mode == "sheets":
        _write_workbook(sheets_out, output)
    else:
        raise ValueError("mode must be 'stack' or 'sheets'")

    return (f"Merged {len(files)} file(s), {total_rows} data row(s).\n"
            f"Saved to: {output}")


# ---------------------------------------------------------------- 2. split
def split_sheets(path: str | Path, output_dir: str | Path) -> str:
    """Save every sheet of one workbook as its own .xlsx file."""
    path, output_dir = Path(path), Path(output_dir)
    sheets = _read_all_sheets(path)
    output_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for name, df in sheets.items():
        out = _unique_path(output_dir / f"{_safe_file_name(name)}.xlsx")
        _write_workbook({name: df}, out)
        written.append(out.name)
    return (f"Created {len(written)} file(s) in: {output_dir}\n  - "
            + "\n  - ".join(written))


def split_by_column(path: str | Path, column: str, output_dir: str | Path,
                    sheet: str | None = None, as_sheets: bool = False) -> str:
    """
    Split one sheet by the values in `column` (e.g. one file per Region).
    as_sheets=True -> one workbook with one sheet per value instead of many files.
    """
    path, output_dir = Path(path), Path(output_dir)
    sheets = _read_all_sheets(path)
    sheet = sheet or next(iter(sheets))
    if sheet not in sheets:
        raise ValueError(f"Sheet '{sheet}' not found. Sheets: {', '.join(sheets)}")
    df = sheets[sheet]
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Columns: "
                         + ", ".join(map(str, df.columns)))

    groups = {("(blank)" if pd.isna(k) else str(k)): g
              for k, g in df.groupby(column, dropna=False, sort=True)}
    output_dir.mkdir(parents=True, exist_ok=True)

    if as_sheets:
        out = _unique_path(output_dir / f"{path.stem} - by {_safe_file_name(column)}.xlsx")
        _write_workbook(groups, out)
        return f"Created 1 file with {len(groups)} sheet(s): {out}"

    written = []
    for value, g in groups.items():
        out = _unique_path(output_dir / f"{_safe_file_name(value)}.xlsx")
        _write_workbook({sheet: g}, out)
        written.append(out.name)
    return (f"Created {len(written)} file(s) in: {output_dir}\n  - "
            + "\n  - ".join(written))


# ---------------------------------------------------------------- 3. dedupe
def remove_duplicates(path: str | Path, output: str | Path,
                      columns: list[str] | None = None, keep: str = "first",
                      ignore_case_and_spaces: bool = True) -> str:
    """
    Remove duplicate rows from every sheet.
    columns: only compare these columns (e.g. ["Email"]); None = whole row.
    keep: "first" or "last" occurrence.
    ignore_case_and_spaces: "  John@Mail.com" equals "john@mail.com".
    Removed rows are saved in an extra sheet so nothing is lost.
    """
    if keep not in ("first", "last"):
        raise ValueError("keep must be 'first' or 'last'")
    path = Path(path)
    output = _unique_path(Path(output))
    sheets = _read_all_sheets(path)

    out_sheets, removed_frames, lines = {}, [], []
    for name, df in sheets.items():
        cols = columns or list(df.columns)
        cols = [c for c in cols if c in df.columns]
        if not cols:
            out_sheets[name] = df
            lines.append(f"  {name}: skipped (column not found)")
            continue

        key = df[cols].copy()
        if ignore_case_and_spaces:
            for c in cols:
                # works for both old (object) and new (string) pandas text columns
                key[c] = key[c].map(
                    lambda v: " ".join(v.split()).lower() if isinstance(v, str) else v)
        dup_mask = key.duplicated(keep=keep)

        out_sheets[name] = df[~dup_mask]
        if dup_mask.any():
            removed = df[dup_mask].copy()
            removed.insert(0, "From Sheet", name)
            removed_frames.append(removed)
        lines.append(f"  {name}: {len(df)} rows -> {len(df) - int(dup_mask.sum())} "
                     f"({int(dup_mask.sum())} duplicate(s) removed)")

    if removed_frames:
        out_sheets["Removed Duplicates"] = pd.concat(removed_frames, ignore_index=True)
    _write_workbook(out_sheets, output)
    return "Duplicate check finished:\n" + "\n".join(lines) + f"\nSaved to: {output}"


def list_columns(path: str | Path, sheet: str | None = None) -> list[str]:
    sheets = _read_all_sheets(Path(path))
    sheet = sheet or next(iter(sheets))
    return [str(c) for c in sheets[sheet].columns]


def list_sheets(path: str | Path) -> list[str]:
    return list(_read_all_sheets(Path(path)).keys())
