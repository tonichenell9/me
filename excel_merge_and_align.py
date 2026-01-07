#!/usr/bin/env python3
"""
Merge two Excel spreadsheets into one, then optionally merge+center identical values per column.

Default behavior:
- Reads the first sheet of each workbook
- Uses row 1 as headers
- Appends rows from file A then file B
- Writes a new workbook
- Merges consecutive identical values vertically within each column and centers them
"""

from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl

from excel_processor import (
    join_tables_on_keys,
    merge_identical_cells_cascading,
    merge_tables_by_headers,
    read_worksheet_as_table,
    sort_rows_by_columns,
    write_table_to_workbook,
)


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Merge two Excel files and optionally merge+center identical values in columns."
    )
    p.add_argument("--file-a", required=True, type=Path, help="First input .xlsx")
    p.add_argument("--file-b", required=True, type=Path, help="Second input .xlsx")
    p.add_argument("--out", required=True, type=Path, help="Output .xlsx path")
    p.add_argument("--sheet-a", default=None, help="Worksheet name in file A (default: active)")
    p.add_argument("--sheet-b", default=None, help="Worksheet name in file B (default: active)")
    p.add_argument("--out-sheet", default="Merged", help="Output worksheet name (default: Merged)")
    p.add_argument(
        "--merge-on",
        default="",
        help=(
            "Comma-separated column name(s) to join on (must exist in BOTH files). "
            "Example: \"Entity ID\" or \"Entity ID,Entity Name\". "
            "If omitted, rows are appended (A then B) instead of joined."
        ),
    )
    p.add_argument(
        "--how",
        default="outer",
        choices=["left", "right", "inner", "outer"],
        help="Join type when using --merge-on (default: outer).",
    )
    p.add_argument(
        "--no-coalesce-overlaps",
        action="store_true",
        help=(
            "When joining, do NOT coalesce overlapping non-key columns. "
            "Instead, keep both with _A/_B suffixes."
        ),
    )
    p.add_argument(
        "--prefer",
        default="a",
        choices=["a", "b"],
        help="When coalescing overlapping columns, prefer values from A or B (default: a).",
    )
    p.add_argument(
        "--include-source",
        action="store_true",
        help="Add a __source__ column with A/B indicating which file the row came from.",
    )
    p.add_argument(
        "--deduplicate",
        action="store_true",
        help="Drop exact duplicate rows after merging (based on all output columns).",
    )
    p.add_argument(
        "--no-merge-identical",
        action="store_true",
        help="Do not merge+center identical consecutive values per column.",
    )
    p.add_argument(
        "--merge-columns",
        default="all",
        help=(
            "Which columns to merge (1-based indices or header names). "
            "Examples: 'all', '1,2,5', 'Status,Owner'. Default: all"
        ),
    )
    p.add_argument(
        "--sort-by",
        default="",
        help=(
            "Comma-separated column name(s) to sort by before formatting (recommended for grouping). "
            "If omitted and --merge-on is provided, it defaults to the same columns as --merge-on."
        ),
    )
    p.add_argument("--header-row", type=int, default=1, help="Header row index (default: 1)")
    p.add_argument("--data-start-row", type=int, default=2, help="Data start row index (default: 2)")
    return p.parse_args()


def _resolve_merge_columns(merge_columns: str, headers: list[str]) -> list[int] | None:
    merge_columns = (merge_columns or "all").strip()
    if merge_columns.lower() == "all":
        return None

    tokens = [t.strip() for t in merge_columns.split(",") if t.strip()]
    if not tokens:
        return None

    # Prefer numeric indices if they look numeric; else treat as header names.
    cols: list[int] = []
    for t in tokens:
        if t.isdigit():
            idx = int(t)
            if idx < 1:
                continue
            cols.append(idx)
        else:
            if t in headers:
                cols.append(headers.index(t) + 1)
    # De-dupe while preserving order
    seen: set[int] = set()
    out: list[int] = []
    for c in cols:
        if c in seen:
            continue
        seen.add(c)
        out.append(c)
    return out or None


def main() -> None:
    args = _parse_args()

    if not args.file_a.exists():
        raise SystemExit(f"file-a not found: {args.file_a}")
    if not args.file_b.exists():
        raise SystemExit(f"file-b not found: {args.file_b}")

    table_a = read_worksheet_as_table(
        args.file_a, args.sheet_a, header_row=args.header_row, data_start_row=args.data_start_row
    )
    table_b = read_worksheet_as_table(
        args.file_b, args.sheet_b, header_row=args.header_row, data_start_row=args.data_start_row
    )

    merge_on = [c.strip() for c in (args.merge_on or "").split(",") if c.strip()]
    if merge_on:
        headers, rows = join_tables_on_keys(
            table_a,
            table_b,
            keys=merge_on,
            how=args.how,
            coalesce_overlaps=not args.no_coalesce_overlaps,
            prefer=args.prefer,
        )
    else:
        headers, rows = merge_tables_by_headers(
            table_a,
            table_b,
            include_source_column=bool(args.include_source),
            deduplicate_rows=bool(args.deduplicate),
        )

    sort_by = [c.strip() for c in (args.sort_by or "").split(",") if c.strip()]
    if not sort_by and merge_on:
        sort_by = list(merge_on)
    if sort_by:
        rows = sort_rows_by_columns(headers, rows, sort_by=sort_by)

    write_table_to_workbook(headers, rows, args.out, sheet_name=args.out_sheet)

    if not args.no_merge_identical:
        wb = openpyxl.load_workbook(args.out)
        try:
            ws = wb[args.out_sheet]
            columns = _resolve_merge_columns(args.merge_columns, list(headers))
            merge_identical_cells_cascading(
                ws,
                columns=columns,
                header_row=args.header_row,
                data_start_row=args.data_start_row,
            )
            wb.save(args.out)
        finally:
            wb.close()


if __name__ == "__main__":
    main()

