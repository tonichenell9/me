#!/usr/bin/env python3
"""
Excel Processor Module
Handles Excel file manipulation, worksheet updates, refresh operations, and copying.
"""

import logging
import math
from pathlib import Path
from typing import Iterable, Optional, Sequence
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Try to import xlwings for Excel refresh functionality
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    logging.warning("xlwings not available. Worksheet 2 refresh will use fallback method.")

DEFAULT_MERGED_CELL_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _normalize_header(value: object, fallback: str) -> str:
    """Convert a header cell value into a usable string."""
    if value is None:
        return fallback
    text = str(value).strip()
    return text if text else fallback


def read_worksheet_as_table(
    workbook_path: Path,
    sheet_name: Optional[str] = None,
    *,
    header_row: int = 1,
    data_start_row: int = 2,
) -> tuple[list[str], list[dict[str, object]]]:
    """
    Read an Excel worksheet into a simple table.

    - Uses the `header_row` to create column names.
    - Returns a list of row dicts keyed by those headers.
    - Skips completely empty rows.
    """
    wb = load_workbook(workbook_path, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        raw_headers: list[object] = [
            ws.cell(row=header_row, column=col).value for col in range(1, ws.max_column + 1)
        ]
        headers: list[str] = []
        seen: set[str] = set()
        for idx, raw in enumerate(raw_headers, start=1):
            base = _normalize_header(raw, f"Column{idx}")
            name = base
            suffix = 2
            while name in seen:
                name = f"{base}_{suffix}"
                suffix += 1
            seen.add(name)
            headers.append(name)

        rows: list[dict[str, object]] = []
        for row in ws.iter_rows(
            min_row=data_start_row,
            max_row=ws.max_row,
            max_col=len(headers),
            values_only=True,
        ):
            if not row or not any(cell is not None and str(cell).strip() != "" for cell in row):
                continue
            rows.append({headers[i]: row[i] for i in range(len(headers))})

        return headers, rows
    finally:
        wb.close()


def merge_tables_by_headers(
    table_a: tuple[Sequence[str], Sequence[dict[str, object]]],
    table_b: tuple[Sequence[str], Sequence[dict[str, object]]],
    *,
    include_source_column: bool = False,
    source_header: str = "__source__",
    source_a_value: str = "A",
    source_b_value: str = "B",
    deduplicate_rows: bool = False,
) -> tuple[list[str], list[dict[str, object]]]:
    """
    Merge two tables by taking the union of headers and appending all rows.

    This is a simple "stack rows" merge (not a key-based join).
    """
    headers_a, rows_a = table_a
    headers_b, rows_b = table_b

    headers: list[str] = []
    for h in list(headers_a) + [h for h in headers_b if h not in headers_a]:
        hs = str(h)
        if hs not in headers:
            headers.append(hs)

    if include_source_column and source_header not in headers:
        headers = [source_header] + headers

    merged_rows: list[dict[str, object]] = []

    def _append_rows(rows: Sequence[dict[str, object]], source_value: str) -> None:
        for r in rows:
            out: dict[str, object] = {}
            if include_source_column:
                out[source_header] = source_value
            for h in headers:
                if include_source_column and h == source_header:
                    continue
                out[h] = r.get(h)
            merged_rows.append(out)

    _append_rows(rows_a, source_a_value)
    _append_rows(rows_b, source_b_value)

    if deduplicate_rows:
        seen: set[tuple[object, ...]] = set()
        deduped: list[dict[str, object]] = []
        for r in merged_rows:
            key = tuple(r.get(h) for h in headers)
            if key in seen:
                continue
            seen.add(key)
            deduped.append(r)
        merged_rows = deduped

    return headers, merged_rows


def write_table_to_workbook(
    headers: Sequence[str],
    rows: Sequence[dict[str, object]],
    output_path: Path,
    *,
    sheet_name: str = "Merged",
    autofit_columns: bool = True,
) -> None:
    """Write a header+row-dicts table to a new xlsx workbook."""
    def _to_excel_value(v: object) -> object:
        # pandas/numpy scalars -> native python
        if hasattr(v, "to_pydatetime"):
            try:
                return v.to_pydatetime()
            except Exception:
                pass
        if hasattr(v, "item"):
            try:
                v = v.item()
            except Exception:
                pass
        # pandas uses NaN / NA for missing
        try:
            if isinstance(v, float) and math.isnan(v):
                return None
        except Exception:
            pass
        return v

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws.append(list(headers))
    for r in rows:
        ws.append([_to_excel_value(r.get(h)) for h in headers])

    if autofit_columns and ws.max_row >= 1:
        for col_idx, header in enumerate(headers, start=1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            max_len = len(str(header)) if header is not None else 0
            for row_idx in range(2, ws.max_row + 1):
                v = ws.cell(row=row_idx, column=col_idx).value
                if v is None:
                    continue
                max_len = max(max_len, len(str(v)))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()


def merge_identical_cells_in_columns(
    worksheet: openpyxl.worksheet.worksheet.Worksheet,
    *,
    columns: Optional[Iterable[int]] = None,
    header_row: int = 1,
    data_start_row: int = 2,
    ignore_blanks: bool = True,
    alignment: Alignment = DEFAULT_MERGED_CELL_ALIGNMENT,
) -> None:
    """
    Merge consecutive identical values vertically within each column and center them.

    This is a formatting operation only; it does not change the underlying values.
    """
    if worksheet.max_row < data_start_row:
        return

    if columns is None:
        columns = range(1, worksheet.max_column + 1)

    def _is_blank(v: object) -> bool:
        if v is None:
            return True
        if isinstance(v, str) and v.strip() == "":
            return True
        return False

    for col in columns:
        run_start = data_start_row
        prev = worksheet.cell(row=data_start_row, column=col).value

        for row in range(data_start_row + 1, worksheet.max_row + 2):  # +2 to flush last run
            cur = worksheet.cell(row=row, column=col).value if row <= worksheet.max_row else object()

            prev_blank = _is_blank(prev) if ignore_blanks else False
            cur_blank = _is_blank(cur) if ignore_blanks else False
            same = (cur == prev) and not prev_blank and not cur_blank
            if same:
                continue

            run_end = row - 1
            if run_end > run_start and not prev_blank:
                worksheet.merge_cells(
                    start_row=run_start,
                    start_column=col,
                    end_row=run_end,
                    end_column=col,
                )
                # Apply alignment to the merged region (top-left cell drives display)
                for r in range(run_start, run_end + 1):
                    worksheet.cell(row=r, column=col).alignment = alignment

            run_start = row
            prev = cur

    # Center header row too (nice default)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(1, worksheet.max_column + 1):
        worksheet.cell(row=header_row, column=col).alignment = header_alignment


def merge_identical_cells_cascading(
    worksheet: openpyxl.worksheet.worksheet.Worksheet,
    *,
    columns: Optional[Iterable[int]] = None,
    header_row: int = 1,
    data_start_row: int = 2,
    ignore_blanks: bool = True,
    alignment: Alignment = DEFAULT_MERGED_CELL_ALIGNMENT,
) -> None:
    """
    Cascading (hierarchical) merge+center from left to right.

    - Column 1 merges identical consecutive values.
    - Column 2 merges identical consecutive values *within* each Column-1 merged group.
    - Column 3 merges within each (Column-1, Column-2) group, etc.

    This is typically what you want for visually grouping rows by multiple columns.
    """
    if worksheet.max_row < data_start_row:
        return

    if columns is None:
        columns = range(1, worksheet.max_column + 1)

    cols = sorted(list(columns))
    if not cols:
        return

    def _is_blank(v: object) -> bool:
        if v is None:
            return True
        if isinstance(v, str) and v.strip() == "":
            return True
        return False

    max_row = worksheet.max_row

    for col in cols:
        run_start = data_start_row
        prev_val = worksheet.cell(row=data_start_row, column=col).value
        prev_group = tuple(
            worksheet.cell(row=data_start_row, column=c).value for c in cols if c < col
        )

        for row in range(data_start_row + 1, max_row + 2):  # +2 to flush last run
            if row <= max_row:
                cur_val = worksheet.cell(row=row, column=col).value
                cur_group = tuple(worksheet.cell(row=row, column=c).value for c in cols if c < col)
            else:
                cur_val = object()
                cur_group = object()  # force flush

            prev_blank = _is_blank(prev_val) if ignore_blanks else False
            cur_blank = _is_blank(cur_val) if ignore_blanks else False

            same_value = (cur_val == prev_val) and not prev_blank and not cur_blank
            same_group = cur_group == prev_group

            if same_value and same_group:
                continue

            run_end = row - 1
            if run_end > run_start and not prev_blank:
                worksheet.merge_cells(
                    start_row=run_start,
                    start_column=col,
                    end_row=run_end,
                    end_column=col,
                )
                for r in range(run_start, run_end + 1):
                    worksheet.cell(row=r, column=col).alignment = alignment

            run_start = row
            prev_val = cur_val
            prev_group = cur_group

    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(1, worksheet.max_column + 1):
        worksheet.cell(row=header_row, column=col).alignment = header_alignment


def join_tables_on_keys(
    table_a: tuple[Sequence[str], Sequence[dict[str, object]]],
    table_b: tuple[Sequence[str], Sequence[dict[str, object]]],
    *,
    keys: Sequence[str],
    how: str = "outer",
    coalesce_overlaps: bool = True,
    prefer: str = "a",
    suffix_a: str = "_A",
    suffix_b: str = "_B",
) -> tuple[list[str], list[dict[str, object]]]:
    """
    Key-based merge ("join") of two tables on specific column names.

    - `keys` must exist in both inputs.
    - `how` is one of: left, right, inner, outer
    - If `coalesce_overlaps` is True, overlapping non-key columns are combined into one:
      prefer A unless A is blank, otherwise take B (or vice versa if prefer="b").
    """
    import pandas as pd

    headers_a, rows_a = table_a
    headers_b, rows_b = table_b
    keys = [k for k in (keys or []) if str(k).strip()]
    if not keys:
        raise ValueError("keys must be provided for a key-based merge")

    df_a = pd.DataFrame(list(rows_a)).reindex(columns=list(headers_a))
    df_b = pd.DataFrame(list(rows_b)).reindex(columns=list(headers_b))

    missing_a = [k for k in keys if k not in df_a.columns]
    missing_b = [k for k in keys if k not in df_b.columns]
    if missing_a or missing_b:
        raise ValueError(
            "Missing merge keys. "
            f"Missing in A: {missing_a or 'none'}, missing in B: {missing_b or 'none'}"
        )

    overlap = (set(df_a.columns) & set(df_b.columns)) - set(keys)

    merged = pd.merge(
        df_a,
        df_b,
        on=keys,
        how=how,
        suffixes=(suffix_a, suffix_b),
    )

    def _is_blank_value(v: object) -> bool:
        if v is None:
            return True
        # pandas NA/NaN
        try:
            if isinstance(v, float) and math.isnan(v):
                return True
        except Exception:
            pass
        if isinstance(v, str) and v.strip() == "":
            return True
        return False

    if coalesce_overlaps and overlap:
        for col in overlap:
            a_col = f"{col}{suffix_a}"
            b_col = f"{col}{suffix_b}"
            if a_col not in merged.columns or b_col not in merged.columns:
                continue

            a_vals = merged[a_col]
            b_vals = merged[b_col]

            a_blank = a_vals.map(_is_blank_value)
            b_blank = b_vals.map(_is_blank_value)

            if prefer.lower() == "b":
                merged[col] = b_vals.where(~b_blank, a_vals)
            else:
                merged[col] = a_vals.where(~a_blank, b_vals)

            merged = merged.drop(columns=[a_col, b_col])

    # Order columns: keys, then A columns, then B-only columns (preserve original orders).
    out_headers: list[str] = []
    for k in keys:
        if k in merged.columns and k not in out_headers:
            out_headers.append(k)

    # A columns (excluding keys). If coalesced, the base name exists; otherwise suffixed exists.
    for col in headers_a:
        if col in keys:
            continue
        if coalesce_overlaps and col in overlap and col in merged.columns and col not in out_headers:
            out_headers.append(col)
        elif not coalesce_overlaps:
            name = col if col not in overlap else f"{col}{suffix_a}"
            if name in merged.columns and name not in out_headers:
                out_headers.append(name)
        else:
            if col in merged.columns and col not in out_headers:
                out_headers.append(col)

    # B-only columns (excluding keys). If not coalescing, include suffixed overlaps too.
    for col in headers_b:
        if col in keys:
            continue
        if coalesce_overlaps and col in overlap:
            # already handled
            continue
        if not coalesce_overlaps and col in overlap:
            name = f"{col}{suffix_b}"
        else:
            name = col
        if name in merged.columns and name not in out_headers:
            out_headers.append(name)

    out_rows = merged[out_headers].to_dict(orient="records")
    return out_headers, out_rows


def sort_rows_by_columns(
    headers: Sequence[str],
    rows: Sequence[dict[str, object]],
    *,
    sort_by: Sequence[str],
) -> list[dict[str, object]]:
    """Stable-sort table rows by the given column names."""
    import pandas as pd

    sort_cols = [c for c in sort_by if c in headers]
    if not sort_cols:
        return list(rows)

    df = pd.DataFrame(list(rows)).reindex(columns=list(headers))
    df = df.sort_values(by=sort_cols, kind="mergesort", na_position="last")
    return df.to_dict(orient="records")


class ExcelProcessor:
    """Handles Excel file manipulation and processing operations."""
    
    def __init__(self, config: dict):
        """
        Initialize Excel processor with configuration.
        
        Args:
            config: Dictionary containing configuration:
                - worksheet_1_name: Name of worksheet 1
                - worksheet_2_name: Name of worksheet 2
                - worksheet_3_name: Name of worksheet 3
        """
        self.config = config
        self.worksheet_1_name = config.get('worksheet_1_name', 'Sheet1')
        self.worksheet_2_name = config.get('worksheet_2_name', 'Sheet2')
        self.worksheet_3_name = config.get('worksheet_3_name', 'Sheet3')
        self.logger = logging.getLogger(__name__)
    
    def load_workbook(self, filepath: Path):
        """
        Load an Excel workbook.
        
        Args:
            filepath: Path to the Excel file
            
        Returns:
            openpyxl Workbook object
            
        Raises:
            FileNotFoundError: If file doesn't exist
            Exception: If file cannot be opened
        """
        try:
            self.logger.info(f"Loading workbook: {filepath}")
            return load_workbook(filepath)
        except FileNotFoundError:
            raise FileNotFoundError(f"Workbook not found: {filepath}")
        except Exception as e:
            raise Exception(f"Error loading workbook {filepath}: {str(e)}")
    
    def update_worksheet_1(self, workbook: openpyxl.Workbook, daily_sheet_path: Path) -> None:
        """
        Replace worksheet 1 with data from the daily email attachment.
        
        Args:
            workbook: The target workbook to update
            daily_sheet_path: Path to the daily worksheet file
            
        Raises:
            FileNotFoundError: If daily sheet file doesn't exist
            ValueError: If worksheet 1 doesn't exist
            Exception: If update fails
        """
        self.logger.info("Updating Worksheet 1...")
        
        try:
            # Load the daily sheet
            if not daily_sheet_path.exists():
                raise FileNotFoundError(f"Daily sheet not found: {daily_sheet_path}")
            
            daily_wb = load_workbook(daily_sheet_path)
            daily_ws = daily_wb.active
            
            # Get or create worksheet 1
            if self.worksheet_1_name in workbook.sheetnames:
                ws1 = workbook[self.worksheet_1_name]
            else:
                self.logger.warning(
                    f"Worksheet '{self.worksheet_1_name}' not found. Using active sheet."
                )
                ws1 = workbook.active
            
            # Clear existing data in worksheet 1
            if ws1.max_row > 0:
                ws1.delete_rows(1, ws1.max_row)
            
            # Copy data from daily sheet to worksheet 1
            for row in daily_ws.iter_rows(values_only=True):
                if any(cell is not None for cell in row):  # Skip completely empty rows
                    ws1.append(row)
            
            # Copy formatting if needed (basic cell copying)
            for row_idx, row in enumerate(daily_ws.iter_rows(), start=1):
                if row_idx > ws1.max_row:
                    break
                for col_idx, cell in enumerate(row, start=1):
                    if cell.has_style:
                        target_cell = ws1.cell(row=row_idx, column=col_idx)
                        target_cell.font = cell.font
                        target_cell.fill = cell.fill
                        target_cell.border = cell.border
                        target_cell.alignment = cell.alignment
                        target_cell.number_format = cell.number_format
            
            daily_wb.close()
            self.logger.info("Worksheet 1 updated successfully.")
            
        except Exception as e:
            self.logger.error(f"Error updating Worksheet 1: {str(e)}")
            raise
    
    def refresh_worksheet_2(self, workbook: openpyxl.Workbook, workbook_path: Optional[Path] = None) -> None:
        """
        Refresh data in worksheet 2 using xlwings (if available) or fallback method.
        
        Args:
            workbook: The workbook containing worksheet 2
            workbook_path: Optional path to the workbook file (required for xlwings)
            
        Raises:
            ValueError: If worksheet 2 doesn't exist
            Exception: If refresh fails
        """
        self.logger.info("Refreshing Worksheet 2...")
        
        if self.worksheet_2_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found in workbook")
        
        # Try xlwings refresh if available and workbook_path is provided
        if XLWINGS_AVAILABLE and workbook_path:
            try:
                self._refresh_with_xlwings(workbook_path)
                self.logger.info("Worksheet 2 refreshed using xlwings.")
                return
            except Exception as e:
                self.logger.warning(
                    f"xlwings refresh failed: {str(e)}. Falling back to openpyxl method."
                )
        
        # Fallback: Use openpyxl (preserves structure but doesn't refresh connections)
        self.logger.info(
            "Using openpyxl fallback. Note: Power Query and external data connections "
            "will not be automatically refreshed. You may need to manually refresh in Excel."
        )
        # The worksheet structure is preserved, but data connections won't refresh
        # This is expected behavior when xlwings is not available
    
    def _refresh_with_xlwings(self, workbook_path: Path) -> None:
        """
        Refresh worksheet 2 using xlwings (requires Excel installed).
        Works on both Windows and macOS.
        
        Args:
            workbook_path: Path to the workbook file
            
        Raises:
            Exception: If xlwings refresh fails
        """
        app = None
        wb = None
        
        try:
            self.logger.info("Attempting to refresh using xlwings...")
            
            # Open Excel application (headless)
            # On Windows, this will use the installed Excel application
            # On macOS, this will use Excel for Mac
            app = xw.App(visible=False)
            app.display_alerts = False
            
            # Open the workbook
            wb = app.books.open(str(workbook_path))
            
            # Get worksheet 2
            if self.worksheet_2_name not in [sheet.name for sheet in wb.sheets]:
                raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found")
            
            ws = wb.sheets[self.worksheet_2_name]
            
            # Try to refresh all data connections/queries
            try:
                # Refresh all queries/connections in the workbook
                wb.api.RefreshAll()
                self.logger.info("Refreshed all data connections in workbook.")
            except Exception as e:
                self.logger.warning(f"Could not refresh all connections: {str(e)}")
                # Try to refresh specific table if it's a ListObject
                try:
                    if ws.api.ListObjects.Count > 0:
                        for list_obj in ws.api.ListObjects:
                            if hasattr(list_obj, 'QueryTable'):
                                list_obj.QueryTable.Refresh(BackgroundQuery=False)
                        self.logger.info("Refreshed table connections.")
                except Exception as e2:
                    self.logger.warning(f"Could not refresh table connections: {str(e2)}")
            
            # Save the workbook
            wb.save()
            self.logger.info("Workbook saved after refresh.")
            
        except Exception as e:
            self.logger.error(f"xlwings refresh error: {str(e)}")
            raise
        finally:
            # Clean up
            if wb:
                try:
                    wb.close()
                except:
                    pass
            if app:
                try:
                    app.quit()
                except:
                    pass
    
    def copy_worksheet_2_to_3(self, workbook: openpyxl.Workbook) -> None:
        """
        Copy worksheet 2 data to worksheet 3 (smartphone-friendly version).
        
        Args:
            workbook: The workbook containing the worksheets
            
        Raises:
            ValueError: If worksheet 2 doesn't exist
            Exception: If copy operation fails
        """
        self.logger.info("Copying Worksheet 2 to Worksheet 3...")
        
        if self.worksheet_2_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found in workbook")
        
        try:
            ws2 = workbook[self.worksheet_2_name]
            
            # Remove existing worksheet 3 if it exists
            if self.worksheet_3_name in workbook.sheetnames:
                workbook.remove(workbook[self.worksheet_3_name])
            
            # Create new worksheet 3
            ws3 = workbook.create_sheet(self.worksheet_3_name)
            
            # Copy data from worksheet 2 to worksheet 3
            for row in ws2.iter_rows(values_only=True):
                if any(cell is not None for cell in row):  # Skip completely empty rows
                    ws3.append(row)
            
            # Copy formatting
            for row_idx, row in enumerate(ws2.iter_rows(), start=1):
                if row_idx > ws3.max_row:
                    break
                for col_idx, cell in enumerate(row, start=1):
                    if cell.has_style:
                        target_cell = ws3.cell(row=row_idx, column=col_idx)
                        target_cell.font = cell.font
                        target_cell.fill = cell.fill
                        target_cell.border = cell.border
                        target_cell.alignment = cell.alignment
                        target_cell.number_format = cell.number_format
            
            # Apply smartphone-friendly formatting (optional enhancements)
            # Adjust column widths for better mobile viewing
            for column in ws3.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 for mobile
                ws3.column_dimensions[column_letter].width = adjusted_width
            
            self.logger.info("Worksheet 3 created successfully.")
            
        except Exception as e:
            self.logger.error(f"Error copying Worksheet 2 to Worksheet 3: {str(e)}")
            raise
    
    def save_workbook(self, workbook: openpyxl.Workbook, filepath: Path) -> None:
        """
        Save the workbook to a file.
        
        Args:
            workbook: The workbook to save
            filepath: Path where to save the workbook
            
        Raises:
            Exception: If save fails
        """
        try:
            self.logger.info(f"Saving workbook to: {filepath}")
            workbook.save(filepath)
            self.logger.info("Workbook saved successfully.")
        except Exception as e:
            self.logger.error(f"Error saving workbook: {str(e)}")
            raise
