#!/usr/bin/env python3
"""
Excel Processor Module
Handles Excel file manipulation, worksheet updates, refresh operations, and copying.
"""

import logging
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
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws.append(list(headers))
    for r in rows:
        ws.append([r.get(h) for h in headers])

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
