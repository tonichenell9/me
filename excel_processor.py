#!/usr/bin/env python3
"""
Excel Processor Module
Handles Excel file manipulation using win32com for robust Windows automation.
"""

import logging
import os
import time
from pathlib import Path
from typing import Optional

# Try to import win32com for Excel COM interface (Windows only)
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    # We will log a warning in __init__ if this is missing on a Windows machine


class ExcelProcessor:
    """Handles Excel file manipulation using Windows COM automation."""
    
    def __init__(self, config: dict):
        """
        Initialize Excel processor with configuration.
        """
        self.config = config
        self.excel_config = config.get('excel', {})
        self.worksheet_1_name = self.excel_config.get('worksheet_1_name', 'large deal report')
        self.worksheet_2_name = self.excel_config.get('worksheet_2_name', 'summary')
        self.worksheet_3_name = self.excel_config.get('worksheet_3_name', 'iphone compatible')
        self.logger = logging.getLogger(__name__)

        if os.name == 'nt' and not WIN32_AVAILABLE:
            self.logger.error("win32com.client is not available. Please install pywin32.")

    def process_report(self, master_workbook_path: Path, daily_sheet_path: Path, output_path: Path) -> None:
        """
        Executes the full Excel workflow:
        1. Open Master Workbook
        2. Copy Daily Sheet content -> Master Worksheet 1
        3. Refresh Master Worksheet 2 (Summary)
        4. Copy Master Worksheet 2 -> Master Worksheet 3 (Values Only)
        5. Save Master as Output Path
        """
        if not WIN32_AVAILABLE:
            raise RuntimeError("This script requires pywin32 to run on Windows.")

        excel = None
        wb_master = None
        wb_daily = None

        try:
            self.logger.info("Starting Excel application...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Keep it hidden during processing
            excel.DisplayAlerts = False

            # 1. Open Master Workbook
            self.logger.info(f"Opening master workbook: {master_workbook_path}")
            if not master_workbook_path.exists():
                raise FileNotFoundError(f"Master workbook not found at {master_workbook_path}")
            
            wb_master = excel.Workbooks.Open(str(master_workbook_path))

            # 2. Open Daily Workbook and Copy Data
            self.logger.info(f"Opening daily sheet: {daily_sheet_path}")
            wb_daily = excel.Workbooks.Open(str(daily_sheet_path))
            
            # Assuming data is in the first sheet of the daily workbook
            ws_daily_source = wb_daily.Sheets(1)
            
            # Target sheet in Master
            try:
                ws_master_target = wb_master.Sheets(self.worksheet_1_name)
            except Exception:
                self.logger.error(f"Worksheet '{self.worksheet_1_name}' not found in master workbook.")
                raise

            self.logger.info(f"Copying data to '{self.worksheet_1_name}'...")
            
            # Clear existing data in target sheet (optional, but good practice)
            ws_master_target.Cells.Clear()
            
            # Copy UsedRange from Daily to Master A1
            ws_daily_source.UsedRange.Copy()
            ws_master_target.Range("A1").PasteSpecial(Paste=-4104) # xlPasteAll
            
            # Close daily workbook
            wb_daily.Close(SaveChanges=False)
            wb_daily = None

            # 3. Refresh Summary Tab
            self.logger.info(f"Refreshing '{self.worksheet_2_name}'...")
            try:
                ws_summary = wb_master.Sheets(self.worksheet_2_name)
                # Refresh all data connections in the workbook is usually safest/easiest
                wb_master.RefreshAll()
                
                # Wait for refresh to complete (if background queries are enabled)
                # Ideally, we should check BackgroundQuery property, but a small sleep helps
                time.sleep(5) 
                
            except Exception as e:
                self.logger.error(f"Error refreshing summary: {e}")
                raise

            # 4. Copy Summary to iPhone Compatible (Values Only)
            self.logger.info(f"Copying '{self.worksheet_2_name}' to '{self.worksheet_3_name}' (Values Only)...")
            try:
                ws_iphone = wb_master.Sheets(self.worksheet_3_name)
            except Exception:
                self.logger.warning(f"Worksheet '{self.worksheet_3_name}' not found. Creating it.")
                ws_iphone = wb_master.Sheets.Add(After=wb_master.Sheets(wb_master.Sheets.Count))
                ws_iphone.Name = self.worksheet_3_name
            
            # Clear target
            ws_iphone.Cells.Clear()

            # Copy Summary Table/UsedRange
            # User said "copy that entire table". We'll assume UsedRange for now.
            ws_summary.UsedRange.Copy()
            
            # Paste Values Only
            # xlPasteValues = -4163
            ws_iphone.Range("A1").PasteSpecial(Paste=-4163)
            
            # Optional: Paste Formats to keep it looking decent? User said "values only", strictly implies stripping formulas/formats.
            # But usually for reports you want column widths at least.
            # ws_iphone.Range("A1").PasteSpecial(Paste=-4122) # xlPasteColumnWidths

            # 5. Save As Output
            self.logger.info(f"Saving report to: {output_path}")
            # xlOpenXMLWorkbook = 51 (for .xlsx)
            wb_master.SaveAs(str(output_path), FileFormat=51)

        except Exception as e:
            self.logger.error(f"Excel processing error: {str(e)}")
            raise

        finally:
            self.logger.info("Cleaning up Excel resources...")
            if wb_daily:
                try:
                    wb_daily.Close(SaveChanges=False)
                except:
                    pass
            
            if wb_master:
                try:
                    wb_master.Close(SaveChanges=False) # Already saved as new file
                except:
                    pass
            
            if excel:
                try:
                    excel.Quit()
                except:
                    pass

