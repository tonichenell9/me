#!/usr/bin/env python3
"""
Large Deal Report Automation Script
Automates the daily generation and distribution of the Large Deal Report.

Workflow:
1. Download the previous day's report from email (from Large Deal Reports folder)
2. Download the daily worksheet from "large trade td" email
3. Paste data into 'large deal report' sheet starting at A1
4. Refresh the 'summary' sheet (right-click refresh on table)
5. Copy summary table to 'iphone compatible' (paste values only)
6. Save as "large deal report - {today's date}.xlsx"
7. Open email preview in Outlook for review before sending
"""

import os
import sys
import logging
from datetime import datetime
from pathlib import Path

from config_reader import read_config
from email_handler import EmailHandler
from excel_processor import ExcelProcessor


class LargeDealReportAutomation:
    """Main automation class that orchestrates the report generation workflow."""
    
    def __init__(self, config_path: str = 'config.txt'):
        """
        Initialize the automation with configuration.
        
        Args:
            config_path: Path to the configuration file
        """
        # Setup basic logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(levelname)s: %(message)s',
            handlers=[logging.StreamHandler(sys.stdout)]
        )
        self.logger = logging.getLogger(__name__)
        
        # Load configuration from config.txt
        self.config = read_config(config_path)
        
        # Setup directories - use absolute paths
        self.reports_dir = Path(self.config.get('reports_directory', './reports')).absolute()
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir = Path(self.config.get('temp_directory', './temp')).absolute()
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        
        self.logger.info(f"Reports folder: {self.reports_dir}")
        self.logger.info(f"Temp folder: {self.temp_dir}")
        
        # Date formats
        now = datetime.now()
        # Long date format: "8 January 2026" (no leading zero)
        self.current_date_long = f"{now.day} {now.strftime('%B %Y')}"
        self.current_date_file = self.current_date_long  # For filenames: "8 January 2026"
        self.current_date_display = self.current_date_long  # For display: "8 January 2026"
        self.current_date_iso = now.strftime('%Y-%m-%d')  # For internal use
        
        # Initialize handlers
        self.email_handler = EmailHandler(self.config['email'])
        self.excel_processor = ExcelProcessor(self.config)
        
        self.logger.info("Large Deal Report Automation initialized")
    
    def cleanup_temp_files(self) -> None:
        """Clean up temporary files in the temp directory."""
        self.logger.info("Cleaning up temporary files...")
        cleaned_count = 0
        
        try:
            for file in self.temp_dir.glob("*.xlsx"):
                try:
                    file.unlink()
                    cleaned_count += 1
                except Exception as e:
                    self.logger.warning(f"Could not delete {file}: {str(e)}")
            
            if cleaned_count > 0:
                self.logger.info(f"Cleaned up {cleaned_count} temporary file(s)")
                
        except Exception as e:
            self.logger.warning(f"Error during cleanup: {str(e)}")
    
    def run(self) -> None:
        """
        Execute the complete automation process.
        
        Workflow:
        1. Download the previous day's report from email
        2. Download the daily worksheet from "large trade td" email
        3. Paste data into 'large deal report' sheet starting at A1
        4. Refresh the 'summary' sheet (table refresh)
        5. Copy summary table to 'iphone compatible' (paste values only)
        6. Save as "large deal report - {date}.xlsx"
        7. Open email preview in Outlook
        """
        workbook = None
        new_report_path = None
        
        try:
            self.logger.info("=" * 60)
            self.logger.info("Large Deal Report Automation - Starting")
            self.logger.info(f"Date: {self.current_date_display}")
            self.logger.info("=" * 60)
            
            # Step 1: Download previous day's report from email
            self.logger.info("\nStep 1: Downloading previous day's report from email...")
            previous_report_path = self.email_handler.download_previous_report(self.temp_dir)
            
            # Step 2: Download daily email attachment (from "large trade td" email)
            self.logger.info("\nStep 2: Downloading daily worksheet from 'Large Trade TD' email...")
            daily_sheet_path = self.email_handler.download_daily_attachment(
                self.temp_dir, 
                self.current_date_iso
            )
            
            # Step 3: Load the previous report and update 'large deal report' sheet
            # Use keep_macros=True to preserve macro buttons
            self.logger.info("\nStep 3: Updating 'large deal report' worksheet with today's data...")
            workbook = self.excel_processor.load_workbook(previous_report_path, keep_macros=True)
            self.excel_processor.update_large_deal_report_sheet(workbook, daily_sheet_path)
            
            # Step 4: Save intermediate version before xlwings refresh
            # Keep same extension as original to preserve macros
            original_ext = previous_report_path.suffix.lower()
            if original_ext == '.xlsm':
                temp_report_path = self.temp_dir / f"temp_report_{self.current_date_iso}.xlsm"
            else:
                temp_report_path = self.temp_dir / f"temp_report_{self.current_date_iso}.xlsx"
            self.excel_processor.save_workbook(workbook, temp_report_path)
            workbook.close()
            
            # Step 5: Refresh 'summary' sheet using xlwings (right-click refresh)
            self.logger.info("\nStep 4: Refreshing 'summary' worksheet (table refresh)...")
            workbook = self.excel_processor.load_workbook(temp_report_path)
            self.excel_processor.refresh_summary_sheet(workbook, temp_report_path)
            
            # Reload workbook after xlwings refresh
            workbook.close()
            workbook = self.excel_processor.load_workbook(temp_report_path, keep_macros=True)
            
            # Step 6: Copy summary to iphone compatible (paste values only)
            self.logger.info("\nStep 5: Copying 'summary' to 'iphone compatible' (values only)...")
            self.excel_processor.copy_summary_to_iphone_compatible(workbook)
            
            # Step 7: Save final report with date-stamped filename
            # Use .xlsm extension if original had macros, otherwise .xlsx
            self.logger.info("\nStep 6: Saving final report...")
            original_ext = previous_report_path.suffix.lower()
            if original_ext == '.xlsm':
                filename = f"Large Deal Report - {self.current_date_file}.xlsm"
            else:
                filename = f"Large Deal Report - {self.current_date_file}.xlsx"
            new_report_path = self.reports_dir / filename
            self.excel_processor.save_workbook(workbook, new_report_path)
            
            # Close workbook before sending email
            if workbook:
                workbook.close()
                workbook = None
            
            # Step 8: Prepare and preview email in Outlook
            self.logger.info("\nStep 7: Preparing email in Outlook...")
            sender_name = self.config.get('sender_name', 'Your Name')
            email_body = self.config.get('email_body', None)
            
            self.email_handler.send_report_email(
                new_report_path,
                self.config['distribution_list'],
                sender_name,
                self.current_date_display,
                email_body
            )
            
            # Step 9: Cleanup
            self.logger.info("\nStep 8: Cleaning up temporary files...")
            self.cleanup_temp_files()
            
            self.logger.info("\n" + "=" * 60)
            self.logger.info("Large Deal Report Automation - Completed!")
            self.logger.info(f"Report saved to: {new_report_path}")
            if self.email_handler.preview_before_send:
                self.logger.info("Email is open in Outlook - review and click Send.")
            self.logger.info("=" * 60)
            
        except FileNotFoundError as e:
            self.logger.error(f"File not found: {str(e)}")
            print(f"\nError: {str(e)}")
            sys.exit(1)
        except ValueError as e:
            self.logger.error(f"Error: {str(e)}")
            print(f"\nError: {str(e)}")
            sys.exit(1)
        except Exception as e:
            self.logger.error(f"Error: {str(e)}")
            print(f"\nError: {str(e)}")
            sys.exit(1)
        finally:
            try:
                if workbook:
                    workbook.close()
                self.cleanup_temp_files()
            except:
                pass


def main():
    """Main entry point."""
    try:
        automation = LargeDealReportAutomation()
        automation.run()
    except KeyboardInterrupt:
        print("\n\nCancelled by user.")
        sys.exit(1)
    except SystemExit:
        raise
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
