#!/usr/bin/env python3
"""
Large Deal Report Automation Script
Automates the daily generation and distribution of the Large Deal Report.

Workflow:
1. Download the daily worksheet from email with subject "large trade td"
2. Copy/paste that data into the 'large deal report' worksheet (starting at A1)
3. Refresh the 'summary' worksheet (right-click refresh on table)
4. Copy the summary table and paste values only into 'iphone compatible'
5. Save the workbook as "large deal report - {today's date}.xlsx"
6. Open email preview (or send) to distribution list
"""

import os
import sys
import json
import logging
import glob
from datetime import datetime
from pathlib import Path
from typing import Optional

from email_handler import EmailHandler
from excel_processor import ExcelProcessor


class LargeDealReportAutomation:
    """Main automation class that orchestrates the report generation workflow."""
    
    def __init__(self, config_path: str = None):
        """
        Initialize the automation with configuration.
        
        Args:
            config_path: Path to the configuration JSON file
        """
        # Setup basic logging first (before loading config)
        logging.basicConfig(
            level=logging.INFO,
            format='%(levelname)s: %(message)s',
            handlers=[logging.StreamHandler(sys.stdout)]
        )
        self.logger = logging.getLogger(__name__)
        
        # Find config file - try config.txt first, then config.json
        if config_path is None:
            if os.path.exists('config.txt'):
                config_path = 'config.txt'
            elif os.path.exists('config.json'):
                config_path = 'config.json'
            else:
                config_path = 'config.txt'  # Default for error message
        
        # Load configuration
        self.config = self.load_config(config_path)
        # Setup logging with config (after config is loaded)
        self.setup_logging()
        
        # Setup directories
        self.reports_dir = Path(self.config.get('reports_directory', './reports'))
        self.reports_dir.mkdir(exist_ok=True)
        self.temp_dir = Path(self.config.get('temp_directory', './temp'))
        self.temp_dir.mkdir(exist_ok=True)
        
        # Date formats
        self.current_date_file = datetime.now().strftime('%d-%m-%Y')  # For filenames (DD-MM-YYYY)
        self.current_date_display = datetime.now().strftime('%d/%m/%Y')  # For display (DD/MM/YYYY)
        self.current_date_iso = datetime.now().strftime('%Y-%m-%d')  # For internal use
        
        # Initialize handlers
        self.email_handler = EmailHandler(self.config['email'])
        self.excel_processor = ExcelProcessor(self.config)
        
        self.logger.info("Large Deal Report Automation initialized")
    
    def setup_logging(self) -> None:
        """Configure logging for the application."""
        log_level = self.config.get('log_level', 'INFO').upper()
        log_file = self.config.get('log_file', None)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Setup root logger
        root_logger = logging.getLogger()
        root_logger.setLevel(getattr(logging, log_level, logging.INFO))
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
        
        # File handler (if specified)
        if log_file:
            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(formatter)
            root_logger.addHandler(file_handler)
    
    def load_config(self, config_path: str) -> dict:
        """
        Load configuration from JSON file.
        
        Args:
            config_path: Path to the configuration file
            
        Returns:
            Configuration dictionary
            
        Raises:
            SystemExit: If configuration file is missing or invalid
        """
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
            
            # Validate required configuration sections
            required_sections = ['email', 'distribution_list']
            for section in required_sections:
                if section not in config:
                    raise ValueError(f"Missing required configuration section: {section}")
            
            # Validate email configuration - need incoming_subject for searching emails
            if 'incoming_subject' not in config['email']:
                config['email']['incoming_subject'] = 'large trade td'  # Default
            
            # Validate SMTP credentials if not using Outlook
            use_outlook = config['email'].get('use_outlook', True)  # Default to Outlook on Windows
            if not use_outlook:
                smtp_required_fields = ['username', 'password']
                for field in smtp_required_fields:
                    if field not in config['email']:
                        raise ValueError(
                            f"Missing required email configuration: {field}\n"
                            "Either provide email credentials or set 'use_outlook' to true."
                        )
            
            return config
            
        except FileNotFoundError:
            print(f"Error: Configuration file not found.")
            print("Please create a config.txt file in the same folder as this script.")
            print("See README.md for setup instructions.")
            sys.exit(1)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid JSON in '{config_path}': {str(e)}")
            sys.exit(1)
        except ValueError as e:
            print(f"Error: Configuration validation failed: {str(e)}")
            sys.exit(1)
        except Exception as e:
            print(f"Error loading configuration: {str(e)}")
            sys.exit(1)
    
    def find_latest_report(self) -> Path:
        """
        Find the most recently generated report file.
        Looks for files matching "large deal report*.xlsx" pattern.
        
        Returns:
            Path to the latest report file
            
        Raises:
            FileNotFoundError: If no reports are found
        """
        # Try multiple patterns to find reports
        patterns = [
            str(self.reports_dir / "large deal report -*.xlsx"),
            str(self.reports_dir / "large deal report - *.xlsx"),
            str(self.reports_dir / "large deal report*.xlsx"),
            str(self.reports_dir / "Large Deal Report*.xlsx"),
        ]
        
        reports = []
        for pattern in patterns:
            reports.extend(glob.glob(pattern))
        
        # Remove duplicates
        reports = list(set(reports))
        
        if not reports:
            error_msg = (
                f"No existing reports found in {self.reports_dir}. "
                "Please ensure at least one report exists with the naming pattern: "
                "'large deal report - DD-MM-YYYY.xlsx'"
            )
            self.logger.error(error_msg)
            raise FileNotFoundError(error_msg)
        
        # Sort by modification time and get the latest
        latest_report = max(reports, key=os.path.getmtime)
        latest_report_path = Path(latest_report)
        
        self.logger.info(f"Found latest report: {latest_report_path}")
        return latest_report_path
    
    def cleanup_temp_files(self) -> None:
        """Clean up temporary files in the temp directory."""
        self.logger.info("Cleaning up temporary files...")
        cleaned_count = 0
        
        try:
            for file in self.temp_dir.glob("daily_sheet_*.xlsx"):
                try:
                    file.unlink()
                    cleaned_count += 1
                    self.logger.debug(f"Deleted temporary file: {file}")
                except Exception as e:
                    self.logger.warning(f"Could not delete {file}: {str(e)}")
            
            if cleaned_count > 0:
                self.logger.info(f"Cleaned up {cleaned_count} temporary file(s)")
            else:
                self.logger.debug("No temporary files to clean up")
                
        except Exception as e:
            self.logger.warning(f"Error during cleanup: {str(e)}")
    
    def run(self) -> None:
        """
        Execute the complete automation process.
        
        Workflow:
        1. Download the previous day's report from email (sent items/inbox)
        2. Download the daily worksheet from "large trade td" email
        3. Paste data into 'large deal report' sheet starting at A1
        4. Refresh the 'summary' sheet (table refresh)
        5. Copy summary table to 'iphone compatible' (paste values only)
        6. Save as "large deal report - {date}.xlsx"
        7. Open email preview (or send) to distribution list
        
        Raises:
            SystemExit: If the automation fails
        """
        workbook = None
        daily_sheet_path = None
        previous_report_path = None
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
            
            # Step 3: Load the previous report and update 'large deal report' sheet with daily data
            self.logger.info("\nStep 3: Updating 'large deal report' worksheet with today's data...")
            workbook = self.excel_processor.load_workbook(previous_report_path)
            self.excel_processor.update_large_deal_report_sheet(workbook, daily_sheet_path)
            
            # Step 4: Save intermediate version before xlwings refresh
            # (xlwings needs to work with a saved file)
            temp_report_path = self.temp_dir / f"temp_report_{self.current_date_iso}.xlsx"
            self.excel_processor.save_workbook(workbook, temp_report_path)
            workbook.close()
            
            # Step 5: Refresh 'summary' sheet using xlwings (right-click refresh)
            self.logger.info("\nStep 4: Refreshing 'summary' worksheet (table refresh)...")
            workbook = self.excel_processor.load_workbook(temp_report_path)
            self.excel_processor.refresh_summary_sheet(workbook, temp_report_path)
            
            # Reload workbook after xlwings refresh
            workbook.close()
            workbook = self.excel_processor.load_workbook(temp_report_path)
            
            # Step 6: Copy summary to iphone compatible (paste values only)
            self.logger.info("\nStep 5: Copying 'summary' to 'iphone compatible' (values only)...")
            self.excel_processor.copy_summary_to_iphone_compatible(workbook)
            
            # Step 7: Save final report with date-stamped filename
            self.logger.info("\nStep 6: Saving final report...")
            filename = f"large deal report - {self.current_date_file}.xlsx"
            new_report_path = self.reports_dir / filename
            self.excel_processor.save_workbook(workbook, new_report_path)
            
            # Close workbook before sending email
            if workbook:
                workbook.close()
                workbook = None
            
            # Step 8: Prepare and preview/send email
            self.logger.info("\nStep 7: Preparing email...")
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
            self.logger.info("Large Deal Report Automation - Completed Successfully!")
            self.logger.info(f"Report saved to: {new_report_path}")
            if self.email_handler.preview_before_send:
                self.logger.info("Email is open for preview - please review and click Send.")
            self.logger.info("=" * 60)
            
        except FileNotFoundError as e:
            self.logger.error(f"File not found error: {str(e)}")
            print(f"\nError: {str(e)}")
            print("Please check that all required files exist.")
            sys.exit(1)
        except ValueError as e:
            self.logger.error(f"Value error: {str(e)}")
            print(f"\nError: {str(e)}")
            print("Please check your configuration and report structure.")
            sys.exit(1)
        except Exception as e:
            self.logger.error(f"Unexpected error during automation: {str(e)}", exc_info=True)
            print(f"\nError during automation: {str(e)}")
            print("Please check the error message and try again.")
            print("See the logs for more details.")
            sys.exit(1)
        finally:
            # Ensure cleanup happens even on error
            try:
                if workbook:
                    workbook.close()
                self.cleanup_temp_files()
            except Exception as e:
                self.logger.warning(f"Error during final cleanup: {str(e)}")


def main():
    """Main entry point."""
    try:
        automation = LargeDealReportAutomation()
        automation.run()
    except KeyboardInterrupt:
        print("\n\nAutomation interrupted by user.")
        sys.exit(1)
    except SystemExit:
        raise
    except Exception as e:
        print(f"\nFatal error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
