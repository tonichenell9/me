#!/usr/bin/env python3
"""
Large Deal Report Automation Script
Automates the daily generation and distribution of the Large Deal Report.
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
    
    def __init__(self, config_path: str = 'config.json'):
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
        
        # Load configuration
        self.config = self.load_config(config_path)
        # Setup logging with config (after config is loaded)
        self.setup_logging()
        
        # Setup directories
        self.reports_dir = Path(self.config.get('reports_directory', './reports'))
        self.reports_dir.mkdir(exist_ok=True)
        self.temp_dir = Path(self.config.get('temp_directory', './temp'))
        self.temp_dir.mkdir(exist_ok=True)
        
        self.current_date = datetime.now().strftime('%Y-%m-%d')
        
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
            
            # Validate email configuration
            required_email_fields = ['sender_email']
            for field in required_email_fields:
                if field not in config['email']:
                    raise ValueError(f"Missing required email configuration: {field}")
            
            # Validate SMTP credentials if not using Outlook for sending
            use_outlook = config['email'].get('use_outlook_for_sending', False)
            if not use_outlook:
                smtp_required_fields = ['username', 'password']
                for field in smtp_required_fields:
                    if field not in config['email']:
                        raise ValueError(
                            f"Missing required email configuration: {field}\n"
                            "Either provide SMTP credentials or set 'use_outlook_for_sending' to true."
                        )
            
            return config
            
        except FileNotFoundError:
            print(f"Error: Configuration file '{config_path}' not found.")
            print("Please create a config.json file from config.json.example")
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
        
        Returns:
            Path to the latest report file
            
        Raises:
            FileNotFoundError: If no reports are found
        """
        pattern = str(self.reports_dir / "large deal report -*.xlsx")
        reports = glob.glob(pattern)
        
        if not reports:
            error_msg = (
                f"No existing reports found in {self.reports_dir}. "
                "Please ensure at least one report exists with the naming pattern: "
                "'large deal report -YYYY-MM-DD.xlsx'"
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
        
        Raises:
            SystemExit: If the automation fails
        """
        workbook = None
        daily_sheet_path = None
        
        try:
            self.logger.info("=" * 60)
            self.logger.info("Large Deal Report Automation - Starting")
            self.logger.info("=" * 60)
            
            # Step 1: Find latest report
            self.logger.info("Step 1: Finding latest report...")
            latest_report_path = self.find_latest_report()
            workbook = self.excel_processor.load_workbook(latest_report_path)
            
            # Step 2: Download daily email attachment
            self.logger.info("Step 2: Downloading daily email attachment...")
            daily_sheet_path = self.email_handler.download_daily_attachment(
                self.temp_dir, 
                self.current_date
            )
            
            # Step 3: Update worksheet 1 with daily data
            self.logger.info("Step 3: Updating Worksheet 1...")
            self.excel_processor.update_worksheet_1(workbook, daily_sheet_path)
            
            # Step 4: Refresh worksheet 2
            self.logger.info("Step 4: Refreshing Worksheet 2...")
            self.excel_processor.refresh_worksheet_2(workbook, latest_report_path)
            
            # Step 5: Copy worksheet 2 to worksheet 3
            self.logger.info("Step 5: Copying Worksheet 2 to Worksheet 3...")
            self.excel_processor.copy_worksheet_2_to_3(workbook)
            
            # Step 6: Save report with date-stamped filename
            self.logger.info("Step 6: Saving report...")
            filename = f"large deal report -{self.current_date}.xlsx"
            new_report_path = self.reports_dir / filename
            self.excel_processor.save_workbook(workbook, new_report_path)
            
            # Close workbook before sending email
            if workbook:
                workbook.close()
                workbook = None
            
            # Step 7: Send email
            self.logger.info("Step 7: Sending email...")
            sender_name = self.config.get('sender_name', 'Automated Report System')
            self.email_handler.send_report_email(
                new_report_path,
                self.config['distribution_list'],
                sender_name,
                self.current_date
            )
            
            # Step 8: Cleanup
            self.logger.info("Step 8: Cleaning up temporary files...")
            self.cleanup_temp_files()
            
            self.logger.info("=" * 60)
            self.logger.info("Large Deal Report Automation - Completed Successfully!")
            self.logger.info(f"Report saved to: {new_report_path}")
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
