#!/usr/bin/env python3
"""
Large Deal Report Automation Script
Automates the daily generation and distribution of the Large Deal Report.
"""

import os
import sys
import json
import logging
from datetime import datetime
from pathlib import Path
import time

from email_handler import EmailHandler
from excel_processor import ExcelProcessor

class LargeDealReportAutomation:
    """Main automation class that orchestrates the report generation workflow."""
    
    def __init__(self, config_path: str = 'config.json'):
        """Initialize the automation with configuration."""
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s: %(message)s',
            handlers=[logging.StreamHandler(sys.stdout)]
        )
        self.logger = logging.getLogger(__name__)
        
        # Load config
        self.config = self.load_config(config_path)
        
        # Directories
        self.reports_dir = Path(self.config.get('reports_directory', './reports'))
        self.reports_dir.mkdir(exist_ok=True)
        self.temp_dir = Path(self.config.get('temp_directory', './temp'))
        self.temp_dir.mkdir(exist_ok=True)
        
        self.current_date = datetime.now().strftime('%d-%b-%Y') # e.g., 07-Jan-2026
        
        # Initialize handlers
        self.email_handler = EmailHandler(self.config['email'])
        self.excel_processor = ExcelProcessor(self.config)
        
        self.logger.info("Large Deal Report Automation initialized")
    
    def load_config(self, config_path: str) -> dict:
        """Load configuration from JSON file."""
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            self.logger.error(f"Error loading config: {e}")
            sys.exit(1)

    def run(self) -> None:
        """Execute the automation process."""
        try:
            self.logger.info("=" * 60)
            self.logger.info(f"Large Deal Report Automation - {self.current_date}")
            self.logger.info("=" * 60)
            
            # 1. Download Daily Attachment
            self.logger.info("Step 1: Downloading daily input file from email...")
            daily_sheet_path = self.email_handler.download_daily_attachment(
                self.temp_dir, 
                self.current_date
            )
            
            # 2. Identify Master Workbook
            # We assume the user has configured a path to a clean Master/Template workbook
            # OR we could pick the most recent one. 
            # For robustness, we'll use the configured path.
            master_path_str = self.config['excel'].get('master_workbook_path')
            if not master_path_str:
                self.logger.error("No 'master_workbook_path' defined in config.json")
                sys.exit(1)
                
            master_path = Path(master_path_str)
            if not master_path.exists():
                self.logger.error(f"Master workbook not found at: {master_path}")
                self.logger.error("Please ensure the file exists or update config.json")
                sys.exit(1)

            # 3. Process Excel
            self.logger.info("Step 2: Processing Excel Report...")
            
            # Define output path
            output_filename = f"large deal report- {self.current_date}.xlsx"
            output_path = self.reports_dir / output_filename
            
            # Determine absolute paths for Excel COM (it often requires absolute paths)
            master_abs = master_path.resolve()
            daily_abs = daily_sheet_path.resolve()
            output_abs = output_path.resolve()
            
            self.excel_processor.process_report(master_abs, daily_abs, output_abs)
            
            # 4. Create Draft Email
            self.logger.info("Step 3: Creating Email Draft...")
            self.email_handler.send_report_email(
                output_abs,
                self.config['distribution_list'],
                self.config.get('sender_name', 'Analyst'),
                self.current_date
            )
            
            self.logger.info("=" * 60)
            self.logger.info("Automation Complete. Please check the Outlook draft.")
            self.logger.info("=" * 60)

        except Exception as e:
            self.logger.error(f"Automation failed: {e}", exc_info=True)
            input("Press Enter to exit...") # Keep window open on error
            sys.exit(1)

if __name__ == "__main__":
    automation = LargeDealReportAutomation()
    automation.run()
