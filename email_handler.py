#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval and sending using Outlook on Windows.
"""

import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import List

# Import win32com for Outlook COM interface (Windows only)
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False


class EmailHandler:
    """Handles email retrieval and sending using Outlook."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - incoming_subject: Subject to search for daily data email
                - previous_report_subject: Subject pattern for previous report
                - previous_report_folder: Outlook folder containing previous reports
                - preview_before_send: If True, displays email for review before sending
        """
        self.email_config = email_config
        self.incoming_subject_template = email_config.get('incoming_subject', 'large trade td')
        self.previous_report_subject_template = email_config.get('previous_report_subject', 'large deal report')
        self.previous_report_folder = email_config.get('previous_report_folder', 'Large Deal Reports')
        self.preview_before_send = email_config.get('preview_before_send', True)
        self.logger = logging.getLogger(__name__)
        
        # Check Outlook is available
        if not OUTLOOK_AVAILABLE:
            raise ImportError(
                "Outlook COM not available.\n"
                "Install it with: pip install pywin32\n"
                "Note: This requires Windows and Outlook to be installed."
            )
        
        self.logger.info("Email handler initialized - using Outlook")
    
    def get_incoming_subject(self) -> str:
        """
        Get the incoming email subject with date placeholders replaced.
        
        Supported placeholders:
            {date_long}      - Long date (e.g., 07 January 2026)
            {date_long_day}  - Long date with day name (e.g., Tuesday, 07 January 2026)
            {date}           - DD/MM/YYYY (e.g., 07/01/2026)
            {date_dash}      - DD-MM-YYYY (e.g., 07-01-2026)
            {date_dot}       - DD.MM.YYYY (e.g., 07.01.2026)
            {day_name}       - Day name (e.g., Tuesday)
            {month_name}     - Month name (e.g., January)
            {dd}             - Day (e.g., 07)
            {d}              - Day without leading zero (e.g., 7)
            {mm}             - Month (e.g., 01)
            {yyyy}           - Year (e.g., 2026)
            {yy}             - Short year (e.g., 26)
        
        Returns:
            The subject string with placeholders replaced
        """
        now = datetime.now()
        
        subject = self.incoming_subject_template
        
        # Long date formats (no leading zero on day)
        # {date_long} = "7 January 2026" (no leading zero)
        subject = subject.replace('{date_long}', f"{now.day} {now.strftime('%B %Y')}")
        subject = subject.replace('{date_long_day}', f"{now.strftime('%A')}, {now.day} {now.strftime('%B %Y')}")
        
        # Short date formats
        subject = subject.replace('{date}', now.strftime('%d/%m/%Y'))
        subject = subject.replace('{date_dash}', now.strftime('%d-%m-%Y'))
        subject = subject.replace('{date_dot}', now.strftime('%d.%m.%Y'))
        
        # Individual components
        subject = subject.replace('{day_name}', now.strftime('%A'))
        subject = subject.replace('{month_name}', now.strftime('%B'))
        subject = subject.replace('{dd}', now.strftime('%d'))
        subject = subject.replace('{d}', str(now.day))
        subject = subject.replace('{mm}', now.strftime('%m'))
        subject = subject.replace('{yyyy}', now.strftime('%Y'))
        subject = subject.replace('{yy}', now.strftime('%y'))
        
        return subject
    
    def get_previous_working_day(self) -> datetime:
        """
        Get the previous working day (Monday-Friday).
        If today is Monday, returns Friday.
        If today is Sunday, returns Friday.
        If today is Saturday, returns Friday.
        Otherwise returns yesterday.
        
        Note: Does not account for bank holidays.
        
        Returns:
            datetime object for the previous working day
        """
        today = datetime.now()
        
        # weekday() returns 0=Monday, 1=Tuesday, ..., 6=Sunday
        if today.weekday() == 0:  # Monday -> Friday
            days_back = 3
        elif today.weekday() == 6:  # Sunday -> Friday
            days_back = 2
        elif today.weekday() == 5:  # Saturday -> Friday
            days_back = 1
        else:  # Tuesday-Friday -> previous day
            days_back = 1
        
        return today - timedelta(days=days_back)
    
    def get_previous_report_subject(self) -> str:
        """
        Get the previous report email subject with date placeholders replaced
        using the previous working day's date.
        
        Returns:
            The subject string with placeholders replaced for previous working day
        """
        prev_day = self.get_previous_working_day()
        
        subject = self.previous_report_subject_template
        
        # Month and year only
        # {month_year} = "January 2026"
        subject = subject.replace('{month_year}', prev_day.strftime('%B %Y'))
        
        # Long date formats (no leading zero on day)
        # {date_long} = "7 January 2026" (no leading zero)
        subject = subject.replace('{date_long}', f"{prev_day.day} {prev_day.strftime('%B %Y')}")
        subject = subject.replace('{date_long_day}', f"{prev_day.strftime('%A')}, {prev_day.day} {prev_day.strftime('%B %Y')}")
        
        # Short date formats
        subject = subject.replace('{date}', prev_day.strftime('%d/%m/%Y'))
        subject = subject.replace('{date_dash}', prev_day.strftime('%d-%m-%Y'))
        subject = subject.replace('{date_dot}', prev_day.strftime('%d.%m.%Y'))
        
        # Individual components
        subject = subject.replace('{day_name}', prev_day.strftime('%A'))
        subject = subject.replace('{month_name}', prev_day.strftime('%B'))
        subject = subject.replace('{dd}', prev_day.strftime('%d'))
        subject = subject.replace('{d}', str(prev_day.day))
        subject = subject.replace('{mm}', prev_day.strftime('%m'))
        subject = subject.replace('{yyyy}', prev_day.strftime('%Y'))
        subject = subject.replace('{yy}', prev_day.strftime('%y'))
        
        return subject
    
    def download_daily_attachment(self, save_directory: Path, current_date: str) -> Path:
        """
        Download the daily worksheet attachment from Outlook.
        Searches for emails with subject containing the configured pattern.
        
        Args:
            save_directory: Directory to save the downloaded attachment
            current_date: Current date string (YYYY-MM-DD)
            
        Returns:
            Path to the downloaded file
            
        Raises:
            Exception: If email retrieval or download fails
        """
        self.logger.info("Connecting to Outlook...")
        
        # Get the subject with date placeholders replaced
        search_subject = self.get_incoming_subject()
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            
            self.logger.info(f"Searching for emails with subject containing '{search_subject}'...")
            
            # Get all inbox items and filter by subject
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # Sort descending (most recent first)
            
            # Search for matching email
            found_email = None
            today = datetime.now().date()
            
            # First try to find one from today
            for message in messages:
                try:
                    subject = message.Subject.lower() if message.Subject else ""
                    received_date = message.ReceivedTime.date()
                    
                    if search_subject.lower() in subject:
                        if received_date == today:
                            found_email = message
                            self.logger.info(f"Found today's email: {message.Subject}")
                            break
                except Exception:
                    continue
            
            # If no email from today, get the most recent matching email
            if not found_email:
                self.logger.warning("No matching email found from today. Searching for most recent...")
                for message in messages:
                    try:
                        subject = message.Subject.lower() if message.Subject else ""
                        if search_subject.lower() in subject:
                            found_email = message
                            self.logger.info(f"Found recent email: {message.Subject}")
                            break
                    except Exception:
                        continue
            
            if not found_email:
                raise Exception(
                    f"No emails found with subject containing '{search_subject}'. "
                    "Please ensure you have received the daily email."
                )
            
            # Find and download Excel attachment
            attachment_path = None
            attachments = found_email.Attachments
            
            # Make sure save directory exists
            save_directory.mkdir(parents=True, exist_ok=True)
            
            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                filename = attachment.FileName
                
                # Flexible check for Excel files
                filename_lower = filename.lower().strip() if filename else ""
                is_excel = '.xlsx' in filename_lower or '.xls' in filename_lower
                
                if is_excel:
                    attachment_path = save_directory / f"daily_sheet_{current_date}.xlsx"
                    # Convert to absolute path for Windows
                    absolute_path = str(attachment_path.absolute())
                    attachment.SaveAsFile(absolute_path)
                    self.logger.info(f"Downloaded attachment: {filename}")
                    break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    "No Excel attachment (.xlsx or .xls) found in email. "
                    "Please ensure the email contains an Excel file attachment."
                )
            
            return attachment_path
            
        except Exception as e:
            self.logger.error(f"Error downloading via Outlook: {str(e)}")
            raise Exception(f"Error downloading email attachment via Outlook: {str(e)}")
    
    def download_previous_report(self, save_directory: Path) -> Path:
        """
        Download the previous working day's report from Outlook.
        Searches in the configured folder (default: "Large Deal Reports").
        
        Args:
            save_directory: Directory to save the downloaded report
            
        Returns:
            Path to the downloaded report file
            
        Raises:
            Exception: If report cannot be found or downloaded
        """
        self.logger.info("Searching for previous day's report in Outlook...")
        
        search_subject = self.get_previous_report_subject()
        prev_day = self.get_previous_working_day()
        
        self.logger.info(f"Today is: {datetime.now().strftime('%A, %d %B %Y')}")
        self.logger.info(f"Previous working day: {prev_day.strftime('%A, %d %B %Y')}")
        self.logger.info(f"Looking for subject containing: '{search_subject}'")
        self.logger.info(f"Searching in folder: '{self.previous_report_folder}'")
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            found_email = None
            
            # First try to find the specified subfolder under Inbox
            try:
                inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
                
                # List all available folders for debugging
                self.logger.info("Available folders in Inbox:")
                for folder in inbox.Folders:
                    self.logger.info(f"  - '{folder.Name}'")
                
                # Try to find the subfolder
                target_folder = None
                
                if self.previous_report_folder:
                    # Search for the subfolder (case-insensitive)
                    for folder in inbox.Folders:
                        if folder.Name.lower() == self.previous_report_folder.lower():
                            target_folder = folder
                            self.logger.info(f"Found matching folder: '{folder.Name}'")
                            break
                    
                    if not target_folder:
                        self.logger.warning(
                            f"Folder '{self.previous_report_folder}' not found under Inbox. "
                            "Searching in Inbox instead."
                        )
                        target_folder = inbox
                else:
                    target_folder = inbox
                
                # Search in the target folder
                self.logger.info(f"Searching in '{target_folder.Name}'...")
                messages = target_folder.Items
                messages.Sort("[ReceivedTime]", True)  # Most recent first
                
                # Show recent emails for debugging
                self.logger.info("Recent emails in this folder:")
                count = 0
                for message in messages:
                    try:
                        if count < 10:  # Show first 10
                            subj = message.Subject if message.Subject else "(no subject)"
                            self.logger.info(f"  - '{subj}'")
                            count += 1
                        else:
                            break
                    except:
                        continue
                
                # Reset and search for matching email
                messages = target_folder.Items
                messages.Sort("[ReceivedTime]", True)
                
                self.logger.info(f"\nSearching for subject containing: '{search_subject}'")
                
                for message in messages:
                    try:
                        subject = message.Subject if message.Subject else ""
                        
                        # Check if subject contains our search term (case-insensitive)
                        if search_subject.lower() in subject.lower():
                            self.logger.info(f"Subject MATCH: '{subject}'")
                            
                            # Check if it has an Excel attachment
                            if message.Attachments.Count > 0:
                                self.logger.info(f"  Attachments found: {message.Attachments.Count}")
                                for i in range(1, message.Attachments.Count + 1):
                                    attachment = message.Attachments.Item(i)
                                    filename = attachment.FileName
                                    self.logger.info(f"    - '{filename}'")
                                    
                                    # Check for Excel file (more flexible check)
                                    filename_lower = filename.lower()
                                    is_excel = (
                                        filename_lower.endswith('.xlsx') or 
                                        filename_lower.endswith('.xls') or
                                        '.xlsx' in filename_lower or
                                        '.xls' in filename_lower
                                    )
                                    
                                    if is_excel:
                                        found_email = message
                                        self.logger.info(f"  Excel file FOUND - using this email")
                                        break
                                
                                if not found_email:
                                    self.logger.info(f"  No Excel attachment found in this email")
                            else:
                                self.logger.info(f"  No attachments in this email")
                            
                            if found_email:
                                break
                    except Exception as e:
                        self.logger.warning(f"Error checking email: {str(e)}")
                        continue
                        
            except Exception as e:
                self.logger.error(f"Error searching folder: {str(e)}")
            
            if not found_email:
                self.logger.error("No matching email with Excel attachment was found!")
                raise Exception(
                    f"\nCould not find previous report.\n"
                    f"Searched for: '{search_subject}'\n"
                    f"In folder: '{self.previous_report_folder}'\n"
                    f"Please check:\n"
                    f"  1. The folder name matches exactly\n"
                    f"  2. The email subject matches the format above\n"
                    f"  3. The email has an Excel attachment (.xlsx or .xls)"
                )
            else:
                self.logger.info(f"Successfully found email to download")
            
            # Download the Excel attachment
            attachment_path = None
            attachments = found_email.Attachments
            
            # Make sure save directory exists
            save_directory.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"Downloading attachment from email...")
            self.logger.info(f"Save folder: {save_directory.absolute()}")
            self.logger.info(f"Number of attachments: {attachments.Count}")
            
            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                filename = attachment.FileName
                self.logger.info(f"Checking attachment: '{filename}'")
                
                # Flexible check for Excel files
                filename_lower = filename.lower().strip()
                is_excel = (
                    '.xlsx' in filename_lower or 
                    '.xls' in filename_lower or
                    filename_lower.endswith('.xlsx') or
                    filename_lower.endswith('.xls')
                )
                
                if is_excel:
                    attachment_path = save_directory / "previous_report.xlsx"
                    # Convert to absolute path for Windows
                    absolute_path = str(attachment_path.absolute())
                    self.logger.info(f"Saving to: {absolute_path}")
                    attachment.SaveAsFile(absolute_path)
                    self.logger.info(f"Downloaded: '{filename}'")
                    break
                else:
                    self.logger.info(f"  Not an Excel file, skipping")
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    f"No Excel attachment found in the previous report email.\n"
                    f"Attachments in email: {attachments.Count}"
                )
            
            return attachment_path
            
        except Exception as e:
            self.logger.error(f"Error downloading previous report: {str(e)}")
            raise
    
    def send_report_email(
        self, 
        report_path: Path, 
        distribution_list: List[str], 
        sender_name: str, 
        current_date: str,
        email_body: str = None
    ) -> None:
        """
        Send or preview the report via Outlook.
        
        If preview_before_send is True (default), the email will be displayed
        in Outlook for review before sending manually.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (DD/MM/YYYY format for display)
            email_body: Optional custom email body text
            
        Raises:
            Exception: If email sending fails
        """
        action = "preview" if self.preview_before_send else "send"
        self.logger.info(f"Preparing email in Outlook for {action}...")
        
        try:
            # Validate report file exists
            if not report_path.exists():
                raise FileNotFoundError(f"Report file not found: {report_path}")
            
            # Connect to Outlook
            self.logger.debug("Connecting to Outlook application...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = MailItem
            
            # Set email properties with proper date format
            mail.Subject = f"Large Deal Report - {current_date}"
            
            # Use custom body if provided, otherwise use default
            if email_body:
                mail.Body = email_body
            else:
                mail.Body = (
                    f"Hi all,\n\n"
                    f"Please find attached today's Large Deal Report for {current_date}.\n\n"
                    f"Kind regards,\n{sender_name}"
                )
            
            # Add recipients (could be distribution lists or individual emails)
            self.logger.info(f"Adding {len(distribution_list)} recipient(s)...")
            for recipient in distribution_list:
                mail.Recipients.Add(recipient)
            
            # Attach report
            self.logger.info(f"Attaching report: {report_path.name}")
            mail.Attachments.Add(str(report_path))
            
            if self.preview_before_send:
                # Display the email for review - user must click Send manually
                self.logger.info("Opening email for preview - please review and click Send when ready...")
                mail.Display(True)  # True = modal window
                self.logger.info("Email displayed for preview. User will send manually.")
            else:
                # Send email automatically
                self.logger.info("Sending email via Outlook...")
                mail.Send()
                self.logger.info("Email sent successfully via Outlook!")
            
        except FileNotFoundError as e:
            self.logger.error(f"Report file not found: {str(e)}")
            raise
        except Exception as e:
            error_msg = f"Error {action}ing email via Outlook: {str(e)}"
            self.logger.error(error_msg)
            if "Outlook.Application" in str(e) or "COM" in str(e) or "win32com" in str(e):
                raise Exception(
                    f"{error_msg}\n"
                    "Make sure Outlook is installed and you're logged into your account.\n"
                    "Also ensure pywin32 is installed: pip install pywin32"
                )
            raise Exception(error_msg)
