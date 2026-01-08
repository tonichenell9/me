#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP/Outlook) and sending (SMTP or Outlook COM) functionality.
"""

import imaplib
import email
import smtplib
import logging
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from pathlib import Path
from typing import Optional, List

# Try to import win32com for Outlook COM interface (Windows only)
try:
    import win32com.client
    OUTLOOK_COM_AVAILABLE = True
except ImportError:
    OUTLOOK_COM_AVAILABLE = False


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address (for IMAP/SMTP)
                - password: App password or password (for IMAP/SMTP)
                - incoming_subject: Subject line to search for incoming emails (e.g., "large trade td")
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address (optional if using Outlook)
                - smtp_port: SMTP server port (optional if using Outlook)
                - use_outlook: If True, use Outlook COM for both reading and sending (Windows only)
                - preview_before_send: If True, displays the email for review before sending
        """
        self.email_config = email_config
        self.username = email_config.get('username')
        self.password = email_config.get('password')
        self.incoming_subject_template = email_config.get('incoming_subject', 'large trade td')
        self.previous_report_subject_template = email_config.get('previous_report_subject', 'large deal report')
        self.imap_server = email_config.get('imap_server', 'outlook.office365.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.office365.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.use_outlook = email_config.get('use_outlook', True)
        self.preview_before_send = email_config.get('preview_before_send', True)
        self.logger = logging.getLogger(__name__)
        
        # Validate Outlook availability if requested
        if self.use_outlook:
            if not OUTLOOK_COM_AVAILABLE:
                raise ImportError(
                    "Outlook COM requested but win32com not available.\n"
                    "Install it with: pip install pywin32\n"
                    "Note: This requires Windows and Outlook to be installed."
                )
            self.logger.info("Will use Outlook COM interface for email operations")
        else:
            # Validate IMAP/SMTP credentials are provided
            if not self.username or not self.password:
                raise ValueError(
                    "Email credentials (username and password) are required "
                    "when not using Outlook."
                )
    
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
        
        # Long date formats
        subject = subject.replace('{date_long}', now.strftime('%d %B %Y'))  # 07 January 2026
        subject = subject.replace('{date_long_day}', now.strftime('%A, %d %B %Y'))  # Tuesday, 07 January 2026
        
        # Short date formats
        subject = subject.replace('{date}', now.strftime('%d/%m/%Y'))
        subject = subject.replace('{date_dash}', now.strftime('%d-%m-%Y'))
        subject = subject.replace('{date_dot}', now.strftime('%d.%m.%Y'))
        
        # Individual components
        subject = subject.replace('{day_name}', now.strftime('%A'))  # Tuesday
        subject = subject.replace('{month_name}', now.strftime('%B'))  # January
        subject = subject.replace('{dd}', now.strftime('%d'))
        subject = subject.replace('{d}', str(now.day))  # Day without leading zero
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
        
        from datetime import timedelta
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
        
        # Long date formats
        subject = subject.replace('{date_long}', prev_day.strftime('%d %B %Y'))
        subject = subject.replace('{date_long_day}', prev_day.strftime('%A, %d %B %Y'))
        
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
    
    def download_previous_report(self, save_directory: Path) -> Path:
        """
        Download the previous working day's report from sent items or inbox.
        This is the report we sent out yesterday that we'll update with today's data.
        
        Args:
            save_directory: Directory to save the downloaded report
            
        Returns:
            Path to the downloaded report file
            
        Raises:
            Exception: If report cannot be found or downloaded
        """
        if self.use_outlook:
            return self._download_previous_report_via_outlook(save_directory)
        else:
            return self._download_previous_report_via_imap(save_directory)
    
    def _download_previous_report_via_outlook(self, save_directory: Path) -> Path:
        """
        Download the previous report using Outlook COM interface.
        Searches Sent Items first, then Inbox.
        
        Args:
            save_directory: Directory to save the downloaded report
            
        Returns:
            Path to the downloaded report file
        """
        self.logger.info("Searching for previous day's report in Outlook...")
        
        search_subject = self.get_previous_report_subject()
        prev_day = self.get_previous_working_day()
        
        self.logger.info(f"Looking for report with subject containing '{search_subject}'")
        self.logger.info(f"Previous working day: {prev_day.strftime('%A, %d %B %Y')}")
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Search in Sent Items first (folder 5), then Inbox (folder 6)
            folders_to_search = [
                (5, "Sent Items"),
                (6, "Inbox")
            ]
            
            found_email = None
            
            for folder_id, folder_name in folders_to_search:
                self.logger.info(f"Searching in {folder_name}...")
                
                try:
                    folder = namespace.GetDefaultFolder(folder_id)
                    messages = folder.Items
                    messages.Sort("[ReceivedTime]", True)  # Most recent first
                    
                    for message in messages:
                        try:
                            subject = message.Subject if message.Subject else ""
                            
                            if search_subject.lower() in subject.lower():
                                # Check if it has an Excel attachment
                                if message.Attachments.Count > 0:
                                    for i in range(1, message.Attachments.Count + 1):
                                        attachment = message.Attachments.Item(i)
                                        if attachment.FileName.lower().endswith(('.xlsx', '.xls')):
                                            found_email = message
                                            self.logger.info(f"Found report in {folder_name}: {subject}")
                                            break
                                if found_email:
                                    break
                        except Exception:
                            continue
                    
                    if found_email:
                        break
                        
                except Exception as e:
                    self.logger.warning(f"Could not search {folder_name}: {str(e)}")
                    continue
            
            if not found_email:
                raise Exception(
                    f"Could not find previous report with subject containing '{search_subject}'. "
                    f"Searched in Sent Items and Inbox. "
                    f"Please ensure you sent/received the report on {prev_day.strftime('%A, %d %B %Y')}."
                )
            
            # Download the Excel attachment
            attachment_path = None
            attachments = found_email.Attachments
            
            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                filename = attachment.FileName
                
                if filename.lower().endswith(('.xlsx', '.xls')):
                    attachment_path = save_directory / f"previous_report.xlsx"
                    attachment.SaveAsFile(str(attachment_path))
                    self.logger.info(f"Downloaded previous report: {filename} -> {attachment_path}")
                    break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception("No Excel attachment found in the previous report email.")
            
            return attachment_path
            
        except Exception as e:
            self.logger.error(f"Error downloading previous report: {str(e)}")
            raise
    
    def _download_previous_report_via_imap(self, save_directory: Path) -> Path:
        """
        Download the previous report using IMAP.
        
        Args:
            save_directory: Directory to save the downloaded report
            
        Returns:
            Path to the downloaded report file
        """
        self.logger.info("Searching for previous day's report via IMAP...")
        
        search_subject = self.get_previous_report_subject()
        prev_day = self.get_previous_working_day()
        
        self.logger.info(f"Looking for report with subject containing '{search_subject}'")
        
        mail = None
        
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            
            # Try Sent folder first, then Inbox
            folders_to_try = ['"[Gmail]/Sent Mail"', '"Sent Items"', '"Sent"', 'INBOX']
            
            found_email_data = None
            
            for folder in folders_to_try:
                try:
                    status, _ = mail.select(folder)
                    if status != 'OK':
                        continue
                    
                    self.logger.info(f"Searching in {folder}...")
                    
                    search_criteria = f'(SUBJECT "{search_subject}")'
                    status, messages = mail.search(None, search_criteria)
                    
                    if status == 'OK' and messages[0]:
                        email_ids = messages[0].split()
                        if email_ids:
                            # Get the most recent
                            latest_email_id = email_ids[-1]
                            status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
                            if status == 'OK':
                                found_email_data = msg_data[0][1]
                                self.logger.info(f"Found report in {folder}")
                                break
                except Exception as e:
                    self.logger.debug(f"Could not search {folder}: {str(e)}")
                    continue
            
            if not found_email_data:
                raise Exception(
                    f"Could not find previous report with subject containing '{search_subject}'. "
                    f"Please ensure you sent/received the report on {prev_day.strftime('%A, %d %B %Y')}."
                )
            
            # Parse email and extract attachment
            msg = email.message_from_bytes(found_email_data)
            
            attachment_path = None
            for part in msg.walk():
                content_disposition = part.get_content_disposition()
                if content_disposition == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        decoded_filename = decode_header(filename)[0][0]
                        if isinstance(decoded_filename, bytes):
                            filename = decoded_filename.decode()
                        
                        if filename.lower().endswith(('.xlsx', '.xls')):
                            attachment_path = save_directory / f"previous_report.xlsx"
                            with open(attachment_path, 'wb') as f:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    f.write(payload)
                            self.logger.info(f"Downloaded previous report: {filename} -> {attachment_path}")
                            break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception("No Excel attachment found in the previous report email.")
            
            return attachment_path
            
        finally:
            if mail:
                try:
                    mail.close()
                    mail.logout()
                except:
                    pass
    
    def download_daily_attachment(self, save_directory: Path, current_date: str) -> Path:
        """
        Download the daily worksheet attachment from email.
        Searches for emails with subject containing 'large trade td'.
        
        Args:
            save_directory: Directory to save the downloaded attachment
            current_date: Current date string (YYYY-MM-DD)
            
        Returns:
            Path to the downloaded file
            
        Raises:
            Exception: If email retrieval or download fails
        """
        if self.use_outlook:
            return self._download_via_outlook(save_directory, current_date)
        else:
            return self._download_via_imap(save_directory, current_date)
    
    def _download_via_outlook(self, save_directory: Path, current_date: str) -> Path:
        """
        Download the daily worksheet using Outlook COM interface.
        Searches for emails with subject containing the configured pattern.
        
        Args:
            save_directory: Directory to save the downloaded attachment
            current_date: Current date string (YYYY-MM-DD)
            
        Returns:
            Path to the downloaded file
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
                except Exception as e:
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
                    except Exception as e:
                        continue
            
            if not found_email:
                raise Exception(
                    f"No emails found with subject containing '{search_subject}'. "
                    "Please ensure you have received the daily email."
                )
            
            # Find and download Excel attachment
            attachment_path = None
            attachments = found_email.Attachments
            
            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                filename = attachment.FileName
                
                if filename and (filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls')):
                    attachment_path = save_directory / f"daily_sheet_{current_date}.xlsx"
                    attachment.SaveAsFile(str(attachment_path))
                    self.logger.info(f"Downloaded attachment: {filename} -> {attachment_path}")
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
    
    def _download_via_imap(self, save_directory: Path, current_date: str) -> Path:
        """
        Download the daily worksheet using IMAP.
        Searches for emails with subject containing the configured pattern.
        
        Args:
            save_directory: Directory to save the downloaded attachment
            current_date: Current date string (YYYY-MM-DD)
            
        Returns:
            Path to the downloaded file
        """
        self.logger.info("Connecting to email server via IMAP...")
        mail = None
        
        # Get the subject with date placeholders replaced
        search_subject = self.get_incoming_subject()
        
        try:
            # Connect to email server
            self.logger.debug(f"Connecting to IMAP server: {self.imap_server}:{self.imap_port}")
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            mail.select('inbox')
            
            self.logger.info(f"Searching for emails with subject containing '{search_subject}'...")
            
            # Search for emails by subject from today
            today = datetime.now().strftime('%d-%b-%Y')
            search_criteria = f'(SUBJECT "{search_subject}" SINCE {today})'
            status, messages = mail.search(None, search_criteria)
            
            if status != 'OK':
                raise Exception("Failed to search emails")
            
            email_ids = messages[0].split()
            
            # If no emails from today, try without date restriction (get most recent)
            if not email_ids:
                self.logger.warning("No emails found from today. Searching for most recent email...")
                search_criteria = f'(SUBJECT "{search_subject}")'
                status, messages = mail.search(None, search_criteria)
                
                if status != 'OK':
                    raise Exception("Failed to search emails")
                
                email_ids = messages[0].split()
            
            if not email_ids:
                raise Exception(
                    f"No emails found with subject containing '{search_subject}'. "
                    "Please ensure you have received the daily email."
                )
            
            # Get the most recent email
            latest_email_id = email_ids[-1]
            self.logger.info(f"Found email ID: {latest_email_id.decode()}")
            
            status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
            
            if status != 'OK':
                raise Exception("Failed to fetch email")
            
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            
            # Decode subject for logging
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            self.logger.info(f"Email subject: {subject}")
            
            # Find and download Excel attachment
            attachment_path = None
            for part in msg.walk():
                content_disposition = part.get_content_disposition()
                if content_disposition == 'attachment':
                    filename = part.get_filename()
                    
                    # Decode filename if needed
                    if filename:
                        decoded_filename = decode_header(filename)[0][0]
                        if isinstance(decoded_filename, bytes):
                            filename = decoded_filename.decode()
                        
                        if filename and (filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls')):
                            attachment_path = save_directory / f"daily_sheet_{current_date}.xlsx"
                            
                            # Download attachment
                            with open(attachment_path, 'wb') as f:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    f.write(payload)
                            
                            self.logger.info(f"Downloaded attachment: {filename} -> {attachment_path}")
                            break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    "No Excel attachment (.xlsx or .xls) found in email. "
                    "Please ensure the email contains an Excel file attachment."
                )
            
            return attachment_path
            
        except imaplib.IMAP4.error as e:
            self.logger.error(f"IMAP error: {str(e)}")
            raise Exception(f"IMAP error: {str(e)}")
        except Exception as e:
            self.logger.error(f"Error downloading email attachment: {str(e)}")
            if "login" in str(e).lower() or "authentication" in str(e).lower():
                error_msg = (
                    f"Email authentication failed: {str(e)}\n"
                    "For Gmail/Outlook, make sure you're using an App Password."
                )
                self.logger.error(error_msg)
                raise Exception(error_msg)
            raise Exception(f"Error downloading email attachment: {str(e)}")
        finally:
            if mail:
                try:
                    mail.close()
                    mail.logout()
                    self.logger.debug("IMAP connection closed")
                except Exception as e:
                    self.logger.warning(f"Error closing IMAP connection: {str(e)}")
    
    def send_report_email(
        self, 
        report_path: Path, 
        distribution_list: List[str], 
        sender_name: str, 
        current_date: str,
        email_body: str = None
    ) -> None:
        """
        Send or preview the report via email to the distribution list.
        Uses Outlook COM if configured, otherwise uses SMTP.
        
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
        if self.use_outlook:
            self._send_via_outlook(report_path, distribution_list, sender_name, current_date, email_body)
        else:
            self._send_via_smtp(report_path, distribution_list, sender_name, current_date, email_body)
    
    def _send_via_outlook(
        self,
        report_path: Path,
        distribution_list: List[str],
        sender_name: str,
        current_date: str,
        email_body: str = None
    ) -> None:
        """
        Send or preview email using Outlook COM interface (Windows only).
        
        If preview_before_send is True, displays the email for review instead of sending.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (DD/MM/YYYY format)
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
    
    def _send_via_smtp(
        self,
        report_path: Path,
        distribution_list: List[str],
        sender_name: str,
        current_date: str,
        email_body: str = None
    ) -> None:
        """
        Send email using SMTP.
        
        Note: SMTP mode does not support preview - emails are sent directly.
        For preview functionality, use Outlook mode instead.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (DD/MM/YYYY format)
            email_body: Optional custom email body text
            
        Raises:
            Exception: If email sending fails
        """
        if self.preview_before_send:
            self.logger.warning(
                "Preview mode is not supported with SMTP. "
                "Email will be sent directly. Use Outlook mode for preview functionality."
            )
        
        self.logger.info("Preparing email via SMTP...")
        server = None
        
        try:
            # Validate report file exists
            if not report_path.exists():
                raise FileNotFoundError(f"Report file not found: {report_path}")
            
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.username
            msg['To'] = ', '.join(distribution_list)
            msg['Subject'] = f"Large Deal Report - {current_date}"
            
            # Email body - use custom or default
            if email_body:
                body = email_body
            else:
                body = (
                    f"Hi all,\n\n"
                    f"Please find attached today's Large Deal Report for {current_date}.\n\n"
                    f"Kind regards,\n{sender_name}"
                )
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach report
            self.logger.info(f"Attaching report: {report_path.name}")
            with open(report_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {report_path.name}'
            )
            msg.attach(part)
            
            # Send email
            self.logger.info(f"Sending email to {len(distribution_list)} recipient(s)...")
            self.logger.debug(f"Connecting to SMTP server: {self.smtp_server}:{self.smtp_port}")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            
            self.logger.info("Email sent successfully via SMTP!")
            
        except smtplib.SMTPAuthenticationError as e:
            error_msg = (
                f"SMTP authentication failed: {str(e)}\n"
                "Make sure you're using an App Password, not your regular password."
            )
            self.logger.error(error_msg)
            raise Exception(error_msg)
        except FileNotFoundError as e:
            self.logger.error(f"Report file not found: {str(e)}")
            raise
        except Exception as e:
            self.logger.error(f"Error sending email: {str(e)}")
            raise Exception(f"Error sending email: {str(e)}")
        finally:
            if server:
                try:
                    server.quit()
                    self.logger.debug("SMTP connection closed")
                except Exception as e:
                    self.logger.warning(f"Error closing SMTP connection: {str(e)}")

