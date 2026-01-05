#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP or Outlook COM) functionality.
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
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address (optional if using Outlook)
                - smtp_port: SMTP server port (optional if using Outlook)
                - use_outlook_for_sending: If True, use Outlook COM instead of SMTP (Windows only)
        """
        self.email_config = email_config
        self.username = email_config.get('username')
        self.password = email_config.get('password')
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.use_outlook_for_sending = email_config.get('use_outlook_for_sending', False)
        self.logger = logging.getLogger(__name__)
        
        # Validate Outlook availability if requested
        if self.use_outlook_for_sending:
            if not OUTLOOK_COM_AVAILABLE:
                raise ImportError(
                    "Outlook COM requested but win32com not available.\n"
                    "Install it with: pip install pywin32\n"
                    "Note: This requires Windows and Outlook to be installed."
                )
            self.logger.info("Will use Outlook COM interface for sending emails")
        else:
            # Validate SMTP credentials are provided
            if not self.username or not self.password:
                raise ValueError(
                    "SMTP credentials (username and password) are required "
                    "when not using Outlook for sending."
                )
    
    def download_daily_attachment(self, save_directory: Path, current_date: str) -> Path:
        """
        Download the daily worksheet attachment from email.
        
        Args:
            save_directory: Directory to save the downloaded attachment
            current_date: Current date string (YYYY-MM-DD)
            
        Returns:
            Path to the downloaded file
            
        Raises:
            Exception: If email retrieval or download fails
        """
        self.logger.info("Connecting to email server...")
        mail = None
        
        try:
            # Connect to email server
            self.logger.debug(f"Connecting to IMAP server: {self.imap_server}:{self.imap_port}")
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            mail.select('inbox')
            
            self.logger.info(f"Searching for emails from {self.sender_email}...")
            
            # Search for emails from today with attachments
            today = datetime.now().strftime('%d-%b-%Y')
            search_criteria = f'(FROM "{self.sender_email}" SINCE {today})'
            status, messages = mail.search(None, search_criteria)
            
            if status != 'OK':
                raise Exception("Failed to search emails")
            
            email_ids = messages[0].split()
            
            # If no emails from today, try without date restriction (get most recent)
            if not email_ids:
                self.logger.warning("No emails found from today. Searching for most recent email...")
                search_criteria = f'(FROM "{self.sender_email}")'
                status, messages = mail.search(None, search_criteria)
                
                if status != 'OK':
                    raise Exception("Failed to search emails")
                
                email_ids = messages[0].split()
            
            if not email_ids:
                raise Exception(
                    f"No emails found from {self.sender_email}. "
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
                        
                        if filename and (filename.endswith('.xlsx') or filename.endswith('.xls')):
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
                    "For Gmail, make sure you're using an App Password, not your regular password."
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
        current_date: str
    ) -> None:
        """
        Send the report via email to the distribution list.
        Uses Outlook COM if configured, otherwise uses SMTP.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        if self.use_outlook_for_sending:
            self._send_via_outlook(report_path, distribution_list, sender_name, current_date)
        else:
            self._send_via_smtp(report_path, distribution_list, sender_name, current_date)
    
    def _send_via_outlook(
        self,
        report_path: Path,
        distribution_list: List[str],
        sender_name: str,
        current_date: str
    ) -> None:
        """
        Send email using Outlook COM interface (Windows only).
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        self.logger.info("Preparing email in Outlook...")
        
        try:
            # Validate report file exists
            if not report_path.exists():
                raise FileNotFoundError(f"Report file not found: {report_path}")
            
            # Connect to Outlook
            self.logger.debug("Connecting to Outlook application...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = MailItem
            
            # Set email properties
            mail.Subject = f"Large Deal Report - {current_date}"
            mail.Body = (
                f"Hi all,\n\n"
                f"Please see the large deal report for {current_date}.\n\n"
                f"Kind regards,\n{sender_name}"
            )
            
            # Add recipients
            self.logger.info(f"Adding {len(distribution_list)} recipient(s)...")
            for recipient in distribution_list:
                mail.Recipients.Add(recipient)
            
            # Attach report
            self.logger.info(f"Attaching report: {report_path.name}")
            mail.Attachments.Add(str(report_path))
            
            # Send email
            self.logger.info("Sending email via Outlook...")
            mail.Send()
            
            self.logger.info("Email sent successfully via Outlook!")
            
        except FileNotFoundError as e:
            self.logger.error(f"Report file not found: {str(e)}")
            raise
        except Exception as e:
            error_msg = f"Error sending email via Outlook: {str(e)}"
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
        current_date: str
    ) -> None:
        """
        Send email using SMTP.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
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
            
            # Email body
            body = (
                f"Hi all,\n\n"
                f"Please see the large deal report for {current_date}.\n\n"
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
                "For Gmail, make sure you're using an App Password, not your regular password."
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

