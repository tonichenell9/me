#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP or Outlook COM) functionality.
"""

import imaplib
import email
import smtplib
import logging
import re
import time
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
        self.use_outlook_for_downloading = email_config.get('use_outlook_for_downloading', False)
        self.incoming_subject_contains = (email_config.get('incoming_subject_contains') or "").strip()
        self.outlook_send_email = bool(email_config.get('outlook_send_email', False))
        self.outlook_use_signature = bool(email_config.get('outlook_use_signature', True))
        self.email_subject_template = email_config.get('email_subject_template') or "Large Deal Report - {date}"
        self.email_body_template = email_config.get('email_body_template') or (
            "Hi all,<br><br>Please see the large deal report for {date}.<br><br>Kind regards,<br>{sender_name}"
        )
        self.logger = logging.getLogger(__name__)
        
        # Validate Outlook availability if requested
        if self.use_outlook_for_sending or self.use_outlook_for_downloading:
            if not OUTLOOK_COM_AVAILABLE:
                raise ImportError(
                    "Outlook COM requested but win32com not available.\n"
                    "Install it with: pip install pywin32\n"
                    "Note: This requires Windows and Outlook to be installed."
                )
            if self.use_outlook_for_sending:
                self.logger.info("Will use Outlook COM interface for sending emails")
            if self.use_outlook_for_downloading:
                self.logger.info("Will use Outlook COM interface for downloading attachments")

        # If sending via SMTP, validate credentials
        if not self.use_outlook_for_sending:
            if not self.username or not self.password:
                raise ValueError(
                    "SMTP credentials (username and password) are required "
                    "when not using Outlook for sending."
                )

    def _render_subject_and_bodies(self, current_date: str, sender_name: str) -> tuple[str, str, str]:
        """
        Render subject + HTML + plain text bodies.
        """
        subject = str(self.email_subject_template).format(date=current_date, sender_name=sender_name)
        body_html = str(self.email_body_template).format(date=current_date, sender_name=sender_name)
        body_text = re.sub(r"<[^>]+>", "", body_html)
        return subject, body_html, body_text
    
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
        if self.use_outlook_for_downloading:
            return self._download_daily_attachment_via_outlook(save_directory, current_date)

        if not self.username or not self.password:
            raise ValueError(
                "IMAP credentials (username and password) are required when not using Outlook for downloading."
            )

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
            
            # Iterate newest -> oldest, pick first message matching subject (if configured) with Excel attachment
            attachment_path = None
            for email_id in reversed(email_ids):
                status, msg_data = mail.fetch(email_id, '(RFC822)')
                if status != 'OK':
                    continue

                email_body = msg_data[0][1]
                msg = email.message_from_bytes(email_body)

                subject = decode_header(msg.get("Subject", ""))[0][0]
                if isinstance(subject, bytes):
                    subject = subject.decode(errors="replace")
                subject_str = str(subject or "")

                if self.incoming_subject_contains:
                    if self.incoming_subject_contains.lower() not in subject_str.lower():
                        continue

                self.logger.info(f"Using email subject: {subject_str}")

                for part in msg.walk():
                    content_disposition = part.get_content_disposition()
                    if content_disposition != 'attachment':
                        continue

                    filename = part.get_filename()
                    if not filename:
                        continue

                    decoded_filename = decode_header(filename)[0][0]
                    if isinstance(decoded_filename, bytes):
                        filename = decoded_filename.decode(errors="replace")
                    filename_str = str(filename or "")

                    if not (filename_str.lower().endswith('.xlsx') or filename_str.lower().endswith('.xls')):
                        continue

                    attachment_path = save_directory / f"daily_sheet_{current_date}.xlsx"
                    with open(attachment_path, 'wb') as f:
                        payload = part.get_payload(decode=True)
                        if payload:
                            f.write(payload)

                    self.logger.info(f"Downloaded attachment: {filename_str} -> {attachment_path}")
                    break

                if attachment_path and attachment_path.exists():
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

    def _download_daily_attachment_via_outlook(self, save_directory: Path, current_date: str) -> Path:
        """
        Download the daily worksheet attachment using Outlook COM (Windows only).
        This is typically the most reliable approach in corporate environments.

        Criteria:
        - Subject contains `incoming_subject_contains` (case-insensitive) if provided
        - Received today (local time) when possible
        - Has an Excel attachment
        """
        self.logger.info("Searching Outlook inbox for daily attachment...")

        if not self.incoming_subject_contains:
            self.logger.warning(
                "No 'incoming_subject_contains' configured; Outlook download will use the newest mail with an Excel attachment."
            )

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

            items = inbox.Items
            items.Sort("[ReceivedTime]", True)  # newest first

            # Restrict to today where possible (Outlook Restrict uses US-style date strings)
            today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            today_str = today_start.strftime("%m/%d/%Y %H:%M %p")
            try:
                items = items.Restrict(f"[ReceivedTime] >= '{today_str}'")
            except Exception as e:
                self.logger.debug(f"Outlook Restrict failed; scanning all items. Details: {e}")

            attachment_path = save_directory / f"daily_sheet_{current_date}.xlsx"

            # Iterate mails
            for item in items:
                try:
                    subject = str(getattr(item, "Subject", "") or "")
                    if self.incoming_subject_contains and self.incoming_subject_contains.lower() not in subject.lower():
                        continue

                    # Best-effort sender filter (may differ for Exchange accounts)
                    sender_addr = str(getattr(item, "SenderEmailAddress", "") or "")
                    if sender_addr and self.sender_email and self.sender_email.lower() not in sender_addr.lower():
                        continue

                    attachments = getattr(item, "Attachments", None)
                    if not attachments or attachments.Count == 0:
                        continue

                    # Pick first Excel attachment
                    for i in range(1, attachments.Count + 1):
                        att = attachments.Item(i)
                        name = str(getattr(att, "FileName", "") or "")
                        if not name:
                            continue
                        if not (name.lower().endswith(".xlsx") or name.lower().endswith(".xls")):
                            continue

                        att.SaveAsFile(str(attachment_path))
                        self.logger.info(f"Downloaded Outlook attachment: {name} -> {attachment_path}")
                        return attachment_path
                except Exception:
                    # Skip problematic items and keep searching
                    continue

            raise Exception(
                "No matching Outlook email with an Excel attachment was found. "
                "Check the inbox subject and that the attachment is an .xlsx/.xls file."
            )

        except Exception as e:
            raise Exception(f"Error downloading attachment via Outlook: {str(e)}")
    
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
            subject, body_html, body_text = self._render_subject_and_bodies(current_date, sender_name)
            mail.Subject = subject
            
            # Add recipients
            self.logger.info(f"Adding {len(distribution_list)} recipient(s)...")
            for recipient in distribution_list:
                mail.Recipients.Add(recipient)
            
            # Attach report
            self.logger.info(f"Attaching report: {report_path.name}")
            mail.Attachments.Add(str(report_path))
            
            # Display draft (recommended) or send immediately if configured
            if self.outlook_use_signature:
                # Display first to let Outlook insert signature, then prepend our body.
                mail.Display()
                try:
                    existing_html = str(getattr(mail, "HTMLBody", "") or "")
                    mail.HTMLBody = body_html + "<br><br>" + existing_html
                except Exception:
                    # Fallback to plain text
                    mail.Body = body_text + "\n\n" + str(getattr(mail, "Body", "") or "")
            else:
                # No signature needed; set body directly and display
                try:
                    mail.HTMLBody = body_html
                except Exception:
                    mail.Body = body_text
                mail.Display()

            if self.outlook_send_email:
                self.logger.info("Sending email via Outlook (outlook_send_email=true)...")
                mail.Send()
                self.logger.info("Email sent successfully via Outlook!")
            else:
                self.logger.info("Outlook email draft opened for review (not auto-sent).")
            
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
            subject, body_html, body_text = self._render_subject_and_bodies(current_date, sender_name)
            msg['Subject'] = subject
            
            # Email body
            msg.attach(MIMEText(body_text, 'plain'))
            msg.attach(MIMEText(body_html, 'html'))
            
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

