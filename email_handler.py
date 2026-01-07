#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (Outlook COM) functionality.
"""

import imaplib
import email
import logging
from email.header import decode_header
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
        """
        self.email_config = email_config
        self.username = email_config.get('username')
        self.password = email_config.get('password')
        self.sender_email = email_config.get('sender_email') # Optional if we just search by subject
        self.target_subject = email_config.get('target_subject_keyword', 'large trade td')
        
        self.imap_server = email_config.get('imap_server', 'outlook.office365.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.use_outlook_for_sending = email_config.get('use_outlook_for_sending', True)
        
        self.logger = logging.getLogger(__name__)
        
        if self.use_outlook_for_sending and not OUTLOOK_COM_AVAILABLE:
             self.logger.warning("Outlook COM requested but win32com not available. Sending will fail.")
    
    def download_daily_attachment(self, save_directory: Path, current_date: str) -> Path:
        """
        Download the daily worksheet attachment from email.
        """
        self.logger.info("Connecting to email server...")
        mail = None
        
        try:
            # Connect to email server
            self.logger.debug(f"Connecting to IMAP server: {self.imap_server}:{self.imap_port}")
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            mail.select('inbox')
            
            self.logger.info(f"Searching for emails with subject '{self.target_subject}'...")
            
            # IMAP search criteria are tricky with partial subject matches in standard IMAP.
            # We'll fetch headers of recent messages and filter in Python for robustness.
            # Fetch emails from the last few days just in case.
            
            status, messages = mail.search(None, 'ALL') # We will filter by date/sender manually or use SINCE
            
            if status != 'OK':
                raise Exception("Failed to search emails")
            
            email_ids = messages[0].split()
            
            # Look at the last 20 emails (assuming volume isn't massive, or we can use SINCE)
            # Using SINCE is better
            today_str = datetime.now().strftime('%d-%b-%Y')
            status, messages = mail.search(None, f'(SINCE "{today_str}")')
             
            if status == 'OK':
                 email_ids = messages[0].split()
            
            if not email_ids:
                self.logger.warning("No emails found from today. Checking last 10 messages regardless of date...")
                status, messages = mail.search(None, 'ALL')
                email_ids = messages[0].split()[-10:]

            # Iterate backwards to find the latest
            target_email_id = None
            found_subject = ""
            
            for eid in reversed(email_ids):
                status, msg_data = mail.fetch(eid, '(RFC822.HEADER)')
                if status != 'OK': continue
                
                header_data = msg_data[0][1]
                msg_header = email.message_from_bytes(header_data)
                
                subject_header = msg_header["Subject"]
                decoded_list = decode_header(subject_header)
                subject = ""
                for part, encoding in decoded_list:
                    if isinstance(part, bytes):
                        subject += part.decode(encoding or 'utf-8')
                    else:
                        subject += part
                
                # Check match
                if self.target_subject.lower() in subject.lower():
                    # Also verify sender if configured
                    if self.sender_email:
                        from_header = msg_header.get("From", "")
                        if self.sender_email.lower() in from_header.lower():
                             target_email_id = eid
                             found_subject = subject
                             break
                    else:
                        target_email_id = eid
                        found_subject = subject
                        break
            
            if not target_email_id:
                raise Exception(f"No email found with subject containing '{self.target_subject}' today.")
                
            self.logger.info(f"Found matching email: {found_subject}")
            
            # Fetch full email body
            status, msg_data = mail.fetch(target_email_id, '(RFC822)')
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            
            # Find and download Excel attachment
            attachment_path = None
            for part in msg.walk():
                content_disposition = part.get_content_disposition()
                if content_disposition == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        # Decode filename
                        decoded_list = decode_header(filename)
                        filename = ""
                        for part_fn, encoding in decoded_list:
                            if isinstance(part_fn, bytes):
                                filename += part_fn.decode(encoding or 'utf-8')
                            else:
                                filename += part_fn

                        if filename.lower().endswith(('.xlsx', '.xls', '.csv')):
                            # Use original filename extension but standardize name if needed, 
                            # or just use a temp name
                            ext = Path(filename).suffix
                            attachment_path = save_directory / f"daily_received_sheet{ext}"
                            
                            with open(attachment_path, 'wb') as f:
                                payload = part.get_payload(decode=True)
                                f.write(payload)
                            
                            self.logger.info(f"Downloaded attachment: {filename}")
                            break
            
            if not attachment_path:
                raise Exception("No Excel attachment found in the target email.")
                
            return attachment_path
            
        except Exception as e:
            self.logger.error(f"Email error: {str(e)}")
            raise
        finally:
            if mail:
                try:
                    mail.close()
                    mail.logout()
                except:
                    pass

    def send_report_email(
        self, 
        report_path: Path, 
        distribution_list: List[str], 
        sender_name: str, 
        current_date: str
    ) -> None:
        """
        Creates a draft email in Outlook with the report attached and displays it.
        """
        if not self.use_outlook_for_sending:
             raise NotImplementedError("Only Outlook sending is supported for the 'preview' feature.")

        self.logger.info("Creating Outlook email draft...")
        
        try:
            if not report_path.exists():
                raise FileNotFoundError(f"Report file not found: {report_path}")
            
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = MailItem
            
            mail.Subject = f"Large Deal Report - {current_date}"
            
            # Construct body
            mail.Body = (
                f"Hi All,\n\n"
                f"Please find attached the Large Deal Report for {current_date}.\n\n"
                f"Kind regards,\n{sender_name}"
            )
            
            # Add recipients
            for recipient in distribution_list:
                mail.Recipients.Add(recipient)
            
            # Attach report
            mail.Attachments.Add(str(report_path.absolute()))
            
            # Display (Preview) instead of Send
            mail.Display()
            
            self.logger.info("Email draft created and displayed.")
            
        except Exception as e:
            self.logger.error(f"Error creating Outlook email: {str(e)}")
            raise
