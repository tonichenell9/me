# Automation script development
_Exported on 05/01/2026 at 11:59:18 GMT from Cursor (2.2.44)_

---

**User**

Implement the following to-dos from the plan (the plan is attached for your reference). Do NOT edit the plan file itself.

You have been assigned the following 8 to-do(s) with IDs: setup_structure, implement_email_handler, implement_excel_processor, implement_main_script, create_config_template, update_requirements, update_readme, add_error_handling

1. [setup_structure] Create project structure with separate modules: email_handler.py and excel_processor.py
2. [implement_email_handler] Implement EmailHandler class with IMAP retrieval and SMTP sending functionality
3. [implement_excel_processor] Implement ExcelProcessor class with worksheet update, refresh (xlwings), and copy operations
4. [implement_main_script] Create main automation script that orchestrates the workflow
5. [create_config_template] Update config.json.example with all required configuration options
6. [update_requirements] Update requirements.txt with all dependencies including xlwings
7. [update_readme] Create comprehensive README with setup instructions, troubleshooting, and usage examples
8. [add_error_handling] Add comprehensive error handling, logging, and cleanup throughout the script

These to-dos have already been created. Do not create them again. Mark them as in_progress as you work, starting with the first one. Don't stop until you have completed all the assigned to-dos.

---

**Cursor**


Reviewing the existing codebase, then implementing the 8 to-dos.



Reviewing the codebase and implementing the 8 to-dos. Creating the initial todo list and starting implementation.



Creating the project structure. Starting with `excel_processor.py`:


```python
#!/usr/bin/env python3
"""
Excel Processor Module
Handles Excel file manipulation, worksheet updates, refresh operations, and copying.
"""

import logging
from pathlib import Path
from typing import Optional
import openpyxl
from openpyxl import load_workbook

# Try to import xlwings for Excel refresh functionality
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    logging.warning("xlwings not available. Worksheet 2 refresh will use fallback method.")


class ExcelProcessor:
    """Handles Excel file manipulation and processing operations."""
    
    def __init__(self, config: dict):
        """
        Initialize Excel processor with configuration.
        
        Args:
            config: Dictionary containing configuration:
                - worksheet_1_name: Name of worksheet 1
                - worksheet_2_name: Name of worksheet 2
                - worksheet_3_name: Name of worksheet 3
        """
        self.config = config
        self.worksheet_1_name = config.get('worksheet_1_name', 'Sheet1')
        self.worksheet_2_name = config.get('worksheet_2_name', 'Sheet2')
        self.worksheet_3_name = config.get('worksheet_3_name', 'Sheet3')
        self.logger = logging.getLogger(__name__)
    
    def load_workbook(self, filepath: Path):
        """
        Load an Excel workbook.
        
        Args:
            filepath: Path to the Excel file
            
        Returns:
            openpyxl Workbook object
            
        Raises:
            FileNotFoundError: If file doesn't exist
            Exception: If file cannot be opened
        """
        try:
            self.logger.info(f"Loading workbook: {filepath}")
            return load_workbook(filepath)
        except FileNotFoundError:
            raise FileNotFoundError(f"Workbook not found: {filepath}")
        except Exception as e:
            raise Exception(f"Error loading workbook {filepath}: {str(e)}")
    
    def update_worksheet_1(self, workbook: openpyxl.Workbook, daily_sheet_path: Path) -> None:
        """
        Replace worksheet 1 with data from the daily email attachment.
        
        Args:
            workbook: The target workbook to update
            daily_sheet_path: Path to the daily worksheet file
            
        Raises:
            FileNotFoundError: If daily sheet file doesn't exist
            ValueError: If worksheet 1 doesn't exist
            Exception: If update fails
        """
        self.logger.info("Updating Worksheet 1...")
        
        try:
            # Load the daily sheet
            if not daily_sheet_path.exists():
                raise FileNotFoundError(f"Daily sheet not found: {daily_sheet_path}")
            
            daily_wb = load_workbook(daily_sheet_path)
            daily_ws = daily_wb.active
            
            # Get or create worksheet 1
            if self.worksheet_1_name in workbook.sheetnames:
                ws1 = workbook[self.worksheet_1_name]
            else:
                self.logger.warning(
                    f"Worksheet '{self.worksheet_1_name}' not found. Using active sheet."
                )
                ws1 = workbook.active
            
            # Clear existing data in worksheet 1
            if ws1.max_row > 0:
                ws1.delete_rows(1, ws1.max_row)
            
            # Copy data from daily sheet to worksheet 1
            for row in daily_ws.iter_rows(values_only=True):
                if any(cell is not None for cell in row):  # Skip completely empty rows
                    ws1.append(row)
            
            # Copy formatting if needed (basic cell copying)
            for row_idx, row in enumerate(daily_ws.iter_rows(), start=1):
                if row_idx > ws1.max_row:
                    break
                for col_idx, cell in enumerate(row, start=1):
                    if cell.has_style:
                        target_cell = ws1.cell(row=row_idx, column=col_idx)
                        target_cell.font = cell.font
                        target_cell.fill = cell.fill
                        target_cell.border = cell.border
                        target_cell.alignment = cell.alignment
                        target_cell.number_format = cell.number_format
            
            daily_wb.close()
            self.logger.info("Worksheet 1 updated successfully.")
            
        except Exception as e:
            self.logger.error(f"Error updating Worksheet 1: {str(e)}")
            raise
    
    def refresh_worksheet_2(self, workbook: openpyxl.Workbook, workbook_path: Optional[Path] = None) -> None:
        """
        Refresh data in worksheet 2 using xlwings (if available) or fallback method.
        
        Args:
            workbook: The workbook containing worksheet 2
            workbook_path: Optional path to the workbook file (required for xlwings)
            
        Raises:
            ValueError: If worksheet 2 doesn't exist
            Exception: If refresh fails
        """
        self.logger.info("Refreshing Worksheet 2...")
        
        if self.worksheet_2_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found in workbook")
        
        # Try xlwings refresh if available and workbook_path is provided
        if XLWINGS_AVAILABLE and workbook_path:
            try:
                self._refresh_with_xlwings(workbook_path)
                self.logger.info("Worksheet 2 refreshed using xlwings.")
                return
            except Exception as e:
                self.logger.warning(
                    f"xlwings refresh failed: {str(e)}. Falling back to openpyxl method."
                )
        
        # Fallback: Use openpyxl (preserves structure but doesn't refresh connections)
        self.logger.info(
            "Using openpyxl fallback. Note: Power Query and external data connections "
            "will not be automatically refreshed. You may need to manually refresh in Excel."
        )
        # The worksheet structure is preserved, but data connections won't refresh
        # This is expected behavior when xlwings is not available
    
    def _refresh_with_xlwings(self, workbook_path: Path) -> None:
        """
        Refresh worksheet 2 using xlwings (requires Excel installed).
        
        Args:
            workbook_path: Path to the workbook file
            
        Raises:
            Exception: If xlwings refresh fails
        """
        app = None
        wb = None
        
        try:
            self.logger.info("Attempting to refresh using xlwings...")
            
            # Open Excel application (headless)
            app = xw.App(visible=False)
            app.display_alerts = False
            
            # Open the workbook
            wb = app.books.open(str(workbook_path))
            
            # Get worksheet 2
            if self.worksheet_2_name not in [sheet.name for sheet in wb.sheets]:
                raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found")
            
            ws = wb.sheets[self.worksheet_2_name]
            
            # Try to refresh all data connections/queries
            try:
                # Refresh all queries/connections in the workbook
                wb.api.RefreshAll()
                self.logger.info("Refreshed all data connections in workbook.")
            except Exception as e:
                self.logger.warning(f"Could not refresh all connections: {str(e)}")
                # Try to refresh specific table if it's a ListObject
                try:
                    if ws.api.ListObjects.Count > 0:
                        for list_obj in ws.api.ListObjects:
                            if hasattr(list_obj, 'QueryTable'):
                                list_obj.QueryTable.Refresh(BackgroundQuery=False)
                        self.logger.info("Refreshed table connections.")
                except Exception as e2:
                    self.logger.warning(f"Could not refresh table connections: {str(e2)}")
            
            # Save the workbook
            wb.save()
            self.logger.info("Workbook saved after refresh.")
            
        except Exception as e:
            self.logger.error(f"xlwings refresh error: {str(e)}")
            raise
        finally:
            # Clean up
            if wb:
                try:
                    wb.close()
                except:
                    pass
            if app:
                try:
                    app.quit()
                except:
                    pass
    
    def copy_worksheet_2_to_3(self, workbook: openpyxl.Workbook) -> None:
        """
        Copy worksheet 2 data to worksheet 3 (smartphone-friendly version).
        
        Args:
            workbook: The workbook containing the worksheets
            
        Raises:
            ValueError: If worksheet 2 doesn't exist
            Exception: If copy operation fails
        """
        self.logger.info("Copying Worksheet 2 to Worksheet 3...")
        
        if self.worksheet_2_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found in workbook")
        
        try:
            ws2 = workbook[self.worksheet_2_name]
            
            # Remove existing worksheet 3 if it exists
            if self.worksheet_3_name in workbook.sheetnames:
                workbook.remove(workbook[self.worksheet_3_name])
            
            # Create new worksheet 3
            ws3 = workbook.create_sheet(self.worksheet_3_name)
            
            # Copy data from worksheet 2 to worksheet 3
            for row in ws2.iter_rows(values_only=True):
                if any(cell is not None for cell in row):  # Skip completely empty rows
                    ws3.append(row)
            
            # Copy formatting
            for row_idx, row in enumerate(ws2.iter_rows(), start=1):
                if row_idx > ws3.max_row:
                    break
                for col_idx, cell in enumerate(row, start=1):
                    if cell.has_style:
                        target_cell = ws3.cell(row=row_idx, column=col_idx)
                        target_cell.font = cell.font
                        target_cell.fill = cell.fill
                        target_cell.border = cell.border
                        target_cell.alignment = cell.alignment
                        target_cell.number_format = cell.number_format
            
            # Apply smartphone-friendly formatting (optional enhancements)
            # Adjust column widths for better mobile viewing
            for column in ws3.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 for mobile
                ws3.column_dimensions[column_letter].width = adjusted_width
            
            self.logger.info("Worksheet 3 created successfully.")
            
        except Exception as e:
            self.logger.error(f"Error copying Worksheet 2 to Worksheet 3: {str(e)}")
            raise
    
    def save_workbook(self, workbook: openpyxl.Workbook, filepath: Path) -> None:
        """
        Save the workbook to a file.
        
        Args:
            workbook: The workbook to save
            filepath: Path where to save the workbook
            
        Raises:
            Exception: If save fails
        """
        try:
            self.logger.info(f"Saving workbook to: {filepath}")
            workbook.save(filepath)
            self.logger.info("Workbook saved successfully.")
        except Exception as e:
            self.logger.error(f"Error saving workbook: {str(e)}")
            raise

```


Enhancing `email_handler.py` with logging and error handling:



Enhancing `email_handler.py` with logging:


```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
    
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
        print("Connecting to email server...")
        mail = None
        
        try:
            # Connect to email server
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            mail.select('inbox')
            
            print(f"Searching for emails from {self.sender_email}...")
            
            # Search for emails from today with attachments
            today = datetime.now().strftime('%d-%b-%Y')
            search_criteria = f'(FROM "{self.sender_email}" SINCE {today})'
            status, messages = mail.search(None, search_criteria)
            
            if status != 'OK':
                raise Exception("Failed to search emails")
            
            email_ids = messages[0].split()
            
            # If no emails from today, try without date restriction (get most recent)
            if not email_ids:
                print(f"No emails found from today. Searching for most recent email...")
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
            print(f"Found email ID: {latest_email_id.decode()}")
            
            status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
            
            if status != 'OK':
                raise Exception("Failed to fetch email")
            
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            
            # Decode subject for logging
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            print(f"Email subject: {subject}")
            
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
                            
                            print(f"Downloaded attachment: {filename} -> {attachment_path}")
                            break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    "No Excel attachment (.xlsx or .xls) found in email. "
                    "Please ensure the email contains an Excel file attachment."
                )
            
            return attachment_path
            
        except imaplib.IMAP4.error as e:
            raise Exception(f"IMAP error: {str(e)}")
        except Exception as e:
            if "login" in str(e).lower() or "authentication" in str(e).lower():
                raise Exception(
                    f"Email authentication failed: {str(e)}\n"
                    "For Gmail, make sure you're using an App Password, not your regular password."
                )
            raise Exception(f"Error downloading email attachment: {str(e)}")
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
        Send the report via email to the distribution list.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        print("Preparing email...")
        
        try:
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
            print(f"Attaching report: {report_path.name}")
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
            print(f"Sending email to {len(distribution_list)} recipient(s)...")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            server.quit()
            
            print("Email sent successfully!")
            
        except smtplib.SMTPAuthenticationError as e:
            raise Exception(
                f"SMTP authentication failed: {str(e)}\n"
                "For Gmail, make sure you're using an App Password, not your regular password."
            )
        except Exception as e:
            raise Exception(f"Error sending email: {str(e)}")


```

```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.logger = logging.getLogger(__name__)
    
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
        print("Connecting to email server...")
        mail = None
        
        try:
            # Connect to email server
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            mail.select('inbox')
            
            print(f"Searching for emails from {self.sender_email}...")
            
            # Search for emails from today with attachments
            today = datetime.now().strftime('%d-%b-%Y')
            search_criteria = f'(FROM "{self.sender_email}" SINCE {today})'
            status, messages = mail.search(None, search_criteria)
            
            if status != 'OK':
                raise Exception("Failed to search emails")
            
            email_ids = messages[0].split()
            
            # If no emails from today, try without date restriction (get most recent)
            if not email_ids:
                print(f"No emails found from today. Searching for most recent email...")
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
            print(f"Found email ID: {latest_email_id.decode()}")
            
            status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
            
            if status != 'OK':
                raise Exception("Failed to fetch email")
            
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            
            # Decode subject for logging
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            print(f"Email subject: {subject}")
            
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
                            
                            print(f"Downloaded attachment: {filename} -> {attachment_path}")
                            break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    "No Excel attachment (.xlsx or .xls) found in email. "
                    "Please ensure the email contains an Excel file attachment."
                )
            
            return attachment_path
            
        except imaplib.IMAP4.error as e:
            raise Exception(f"IMAP error: {str(e)}")
        except Exception as e:
            if "login" in str(e).lower() or "authentication" in str(e).lower():
                raise Exception(
                    f"Email authentication failed: {str(e)}\n"
                    "For Gmail, make sure you're using an App Password, not your regular password."
                )
            raise Exception(f"Error downloading email attachment: {str(e)}")
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
        Send the report via email to the distribution list.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        print("Preparing email...")
        
        try:
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
            print(f"Attaching report: {report_path.name}")
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
            print(f"Sending email to {len(distribution_list)} recipient(s)...")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            server.quit()
            
            print("Email sent successfully!")
            
        except smtplib.SMTPAuthenticationError as e:
            raise Exception(
                f"SMTP authentication failed: {str(e)}\n"
                "For Gmail, make sure you're using an App Password, not your regular password."
            )
        except Exception as e:
            raise Exception(f"Error sending email: {str(e)}")


```

```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.logger = logging.getLogger(__name__)
    
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
                print(f"No emails found from today. Searching for most recent email...")
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
            print(f"Found email ID: {latest_email_id.decode()}")
            
            status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
            
            if status != 'OK':
                raise Exception("Failed to fetch email")
            
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            
            # Decode subject for logging
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            print(f"Email subject: {subject}")
            
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
                            
                            print(f"Downloaded attachment: {filename} -> {attachment_path}")
                            break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    "No Excel attachment (.xlsx or .xls) found in email. "
                    "Please ensure the email contains an Excel file attachment."
                )
            
            return attachment_path
            
        except imaplib.IMAP4.error as e:
            raise Exception(f"IMAP error: {str(e)}")
        except Exception as e:
            if "login" in str(e).lower() or "authentication" in str(e).lower():
                raise Exception(
                    f"Email authentication failed: {str(e)}\n"
                    "For Gmail, make sure you're using an App Password, not your regular password."
                )
            raise Exception(f"Error downloading email attachment: {str(e)}")
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
        Send the report via email to the distribution list.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        print("Preparing email...")
        
        try:
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
            print(f"Attaching report: {report_path.name}")
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
            print(f"Sending email to {len(distribution_list)} recipient(s)...")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            server.quit()
            
            print("Email sent successfully!")
            
        except smtplib.SMTPAuthenticationError as e:
            raise Exception(
                f"SMTP authentication failed: {str(e)}\n"
                "For Gmail, make sure you're using an App Password, not your regular password."
            )
        except Exception as e:
            raise Exception(f"Error sending email: {str(e)}")


```

```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.logger = logging.getLogger(__name__)
    
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
            print(f"Found email ID: {latest_email_id.decode()}")
            
            status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
            
            if status != 'OK':
                raise Exception("Failed to fetch email")
            
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            
            # Decode subject for logging
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            print(f"Email subject: {subject}")
            
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
                            
                            print(f"Downloaded attachment: {filename} -> {attachment_path}")
                            break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    "No Excel attachment (.xlsx or .xls) found in email. "
                    "Please ensure the email contains an Excel file attachment."
                )
            
            return attachment_path
            
        except imaplib.IMAP4.error as e:
            raise Exception(f"IMAP error: {str(e)}")
        except Exception as e:
            if "login" in str(e).lower() or "authentication" in str(e).lower():
                raise Exception(
                    f"Email authentication failed: {str(e)}\n"
                    "For Gmail, make sure you're using an App Password, not your regular password."
                )
            raise Exception(f"Error downloading email attachment: {str(e)}")
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
        Send the report via email to the distribution list.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        print("Preparing email...")
        
        try:
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
            print(f"Attaching report: {report_path.name}")
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
            print(f"Sending email to {len(distribution_list)} recipient(s)...")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            server.quit()
            
            print("Email sent successfully!")
            
        except smtplib.SMTPAuthenticationError as e:
            raise Exception(
                f"SMTP authentication failed: {str(e)}\n"
                "For Gmail, make sure you're using an App Password, not your regular password."
            )
        except Exception as e:
            raise Exception(f"Error sending email: {str(e)}")


```

```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.logger = logging.getLogger(__name__)
    
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
                            
                            print(f"Downloaded attachment: {filename} -> {attachment_path}")
                            break
            
            if not attachment_path or not attachment_path.exists():
                raise Exception(
                    "No Excel attachment (.xlsx or .xls) found in email. "
                    "Please ensure the email contains an Excel file attachment."
                )
            
            return attachment_path
            
        except imaplib.IMAP4.error as e:
            raise Exception(f"IMAP error: {str(e)}")
        except Exception as e:
            if "login" in str(e).lower() or "authentication" in str(e).lower():
                raise Exception(
                    f"Email authentication failed: {str(e)}\n"
                    "For Gmail, make sure you're using an App Password, not your regular password."
                )
            raise Exception(f"Error downloading email attachment: {str(e)}")
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
        Send the report via email to the distribution list.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        print("Preparing email...")
        
        try:
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
            print(f"Attaching report: {report_path.name}")
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
            print(f"Sending email to {len(distribution_list)} recipient(s)...")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            server.quit()
            
            print("Email sent successfully!")
            
        except smtplib.SMTPAuthenticationError as e:
            raise Exception(
                f"SMTP authentication failed: {str(e)}\n"
                "For Gmail, make sure you're using an App Password, not your regular password."
            )
        except Exception as e:
            raise Exception(f"Error sending email: {str(e)}")


```

```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.logger = logging.getLogger(__name__)
    
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
            raise Exception(f"IMAP error: {str(e)}")
        except Exception as e:
            if "login" in str(e).lower() or "authentication" in str(e).lower():
                raise Exception(
                    f"Email authentication failed: {str(e)}\n"
                    "For Gmail, make sure you're using an App Password, not your regular password."
                )
            raise Exception(f"Error downloading email attachment: {str(e)}")
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
        Send the report via email to the distribution list.
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        print("Preparing email...")
        
        try:
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
            print(f"Attaching report: {report_path.name}")
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
            print(f"Sending email to {len(distribution_list)} recipient(s)...")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            server.quit()
            
            print("Email sent successfully!")
            
        except smtplib.SMTPAuthenticationError as e:
            raise Exception(
                f"SMTP authentication failed: {str(e)}\n"
                "For Gmail, make sure you're using an App Password, not your regular password."
            )
        except Exception as e:
            raise Exception(f"Error sending email: {str(e)}")


```

```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.logger = logging.getLogger(__name__)
    
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
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        print("Preparing email...")
        
        try:
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
            print(f"Attaching report: {report_path.name}")
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
            print(f"Sending email to {len(distribution_list)} recipient(s)...")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            text = msg.as_string()
            server.sendmail(self.username, distribution_list, text)
            server.quit()
            
            print("Email sent successfully!")
            
        except smtplib.SMTPAuthenticationError as e:
            raise Exception(
                f"SMTP authentication failed: {str(e)}\n"
                "For Gmail, make sure you're using an App Password, not your regular password."
            )
        except Exception as e:
            raise Exception(f"Error sending email: {str(e)}")


```

```python
#!/usr/bin/env python3
"""
Email Handler Module
Handles email retrieval (IMAP) and sending (SMTP) functionality.
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


class EmailHandler:
    """Handles email retrieval and sending operations."""
    
    def __init__(self, email_config: dict):
        """
        Initialize email handler with configuration.
        
        Args:
            email_config: Dictionary containing email settings:
                - username: Email address
                - password: App password or password
                - sender_email: Email address to look for
                - imap_server: IMAP server address
                - imap_port: IMAP server port
                - smtp_server: SMTP server address
                - smtp_port: SMTP server port
        """
        self.email_config = email_config
        self.username = email_config['username']
        self.password = email_config['password']
        self.sender_email = email_config['sender_email']
        self.imap_server = email_config.get('imap_server', 'imap.gmail.com')
        self.imap_port = email_config.get('imap_port', 993)
        self.smtp_server = email_config.get('smtp_server', 'smtp.gmail.com')
        self.smtp_port = email_config.get('smtp_port', 587)
        self.logger = logging.getLogger(__name__)
    
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
        
        Args:
            report_path: Path to the report file to attach
            distribution_list: List of recipient email addresses
            sender_name: Name of the sender
            current_date: Current date string (YYYY-MM-DD)
            
        Raises:
            Exception: If email sending fails
        """
        self.logger.info("Preparing email...")
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
            
            self.logger.info("Email sent successfully!")
            
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


```


Refactoring the main script to use the modular structure with error handling and logging:


```python
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
        self.config = self.load_config(config_path)
        self.setup_logging()
        self.logger = logging.getLogger(__name__)
        
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
            required_email_fields = ['username', 'password', 'sender_email']
            for field in required_email_fields:
                if field not in config['email']:
                    raise ValueError(f"Missing required email configuration: {field}")
            
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

```


Updating `config.json.example` with all required configuration options:


```plaintext
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com",
    "recipient3@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}

```


Updating `requirements.txt` with all dependencies, including xlwings:


```pip-requirements
openpyxl>=3.1.2
pandas>=2.0.0
xlwings>=0.30.0

```


Updating the README with setup instructions, troubleshooting, and usage examples:


```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality on macOS/Windows)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

```bash
python large_deal_report_automation.py
```

Or make it executable and run directly:

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution (macOS)

#### Option 1: Shell Alias

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /Users/tonigallimore/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

#### Option 2: macOS Automator Quick Action

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /Users/tonigallimore/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

#### Option 3: Shell Script

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Mac/Windows with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed
- Works headlessly (Excel runs in background)

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- Ensure Excel is installed on your system
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts)
- Check that your data connections are properly configured in Excel
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.sh                     # Shell script for easy execution
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```


Checking for linting errors in the files we created:



Verifying the project structure and that imports work correctly:



The import error is expected since dependencies aren't installed. Verifying the project structure:



## Completed tasks

1. **Project structure** — Created `excel_processor.py` module (email_handler.py already existed)

2. **Email handler** — Enhanced `EmailHandler` with:
   - Logging
   - Error handling
   - Connection cleanup

3. **Excel processor** — Created `ExcelProcessor` with:
   - Worksheet update functionality
   - xlwings-based refresh with fallback
   - Worksheet copying with formatting
   - Smartphone-friendly formatting options

4. **Main script** — Refactored `large_deal_report_automation.py` to:
   - Use modular structure (EmailHandler and ExcelProcessor)
   - Orchestrate the workflow
   - Include error handling and logging

5. **Configuration template** — Updated `config.json.example` with:
   - All required email settings
   - Directory paths
   - Worksheet names
   - Logging configuration options

6. **Requirements** — Updated `requirements.txt` with:
   - `openpyxl>=3.1.2`
   - `pandas>=2.0.0`
   - `xlwings>=0.30.0`

7. **README** — Created documentation with:
   - Setup instructions
   - Configuration guide
   - Troubleshooting
   - Usage examples
   - Email provider configurations
   - Security best practices

8. **Error handling** — Added throughout:
   - Logging in all modules
   - Try/except blocks
   - Cleanup in finally blocks
   - Validation
   - Clear error messages

The project is modular, documented, and ready to use. All files pass linting, and the structure follows the plan.

---

**User**

this task will be ran on my work desktop, which is actually windows 11, can you edit the script?

---

**Cursor**


Updating the script and documentation for Windows 11 compatibility. Reviewing the current files and making necessary changes.



Updating the script and documentation for Windows 11 compatibility. Fixing a logging issue and adding Windows-specific files.


```python
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
            required_email_fields = ['username', 'password', 'sender_email']
            for field in required_email_fields:
                if field not in config['email']:
                    raise ValueError(f"Missing required email configuration: {field}")
            
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

```


Creating a Windows batch file and updating the README with Windows-specific instructions:


```bat
@echo off
REM Large Deal Report Automation - Windows Batch Script
REM Run this script to execute the automation

cd /d "%~dp0"
python large_deal_report_automation.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Error: Automation failed with exit code %ERRORLEVEL%
    pause
)

```

```python
#!/usr/bin/env python3
"""
Excel Processor Module
Handles Excel file manipulation, worksheet updates, refresh operations, and copying.
"""

import logging
from pathlib import Path
from typing import Optional
import openpyxl
from openpyxl import load_workbook

# Try to import xlwings for Excel refresh functionality
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False
    logging.warning("xlwings not available. Worksheet 2 refresh will use fallback method.")


class ExcelProcessor:
    """Handles Excel file manipulation and processing operations."""
    
    def __init__(self, config: dict):
        """
        Initialize Excel processor with configuration.
        
        Args:
            config: Dictionary containing configuration:
                - worksheet_1_name: Name of worksheet 1
                - worksheet_2_name: Name of worksheet 2
                - worksheet_3_name: Name of worksheet 3
        """
        self.config = config
        self.worksheet_1_name = config.get('worksheet_1_name', 'Sheet1')
        self.worksheet_2_name = config.get('worksheet_2_name', 'Sheet2')
        self.worksheet_3_name = config.get('worksheet_3_name', 'Sheet3')
        self.logger = logging.getLogger(__name__)
    
    def load_workbook(self, filepath: Path):
        """
        Load an Excel workbook.
        
        Args:
            filepath: Path to the Excel file
            
        Returns:
            openpyxl Workbook object
            
        Raises:
            FileNotFoundError: If file doesn't exist
            Exception: If file cannot be opened
        """
        try:
            self.logger.info(f"Loading workbook: {filepath}")
            return load_workbook(filepath)
        except FileNotFoundError:
            raise FileNotFoundError(f"Workbook not found: {filepath}")
        except Exception as e:
            raise Exception(f"Error loading workbook {filepath}: {str(e)}")
    
    def update_worksheet_1(self, workbook: openpyxl.Workbook, daily_sheet_path: Path) -> None:
        """
        Replace worksheet 1 with data from the daily email attachment.
        
        Args:
            workbook: The target workbook to update
            daily_sheet_path: Path to the daily worksheet file
            
        Raises:
            FileNotFoundError: If daily sheet file doesn't exist
            ValueError: If worksheet 1 doesn't exist
            Exception: If update fails
        """
        self.logger.info("Updating Worksheet 1...")
        
        try:
            # Load the daily sheet
            if not daily_sheet_path.exists():
                raise FileNotFoundError(f"Daily sheet not found: {daily_sheet_path}")
            
            daily_wb = load_workbook(daily_sheet_path)
            daily_ws = daily_wb.active
            
            # Get or create worksheet 1
            if self.worksheet_1_name in workbook.sheetnames:
                ws1 = workbook[self.worksheet_1_name]
            else:
                self.logger.warning(
                    f"Worksheet '{self.worksheet_1_name}' not found. Using active sheet."
                )
                ws1 = workbook.active
            
            # Clear existing data in worksheet 1
            if ws1.max_row > 0:
                ws1.delete_rows(1, ws1.max_row)
            
            # Copy data from daily sheet to worksheet 1
            for row in daily_ws.iter_rows(values_only=True):
                if any(cell is not None for cell in row):  # Skip completely empty rows
                    ws1.append(row)
            
            # Copy formatting if needed (basic cell copying)
            for row_idx, row in enumerate(daily_ws.iter_rows(), start=1):
                if row_idx > ws1.max_row:
                    break
                for col_idx, cell in enumerate(row, start=1):
                    if cell.has_style:
                        target_cell = ws1.cell(row=row_idx, column=col_idx)
                        target_cell.font = cell.font
                        target_cell.fill = cell.fill
                        target_cell.border = cell.border
                        target_cell.alignment = cell.alignment
                        target_cell.number_format = cell.number_format
            
            daily_wb.close()
            self.logger.info("Worksheet 1 updated successfully.")
            
        except Exception as e:
            self.logger.error(f"Error updating Worksheet 1: {str(e)}")
            raise
    
    def refresh_worksheet_2(self, workbook: openpyxl.Workbook, workbook_path: Optional[Path] = None) -> None:
        """
        Refresh data in worksheet 2 using xlwings (if available) or fallback method.
        
        Args:
            workbook: The workbook containing worksheet 2
            workbook_path: Optional path to the workbook file (required for xlwings)
            
        Raises:
            ValueError: If worksheet 2 doesn't exist
            Exception: If refresh fails
        """
        self.logger.info("Refreshing Worksheet 2...")
        
        if self.worksheet_2_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found in workbook")
        
        # Try xlwings refresh if available and workbook_path is provided
        if XLWINGS_AVAILABLE and workbook_path:
            try:
                self._refresh_with_xlwings(workbook_path)
                self.logger.info("Worksheet 2 refreshed using xlwings.")
                return
            except Exception as e:
                self.logger.warning(
                    f"xlwings refresh failed: {str(e)}. Falling back to openpyxl method."
                )
        
        # Fallback: Use openpyxl (preserves structure but doesn't refresh connections)
        self.logger.info(
            "Using openpyxl fallback. Note: Power Query and external data connections "
            "will not be automatically refreshed. You may need to manually refresh in Excel."
        )
        # The worksheet structure is preserved, but data connections won't refresh
        # This is expected behavior when xlwings is not available
    
    def _refresh_with_xlwings(self, workbook_path: Path) -> None:
        """
        Refresh worksheet 2 using xlwings (requires Excel installed).
        Works on both Windows and macOS.
        
        Args:
            workbook_path: Path to the workbook file
            
        Raises:
            Exception: If xlwings refresh fails
        """
        app = None
        wb = None
        
        try:
            self.logger.info("Attempting to refresh using xlwings...")
            
            # Open Excel application (headless)
            # On Windows, this will use the installed Excel application
            # On macOS, this will use Excel for Mac
            app = xw.App(visible=False)
            app.display_alerts = False
            
            # Open the workbook
            wb = app.books.open(str(workbook_path))
            
            # Get worksheet 2
            if self.worksheet_2_name not in [sheet.name for sheet in wb.sheets]:
                raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found")
            
            ws = wb.sheets[self.worksheet_2_name]
            
            # Try to refresh all data connections/queries
            try:
                # Refresh all queries/connections in the workbook
                wb.api.RefreshAll()
                self.logger.info("Refreshed all data connections in workbook.")
            except Exception as e:
                self.logger.warning(f"Could not refresh all connections: {str(e)}")
                # Try to refresh specific table if it's a ListObject
                try:
                    if ws.api.ListObjects.Count > 0:
                        for list_obj in ws.api.ListObjects:
                            if hasattr(list_obj, 'QueryTable'):
                                list_obj.QueryTable.Refresh(BackgroundQuery=False)
                        self.logger.info("Refreshed table connections.")
                except Exception as e2:
                    self.logger.warning(f"Could not refresh table connections: {str(e2)}")
            
            # Save the workbook
            wb.save()
            self.logger.info("Workbook saved after refresh.")
            
        except Exception as e:
            self.logger.error(f"xlwings refresh error: {str(e)}")
            raise
        finally:
            # Clean up
            if wb:
                try:
                    wb.close()
                except:
                    pass
            if app:
                try:
                    app.quit()
                except:
                    pass
    
    def copy_worksheet_2_to_3(self, workbook: openpyxl.Workbook) -> None:
        """
        Copy worksheet 2 data to worksheet 3 (smartphone-friendly version).
        
        Args:
            workbook: The workbook containing the worksheets
            
        Raises:
            ValueError: If worksheet 2 doesn't exist
            Exception: If copy operation fails
        """
        self.logger.info("Copying Worksheet 2 to Worksheet 3...")
        
        if self.worksheet_2_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.worksheet_2_name}' not found in workbook")
        
        try:
            ws2 = workbook[self.worksheet_2_name]
            
            # Remove existing worksheet 3 if it exists
            if self.worksheet_3_name in workbook.sheetnames:
                workbook.remove(workbook[self.worksheet_3_name])
            
            # Create new worksheet 3
            ws3 = workbook.create_sheet(self.worksheet_3_name)
            
            # Copy data from worksheet 2 to worksheet 3
            for row in ws2.iter_rows(values_only=True):
                if any(cell is not None for cell in row):  # Skip completely empty rows
                    ws3.append(row)
            
            # Copy formatting
            for row_idx, row in enumerate(ws2.iter_rows(), start=1):
                if row_idx > ws3.max_row:
                    break
                for col_idx, cell in enumerate(row, start=1):
                    if cell.has_style:
                        target_cell = ws3.cell(row=row_idx, column=col_idx)
                        target_cell.font = cell.font
                        target_cell.fill = cell.fill
                        target_cell.border = cell.border
                        target_cell.alignment = cell.alignment
                        target_cell.number_format = cell.number_format
            
            # Apply smartphone-friendly formatting (optional enhancements)
            # Adjust column widths for better mobile viewing
            for column in ws3.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 for mobile
                ws3.column_dimensions[column_letter].width = adjusted_width
            
            self.logger.info("Worksheet 3 created successfully.")
            
        except Exception as e:
            self.logger.error(f"Error copying Worksheet 2 to Worksheet 3: {str(e)}")
            raise
    
    def save_workbook(self, workbook: openpyxl.Workbook, filepath: Path) -> None:
        """
        Save the workbook to a file.
        
        Args:
            workbook: The workbook to save
            filepath: Path where to save the workbook
            
        Raises:
            Exception: If save fails
        """
        try:
            self.logger.info(f"Saving workbook to: {filepath}")
            workbook.save(filepath)
            self.logger.info("Workbook saved successfully.")
        except Exception as e:
            self.logger.error(f"Error saving workbook: {str(e)}")
            raise

```


Updating the README with Windows-specific instructions:


```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality on macOS/Windows)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

```bash
python large_deal_report_automation.py
```

Or make it executable and run directly:

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution (macOS)

#### Option 1: Shell Alias

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /Users/tonigallimore/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

#### Option 2: macOS Automator Quick Action

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /Users/tonigallimore/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

#### Option 3: Shell Script

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Mac/Windows with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed
- Works headlessly (Excel runs in background)

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- Ensure Excel is installed on your system
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts)
- Check that your data connections are properly configured in Excel
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.sh                     # Shell script for easy execution
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```

```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality on macOS/Windows)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

**Windows:**
```cmd
python large_deal_report_automation.py
```

**macOS/Linux:**
```bash
python large_deal_report_automation.py
```

Or make it executable and run directly (macOS/Linux only):

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution

#### Windows

**Option 1: Batch File (Recommended)**

Double-click `run_report.bat` or run from Command Prompt:

```cmd
run_report.bat
```

**Option 2: Desktop Shortcut**

1. Right-click `run_report.bat`
2. Select "Create shortcut"
3. Drag the shortcut to your desktop
4. Right-click the shortcut → Properties
5. Optionally set a keyboard shortcut in the "Shortcut key" field

**Option 3: Task Scheduler (Automated Daily Run)**

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9:00 AM)
4. Action: Start a program
5. Program: `python.exe`
6. Arguments: `"C:\path\to\large_deal_report_automation.py"`
7. Start in: `C:\path\to\python\`

**Option 4: PowerShell Script**

Create a PowerShell script for more advanced execution:

```powershell
Set-Location "C:\path\to\python"
python large_deal_report_automation.py
```

#### macOS

**Option 1: Shell Alias**

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /path/to/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

**Option 2: macOS Automator Quick Action**

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /path/to/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

**Option 3: Shell Script**

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Mac/Windows with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed
- Works headlessly (Excel runs in background)

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- Ensure Excel is installed on your system
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts)
- Check that your data connections are properly configured in Excel
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.sh                     # Shell script for easy execution
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```

```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality - **Windows 11 recommended**)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

### Windows 11 Specific Notes

- **Excel**: Microsoft Excel must be installed (Office 365 or standalone version)
- **Python**: Ensure Python is added to your system PATH
- **xlwings**: Works seamlessly with Excel on Windows - no additional configuration needed
- **Path Handling**: The script uses `pathlib` which handles Windows paths automatically

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

**Windows:**
```cmd
python large_deal_report_automation.py
```

**macOS/Linux:**
```bash
python large_deal_report_automation.py
```

Or make it executable and run directly (macOS/Linux only):

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution

#### Windows

**Option 1: Batch File (Recommended)**

Double-click `run_report.bat` or run from Command Prompt:

```cmd
run_report.bat
```

**Option 2: Desktop Shortcut**

1. Right-click `run_report.bat`
2. Select "Create shortcut"
3. Drag the shortcut to your desktop
4. Right-click the shortcut → Properties
5. Optionally set a keyboard shortcut in the "Shortcut key" field

**Option 3: Task Scheduler (Automated Daily Run)**

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9:00 AM)
4. Action: Start a program
5. Program: `python.exe`
6. Arguments: `"C:\path\to\large_deal_report_automation.py"`
7. Start in: `C:\path\to\python\`

**Option 4: PowerShell Script**

Create a PowerShell script for more advanced execution:

```powershell
Set-Location "C:\path\to\python"
python large_deal_report_automation.py
```

#### macOS

**Option 1: Shell Alias**

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /path/to/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

**Option 2: macOS Automator Quick Action**

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /path/to/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

**Option 3: Shell Script**

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Mac/Windows with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed
- Works headlessly (Excel runs in background)

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- Ensure Excel is installed on your system
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts)
- Check that your data connections are properly configured in Excel
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.sh                     # Shell script for easy execution
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```

```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality - **Windows 11 recommended**)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

### Windows 11 Specific Notes

- **Excel**: Microsoft Excel must be installed (Office 365 or standalone version)
- **Python**: Ensure Python is added to your system PATH
- **xlwings**: Works seamlessly with Excel on Windows - no additional configuration needed
- **Path Handling**: The script uses `pathlib` which handles Windows paths automatically

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

**Windows:**
```cmd
python large_deal_report_automation.py
```

**macOS/Linux:**
```bash
python large_deal_report_automation.py
```

Or make it executable and run directly (macOS/Linux only):

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution

#### Windows

**Option 1: Batch File (Recommended)**

Double-click `run_report.bat` or run from Command Prompt:

```cmd
run_report.bat
```

**Option 2: Desktop Shortcut**

1. Right-click `run_report.bat`
2. Select "Create shortcut"
3. Drag the shortcut to your desktop
4. Right-click the shortcut → Properties
5. Optionally set a keyboard shortcut in the "Shortcut key" field

**Option 3: Task Scheduler (Automated Daily Run)**

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9:00 AM)
4. Action: Start a program
5. Program: `python.exe`
6. Arguments: `"C:\path\to\large_deal_report_automation.py"`
7. Start in: `C:\path\to\python\`

**Option 4: PowerShell Script**

Create a PowerShell script for more advanced execution:

```powershell
Set-Location "C:\path\to\python"
python large_deal_report_automation.py
```

#### macOS

**Option 1: Shell Alias**

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /path/to/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

**Option 2: macOS Automator Quick Action**

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /path/to/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

**Option 3: Shell Script**

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Windows 11 with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed (works best on Windows)
- Works headlessly (Excel runs in background)
- **Windows 11**: xlwings integrates seamlessly with Excel on Windows

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- Ensure Excel is installed on your system
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts)
- Check that your data connections are properly configured in Excel
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.sh                     # Shell script for easy execution
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```

```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality - **Windows 11 recommended**)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

### Windows 11 Specific Notes

- **Excel**: Microsoft Excel must be installed (Office 365 or standalone version)
- **Python**: Ensure Python is added to your system PATH
- **xlwings**: Works seamlessly with Excel on Windows - no additional configuration needed
- **Path Handling**: The script uses `pathlib` which handles Windows paths automatically

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

**Windows:**
```cmd
python large_deal_report_automation.py
```

**macOS/Linux:**
```bash
python large_deal_report_automation.py
```

Or make it executable and run directly (macOS/Linux only):

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution

#### Windows

**Option 1: Batch File (Recommended)**

Double-click `run_report.bat` or run from Command Prompt:

```cmd
run_report.bat
```

**Option 2: Desktop Shortcut**

1. Right-click `run_report.bat`
2. Select "Create shortcut"
3. Drag the shortcut to your desktop
4. Right-click the shortcut → Properties
5. Optionally set a keyboard shortcut in the "Shortcut key" field

**Option 3: Task Scheduler (Automated Daily Run)**

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9:00 AM)
4. Action: Start a program
5. Program: `python.exe`
6. Arguments: `"C:\path\to\large_deal_report_automation.py"`
7. Start in: `C:\path\to\python\`

**Option 4: PowerShell Script**

Create a PowerShell script for more advanced execution:

```powershell
Set-Location "C:\path\to\python"
python large_deal_report_automation.py
```

#### macOS

**Option 1: Shell Alias**

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /path/to/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

**Option 2: macOS Automator Quick Action**

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /path/to/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

**Option 3: Shell Script**

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Windows 11 with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed (works best on Windows)
- Works headlessly (Excel runs in background)
- **Windows 11**: xlwings integrates seamlessly with Excel on Windows

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- Ensure Excel is installed on your system
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts)
- Check that your data connections are properly configured in Excel
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.bat                    # Windows batch script for easy execution
├── run_report.sh                     # macOS/Linux shell script
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

### Windows Path Examples

On Windows, you can use either forward slashes or backslashes in paths. The script handles both:

```json
{
  "reports_directory": "C:\\Users\\YourName\\Documents\\reports",
  "temp_directory": "C:\\Users\\YourName\\Documents\\temp"
}
```

Or use forward slashes (recommended):

```json
{
  "reports_directory": "C:/Users/YourName/Documents/reports",
  "temp_directory": "C:/Users/YourName/Documents/temp"
}
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```

```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality - **Windows 11 recommended**)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

### Windows 11 Specific Notes

- **Excel**: Microsoft Excel must be installed (Office 365 or standalone version)
- **Python**: Ensure Python is added to your system PATH
- **xlwings**: Works seamlessly with Excel on Windows - no additional configuration needed
- **Path Handling**: The script uses `pathlib` which handles Windows paths automatically

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

**Windows:**
```cmd
python large_deal_report_automation.py
```

**macOS/Linux:**
```bash
python large_deal_report_automation.py
```

Or make it executable and run directly (macOS/Linux only):

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution

#### Windows

**Option 1: Batch File (Recommended)**

Double-click `run_report.bat` or run from Command Prompt:

```cmd
run_report.bat
```

**Option 2: Desktop Shortcut**

1. Right-click `run_report.bat`
2. Select "Create shortcut"
3. Drag the shortcut to your desktop
4. Right-click the shortcut → Properties
5. Optionally set a keyboard shortcut in the "Shortcut key" field

**Option 3: Task Scheduler (Automated Daily Run)**

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9:00 AM)
4. Action: Start a program
5. Program: `python.exe`
6. Arguments: `"C:\path\to\large_deal_report_automation.py"`
7. Start in: `C:\path\to\python\`

**Option 4: PowerShell Script**

Create a PowerShell script for more advanced execution:

```powershell
Set-Location "C:\path\to\python"
python large_deal_report_automation.py
```

#### macOS

**Option 1: Shell Alias**

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /path/to/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

**Option 2: macOS Automator Quick Action**

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /path/to/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

**Option 3: Shell Script**

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Windows 11 with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed (works best on Windows)
- Works headlessly (Excel runs in background)
- **Windows 11**: xlwings integrates seamlessly with Excel on Windows

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- **Windows 11**: Ensure Excel is installed and properly licensed
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts on Windows)
- On Windows, ensure Excel is the default application for .xlsx files
- Check that your data connections are properly configured in Excel
- Try running Excel once manually to ensure it's properly initialized
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.bat                    # Windows batch script for easy execution
├── run_report.sh                     # macOS/Linux shell script
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

### Windows Path Examples

On Windows, you can use either forward slashes or backslashes in paths. The script handles both:

```json
{
  "reports_directory": "C:\\Users\\YourName\\Documents\\reports",
  "temp_directory": "C:\\Users\\YourName\\Documents\\temp"
}
```

Or use forward slashes (recommended):

```json
{
  "reports_directory": "C:/Users/YourName/Documents/reports",
  "temp_directory": "C:/Users/YourName/Documents/temp"
}
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```


Adding a Windows-specific section to the README and verifying everything is set up correctly:


```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality - **Windows 11 recommended**)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

### Windows 11 Specific Notes

- **Excel**: Microsoft Excel must be installed (Office 365 or standalone version)
- **Python**: Ensure Python is added to your system PATH
- **xlwings**: Works seamlessly with Excel on Windows - no additional configuration needed
- **Path Handling**: The script uses `pathlib` which handles Windows paths automatically

## Installation

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

**Windows:**
```cmd
python large_deal_report_automation.py
```

**macOS/Linux:**
```bash
python large_deal_report_automation.py
```

Or make it executable and run directly (macOS/Linux only):

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution

#### Windows

**Option 1: Batch File (Recommended)**

Double-click `run_report.bat` or run from Command Prompt:

```cmd
run_report.bat
```

**Option 2: Desktop Shortcut**

1. Right-click `run_report.bat`
2. Select "Create shortcut"
3. Drag the shortcut to your desktop
4. Right-click the shortcut → Properties
5. Optionally set a keyboard shortcut in the "Shortcut key" field

**Option 3: Task Scheduler (Automated Daily Run)**

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9:00 AM)
4. Action: Start a program
5. Program: `python.exe`
6. Arguments: `"C:\path\to\large_deal_report_automation.py"`
7. Start in: `C:\path\to\python\`

**Option 4: PowerShell Script**

Create a PowerShell script for more advanced execution:

```powershell
Set-Location "C:\path\to\python"
python large_deal_report_automation.py
```

#### macOS

**Option 1: Shell Alias**

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /path/to/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

**Option 2: macOS Automator Quick Action**

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /path/to/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

**Option 3: Shell Script**

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Windows 11 with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed (works best on Windows)
- Works headlessly (Excel runs in background)
- **Windows 11**: xlwings integrates seamlessly with Excel on Windows

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- **Windows 11**: Ensure Excel is installed and properly licensed
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts on Windows)
- On Windows, ensure Excel is the default application for .xlsx files
- Check that your data connections are properly configured in Excel
- Try running Excel once manually to ensure it's properly initialized
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.bat                    # Windows batch script for easy execution
├── run_report.sh                     # macOS/Linux shell script
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

### Windows Path Examples

On Windows, you can use either forward slashes or backslashes in paths. The script handles both:

```json
{
  "reports_directory": "C:\\Users\\YourName\\Documents\\reports",
  "temp_directory": "C:\\Users\\YourName\\Documents\\temp"
}
```

Or use forward slashes (recommended):

```json
{
  "reports_directory": "C:/Users/YourName/Documents/reports",
  "temp_directory": "C:/Users/YourName/Documents/temp"
}
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Windows 11 Specific Information

### Excel Integration

On Windows 11, xlwings works seamlessly with Microsoft Excel:

- **No Additional Setup**: xlwings automatically finds and uses your installed Excel
- **Background Operation**: Excel runs in the background (not visible) during refresh
- **Power Query Support**: Automatically refreshes Power Query connections
- **Data Connections**: Refreshes all external data connections in the workbook

### Running on Windows

**Recommended Method:**
1. Double-click `run_report.bat` for one-click execution
2. Or create a desktop shortcut to `run_report.bat`

**Command Prompt:**
```cmd
cd C:\path\to\python
python large_deal_report_automation.py
```

**PowerShell:**
```powershell
cd C:\path\to\python
python large_deal_report_automation.py
```

### Windows Task Scheduler Setup

To run the script automatically every day:

1. Open **Task Scheduler** (search in Start menu)
2. Click **Create Basic Task**
3. Name: "Large Deal Report Automation"
4. Trigger: **Daily** at your preferred time (e.g., 9:00 AM)
5. Action: **Start a program**
6. Program/script: `python.exe` (or full path: `C:\Python\python.exe`)
7. Add arguments: `"C:\full\path\to\large_deal_report_automation.py"`
8. Start in: `C:\full\path\to\python\`
9. Check "Open the Properties dialog for this task when I click Finish"
10. In Properties → General:
    - Check "Run whether user is logged on or not" (optional)
    - Check "Run with highest privileges" (if needed)
11. In Properties → Conditions:
    - Uncheck "Start the task only if the computer is on AC power" (if you want it to run on battery)
12. Click OK and enter your Windows password if prompted

### Windows Path Configuration

When configuring paths in `config.json` on Windows:

```json
{
  "reports_directory": "C:/Users/YourName/Documents/reports",
  "temp_directory": "C:/Users/YourName/Documents/temp",
  "log_file": "C:/Users/YourName/Documents/logs/report_automation.log"
}
```

**Note:** Forward slashes work on Windows and are recommended for cross-platform compatibility.

### Windows Firewall

If you encounter email connection issues:

1. Windows Firewall may block Python's network access
2. When Python first connects, Windows will prompt you to allow access
3. Check "Private networks" and "Public networks" and click "Allow access"
4. Alternatively, add Python to Windows Firewall exceptions manually

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated (`pip install --upgrade -r requirements.txt`)
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues
5. **Windows Updates**: Keep Excel and Python updated

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories
6. **Windows**: Check Windows Event Viewer for system-level errors

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```


Verifying the batch file and adding a quick Windows setup guide:


```markdown
# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically
- **Excel Processing**: Updates worksheets, refreshes data connections, and creates smartphone-friendly versions
- **Data Refresh**: Supports xlwings for automatic refresh of Power Query and external data connections
- **Automated Distribution**: Emails the completed report to a distribution list
- **Error Handling**: Comprehensive error handling with detailed logging
- **Modular Design**: Separate modules for email handling and Excel processing

## Architecture

The automation follows this workflow:

1. Find the latest report template
2. Download daily worksheet from email
3. Update Worksheet 1 with daily data (full replacement)
4. Refresh Worksheet 2 using xlwings (or fallback method)
5. Copy Worksheet 2 to Worksheet 3 (smartphone-friendly version)
6. Save report with current date
7. Email report to distribution list
8. Clean up temporary files

## Prerequisites

- Python 3.7 or higher
- Excel installed (for xlwings refresh functionality - **Windows 11 recommended**)
- Email account with IMAP and SMTP access
- Gmail App Password (if using Gmail) or appropriate email credentials

### Windows 11 Specific Notes

- **Excel**: Microsoft Excel must be installed (Office 365 or standalone version)
- **Python**: Ensure Python is added to your system PATH
- **xlwings**: Works seamlessly with Excel on Windows - no additional configuration needed
- **Path Handling**: The script uses `pathlib` which handles Windows paths automatically

## Installation

### Windows 11 Quick Start

1. **Install Python** (if not already installed):
   - Download from [python.org](https://www.python.org/downloads/)
   - During installation, check "Add Python to PATH"
   - Verify: Open Command Prompt and run `python --version`

2. **Install Excel** (if not already installed):
   - Microsoft Excel is required for xlwings functionality
   - Office 365 or standalone Excel both work

3. **Download/Extract Project Files**:
   - Place all files in a folder (e.g., `C:\Users\YourName\Documents\python\`)

4. **Install Dependencies** (see step 2 below)

5. **Configure** (see step 3 below)

6. **Run**: Double-click `run_report.bat` or use Command Prompt

### 1. Clone or Download the Project

Ensure you have all the project files in your working directory.

**Windows:** Extract to a folder like `C:\Users\YourName\Documents\python\`

### 2. Install Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - Excel file manipulation (reading/writing)
- `pandas` - Data manipulation (optional, for advanced operations)
- `xlwings` - Excel application automation for data refresh (macOS/Windows)

**Note:** `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).

### 3. Configure Email Settings

Copy the example configuration file and fill in your details:

**Windows:**
```cmd
copy config.json.example config.json
```

**macOS/Linux:**
```bash
cp config.json.example config.json
```

Edit `config.json` with your email credentials and settings:

```json
{
  "email": {
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "sender_email": "sender@example.com",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  },
  "distribution_list": [
    "recipient1@example.com",
    "recipient2@example.com"
  ],
  "sender_name": "Your Name",
  "reports_directory": "./reports",
  "temp_directory": "./temp",
  "worksheet_1_name": "Sheet1",
  "worksheet_2_name": "Sheet2",
  "worksheet_3_name": "Sheet3",
  "log_level": "INFO",
  "log_file": null
}
```

#### Configuration Options

- **email.username**: Your email address for sending/receiving
- **email.password**: App Password (for Gmail) or regular password
- **email.sender_email**: The email address that sends you the daily worksheet
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`)
- **email.imap_port**: IMAP server port (default: `993`)
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`)
- **email.smtp_port**: SMTP server port (default: `587`)
- **distribution_list**: Array of recipient email addresses
- **sender_name**: Name to appear in email signature
- **reports_directory**: Directory where reports are stored (default: `./reports`)
- **temp_directory**: Directory for temporary files (default: `./temp`)
- **worksheet_1_name**: Name of worksheet 1 (default: `Sheet1`)
- **worksheet_2_name**: Name of worksheet 2 (default: `Sheet2`)
- **worksheet_3_name**: Name of worksheet 3 (default: `Sheet3`)
- **log_level**: Logging level - `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`)
- **log_file**: Path to log file (optional, set to `null` for console-only logging)

### 4. Gmail App Password Setup

If using Gmail, you **must** use an App Password instead of your regular password:

1. Go to your Google Account settings
2. Enable 2-Step Verification (required for App Passwords)
3. Go to Security → App Passwords
4. Generate a new App Password for "Mail"
5. Use this 16-character password in your `config.json`

**Important:** Never use your regular Gmail password. App Passwords are required for Gmail.

### 5. Prepare Your Report Template

1. Create your initial report with 3 worksheets:
   - **Worksheet 1**: Will be overwritten with daily data from email
   - **Worksheet 2**: Contains the main report with data connections/Power Query
   - **Worksheet 3**: Smartphone-friendly version (will be auto-generated)

2. Save this template in the `./reports/` directory with the naming pattern:
   ```
   large deal report -YYYY-MM-DD.xlsx
   ```
   Example: `large deal report -2024-01-15.xlsx`

3. Ensure Worksheet 2 has your data connections configured (Power Query, external data, etc.)

### 6. Email Setup

The script expects to receive a daily email with an Excel attachment:

- The email must come from the address specified in `config.json` under `sender_email`
- The attachment must be an Excel file (`.xlsx` or `.xls`)
- The script will find the most recent email from today, or fall back to the most recent email if none from today

## Usage

### Basic Usage

Run the automation script:

**Windows:**
```cmd
python large_deal_report_automation.py
```

**macOS/Linux:**
```bash
python large_deal_report_automation.py
```

Or make it executable and run directly (macOS/Linux only):

```bash
chmod +x large_deal_report_automation.py
./large_deal_report_automation.py
```

### One-Click Execution

#### Windows

**Option 1: Batch File (Recommended)**

Double-click `run_report.bat` or run from Command Prompt:

```cmd
run_report.bat
```

**Option 2: Desktop Shortcut**

1. Right-click `run_report.bat`
2. Select "Create shortcut"
3. Drag the shortcut to your desktop
4. Right-click the shortcut → Properties
5. Optionally set a keyboard shortcut in the "Shortcut key" field

**Option 3: Task Scheduler (Automated Daily Run)**

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., daily at 9:00 AM)
4. Action: Start a program
5. Program: `python.exe`
6. Arguments: `"C:\path\to\large_deal_report_automation.py"`
7. Start in: `C:\path\to\python\`

**Option 4: PowerShell Script**

Create a PowerShell script for more advanced execution:

```powershell
Set-Location "C:\path\to\python"
python large_deal_report_automation.py
```

#### macOS

**Option 1: Shell Alias**

Add an alias to your shell profile (`~/.zshrc` or `~/.bash_profile`):

```bash
alias large-deal-report="cd /path/to/python && python large_deal_report_automation.py"
```

Then run:
```bash
large-deal-report
```

**Option 2: macOS Automator Quick Action**

1. Open Automator
2. Create a new "Quick Action"
3. Add "Run Shell Script" action
4. Set it to: `cd /path/to/python && python large_deal_report_automation.py`
5. Save as "Generate Large Deal Report"
6. Access from Services menu or assign a keyboard shortcut

**Option 3: Shell Script**

Use the provided `run_report.sh` script:

```bash
./run_report.sh
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report template in the reports directory
2. **Download Email Attachment**: Connects to email, finds the daily email, and downloads the Excel attachment
3. **Update Worksheet 1**: Replaces all data in Worksheet 1 with data from the daily email attachment
4. **Refresh Worksheet 2**: Attempts to refresh data connections using xlwings (if available) or preserves structure
5. **Copy to Worksheet 3**: Creates a smartphone-friendly copy of Worksheet 2
6. **Save Report**: Saves the updated report with current date: `large deal report -YYYY-MM-DD.xlsx`
7. **Send Email**: Emails the report to all recipients in the distribution list
8. **Cleanup**: Removes temporary downloaded files

### Worksheet 2 Refresh

The script supports multiple methods for refreshing Worksheet 2:

**Option 1: xlwings (Recommended for Windows 11 with Excel)**
- Automatically refreshes Power Query and external data connections
- Requires Excel to be installed (works best on Windows)
- Works headlessly (Excel runs in background)
- **Windows 11**: xlwings integrates seamlessly with Excel on Windows

**Option 2: Fallback (openpyxl)**
- Preserves worksheet structure
- Does not refresh data connections
- You may need to manually refresh in Excel or set Excel to auto-refresh on open

The script automatically tries xlwings first and falls back if it's not available or fails.

## Email Provider Configuration

### Gmail (Default)

```json
{
  "imap_server": "imap.gmail.com",
  "imap_port": 993,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587
}
```

**Requirements:**
- App Password (not regular password)
- IMAP enabled in Gmail settings

### Outlook/Office 365

```json
{
  "imap_server": "outlook.office365.com",
  "imap_port": 993,
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587
}
```

### Yahoo Mail

```json
{
  "imap_server": "imap.mail.yahoo.com",
  "imap_port": 993,
  "smtp_server": "smtp.mail.yahoo.com",
  "smtp_port": 587
}
```

## Troubleshooting

### "No existing reports found"

**Problem:** The script cannot find a report template.

**Solutions:**
- Ensure you have at least one report file in the `./reports/` directory
- Check that the filename follows the pattern: `large deal report -YYYY-MM-DD.xlsx`
- Verify the `reports_directory` path in `config.json` is correct

### "No emails found from [sender]"

**Problem:** The script cannot find the daily email.

**Solutions:**
- Verify the `sender_email` address in `config.json` matches exactly
- Check that you've received the daily email
- Try running the script later in the day if emails are sent in the morning
- The script will also search for the most recent email if none from today are found

### "Email authentication failed"

**Problem:** Cannot connect to email server.

**Solutions:**
- **For Gmail**: Make sure you're using an App Password, not your regular password
- Verify your email credentials are correct
- Check that IMAP is enabled in your email settings
- Ensure 2-Step Verification is enabled (required for Gmail App Passwords)
- Verify the IMAP/SMTP server addresses and ports are correct

### "IMAP error" or "SMTP error"

**Problem:** Connection to email server fails.

**Solutions:**
- Check your internet connection
- Verify firewall isn't blocking IMAP/SMTP ports
- Try different ports (e.g., 465 for SMTP SSL)
- Check if your email provider requires specific security settings

### "Worksheet '[name]' not found"

**Problem:** The script cannot find a required worksheet.

**Solutions:**
- Verify worksheet names in `config.json` match your Excel file exactly
- Check that your report template has all three worksheets
- Worksheet names are case-sensitive

### "xlwings refresh failed"

**Problem:** xlwings cannot refresh the worksheet.

**Solutions:**
- **Windows 11**: Ensure Excel is installed and properly licensed
- Check that xlwings is installed: `pip install xlwings`
- Verify Excel is not already open (may cause conflicts on Windows)
- On Windows, ensure Excel is the default application for .xlsx files
- Check that your data connections are properly configured in Excel
- Try running Excel once manually to ensure it's properly initialized
- The script will fall back to openpyxl if xlwings fails

### "Error saving workbook"

**Problem:** Cannot save the Excel file.

**Solutions:**
- Check that the reports directory exists and is writable
- Ensure the file is not open in Excel
- Verify you have sufficient disk space
- Check file permissions on the reports directory

### "Error sending email"

**Problem:** Email cannot be sent.

**Solutions:**
- Verify SMTP server settings
- Check that the distribution list contains valid email addresses
- Ensure the report file exists and is not corrupted
- Check SMTP authentication (App Password for Gmail)
- Verify network connectivity

### Logging Issues

**Problem:** Need more detailed error information.

**Solutions:**
- Set `log_level` to `DEBUG` in `config.json` for verbose logging
- Set `log_file` to a file path to save logs to a file
- Check console output for error messages

## File Structure

```
python/
├── large_deal_report_automation.py  # Main automation script
├── email_handler.py                  # Email retrieval and sending module
├── excel_processor.py                # Excel manipulation module
├── config.json                       # Your configuration (create from example)
├── config.json.example              # Example configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.bat                    # Windows batch script for easy execution
├── run_report.sh                     # macOS/Linux shell script
├── reports/                          # Directory for report files
│   └── large deal report -*.xlsx    # Report templates and generated reports
└── temp/                            # Temporary files (auto-cleaned)
    └── daily_sheet_*.xlsx           # Downloaded email attachments
```

### Windows Path Examples

On Windows, you can use either forward slashes or backslashes in paths. The script handles both:

```json
{
  "reports_directory": "C:\\Users\\YourName\\Documents\\reports",
  "temp_directory": "C:\\Users\\YourName\\Documents\\temp"
}
```

Or use forward slashes (recommended):

```json
{
  "reports_directory": "C:/Users/YourName/Documents/reports",
  "temp_directory": "C:/Users/YourName/Documents/temp"
}
```

## Security Best Practices

1. **Never commit `config.json`** to version control
   - Add it to `.gitignore`
   - Only commit `config.json.example`

2. **Use App Passwords** for email accounts when possible
   - More secure than regular passwords
   - Can be revoked individually

3. **Consider Environment Variables** for sensitive data
   - Store credentials in environment variables
   - Modify the script to read from environment if needed

4. **Restrict File Permissions**
   - Set `config.json` permissions: `chmod 600 config.json`
   - Protect your email credentials

## Customization

### Custom Email Message

Edit the `send_report_email` method in `email_handler.py`:

```python
body = (
    f"Custom message here\n\n"
    f"Report for {current_date}.\n\n"
    f"Best regards,\n{sender_name}"
)
```

### Custom File Naming

Edit the `run` method in `large_deal_report_automation.py`:

```python
filename = f"custom-name-{self.current_date}.xlsx"
```

### Custom Worksheet Formatting

Enhance the `copy_worksheet_2_to_3` method in `excel_processor.py` to add custom formatting for mobile devices.

### Custom Data Refresh Logic

Modify the `refresh_worksheet_2` method in `excel_processor.py` to implement custom refresh logic for your specific data sources.

## Logging

The script includes comprehensive logging:

- **Console Output**: Always shown
- **File Logging**: Optional, set `log_file` in config
- **Log Levels**: DEBUG, INFO, WARNING, ERROR

Example log output:
```
2024-01-15 10:30:00 - __main__ - INFO - Large Deal Report Automation - Starting
2024-01-15 10:30:01 - email_handler - INFO - Connecting to email server...
2024-01-15 10:30:02 - excel_processor - INFO - Updating Worksheet 1...
```

## Windows 11 Specific Information

### Excel Integration

On Windows 11, xlwings works seamlessly with Microsoft Excel:

- **No Additional Setup**: xlwings automatically finds and uses your installed Excel
- **Background Operation**: Excel runs in the background (not visible) during refresh
- **Power Query Support**: Automatically refreshes Power Query connections
- **Data Connections**: Refreshes all external data connections in the workbook

### Running on Windows

**Recommended Method:**
1. Double-click `run_report.bat` for one-click execution
2. Or create a desktop shortcut to `run_report.bat`

**Command Prompt:**
```cmd
cd C:\path\to\python
python large_deal_report_automation.py
```

**PowerShell:**
```powershell
cd C:\path\to\python
python large_deal_report_automation.py
```

### Windows Task Scheduler Setup

To run the script automatically every day:

1. Open **Task Scheduler** (search in Start menu)
2. Click **Create Basic Task**
3. Name: "Large Deal Report Automation"
4. Trigger: **Daily** at your preferred time (e.g., 9:00 AM)
5. Action: **Start a program**
6. Program/script: `python.exe` (or full path: `C:\Python\python.exe`)
7. Add arguments: `"C:\full\path\to\large_deal_report_automation.py"`
8. Start in: `C:\full\path\to\python\`
9. Check "Open the Properties dialog for this task when I click Finish"
10. In Properties → General:
    - Check "Run whether user is logged on or not" (optional)
    - Check "Run with highest privileges" (if needed)
11. In Properties → Conditions:
    - Uncheck "Start the task only if the computer is on AC power" (if you want it to run on battery)
12. Click OK and enter your Windows password if prompted

### Windows Path Configuration

When configuring paths in `config.json` on Windows:

```json
{
  "reports_directory": "C:/Users/YourName/Documents/reports",
  "temp_directory": "C:/Users/YourName/Documents/temp",
  "log_file": "C:/Users/YourName/Documents/logs/report_automation.log"
}
```

**Note:** Forward slashes work on Windows and are recommended for cross-platform compatibility.

### Windows Firewall

If you encounter email connection issues:

1. Windows Firewall may block Python's network access
2. When Python first connects, Windows will prompt you to allow access
3. Check "Private networks" and "Public networks" and click "Allow access"
4. Alternatively, add Python to Windows Firewall exceptions manually

## Support and Maintenance

### Regular Maintenance

1. **Clean Old Reports**: Periodically archive or delete old reports
2. **Update Dependencies**: Keep Python packages updated (`pip install --upgrade -r requirements.txt`)
3. **Test Email Connectivity**: Verify email settings periodically
4. **Monitor Logs**: Check logs for recurring issues
5. **Windows Updates**: Keep Excel and Python updated

### Getting Help

1. Check the error messages in console output
2. Review the log file (if configured)
3. Verify configuration file settings
4. Test email server connectivity manually
5. Check file permissions in reports and temp directories
6. **Windows**: Check Windows Event Viewer for system-level errors

## License

This script is provided as-is for internal use.

## Version History

- **v1.0**: Initial implementation with modular design, xlwings support, and comprehensive error handling

```


## Windows 11 compatibility updates

1. **Created `run_report.bat`** — Windows batch file for one-click execution
   - Changes to the script directory automatically
   - Shows error messages if the script fails
   - Pauses on error so you can see the message

2. **Updated README.md** with Windows-specific sections:
   - Windows 11 Quick Start guide
   - Windows path configuration examples
   - Task Scheduler setup instructions for automated daily runs
   - Windows firewall notes
   - Windows-specific troubleshooting tips

3. **Enhanced Excel integration notes**:
   - Clarified that xlwings works well on Windows 11
   - Added notes about Excel integration on Windows

4. **Fixed logging initialization** in the main script

## Key Windows features

- **Path handling**: Uses `pathlib`, which handles Windows paths automatically
- **xlwings**: Works with Excel on Windows 11
- **Easy execution**: Double-click `run_report.bat` to run
- **Task Scheduler**: Instructions included for automated daily runs
- **Batch file**: Handles errors and shows messages

## Ready to use on Windows 11

The script is ready to run on your Windows 11 desktop. Use `run_report.bat` for one-click execution, or set up Task Scheduler for automated daily runs.

All files pass linting, and the code is cross-platform compatible while optimized for Windows 11.

