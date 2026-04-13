# Reusable Python Automation Snippets

This handout has two goals:

1. Confirm the presentation code snippets align with the actual Large Deal automation code.
2. Provide small, reusable snippets your team can copy for other automation tasks.

---

## Part A) Snippet Accuracy Check (Presentation vs Actual Code)

### 1) One-click launcher snippet
**Presentation snippet**

```bash
#!/bin/bash
cd "$(dirname "$0")"
python3 large_deal_report_automation.py
```

**Status:** Exact match with `run_report.sh` lines 1-5.

---

### 2) Step-by-step workflow snippet
**Presentation snippet (condensed)**

```python
latest_report_path = self.find_latest_report()
daily_sheet_path = self.email_handler.download_daily_attachment(self.temp_dir, self.current_date)
self.excel_processor.update_worksheet_1(workbook, daily_sheet_path)
self.excel_processor.refresh_worksheet_2(workbook, latest_report_path)
self.excel_processor.copy_worksheet_2_to_3(workbook)
self.excel_processor.save_workbook(workbook, new_report_path)
self.email_handler.send_report_email(new_report_path, self.config["distribution_list"], sender_name, self.current_date)
self.cleanup_temp_files()
```

**Status:** Accurate sequence from `large_deal_report_automation.py` run flow (lines 206-253), intentionally condensed for presentation readability.

---

### 3) Manual copy/paste replaced snippet
**Presentation snippet**

```python
if ws1.max_row > 0:
    ws1.delete_rows(1, ws1.max_row)

for row in daily_ws.iter_rows(values_only=True):
    if any(cell is not None for cell in row):
        ws1.append(row)
```

**Status:** Exact logic match with `excel_processor.py` lines 95-103.

---

## Part B) Reusable Mini-Snippets for Other Scenarios

Use these as copy/paste starters for other team automations.

### Snippet 1: Find latest file by pattern
**Use for:** latest report, latest export, latest drop file.

```python
from pathlib import Path
import glob
import os


def find_latest_file(directory: str, pattern: str) -> Path:
    files = glob.glob(str(Path(directory) / pattern))
    if not files:
        raise FileNotFoundError(f"No files found for pattern: {pattern}")
    return Path(max(files, key=os.path.getmtime))
```

---

### Snippet 2: Safe config loading with required keys
**Use for:** any automation with environment/config settings.

```python
import json


def load_config(config_path: str, required_keys: list[str]) -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
    missing = [k for k in required_keys if k not in config]
    if missing:
        raise ValueError(f"Missing required config keys: {missing}")
    return config
```

---

### Snippet 3: Clean temporary files by glob
**Use for:** temp folders, old exports, staged attachments.

```python
from pathlib import Path


def cleanup_files(directory: str, glob_pattern: str) -> int:
    deleted = 0
    for path in Path(directory).glob(glob_pattern):
        path.unlink(missing_ok=True)
        deleted += 1
    return deleted
```

---

### Snippet 4: IMAP search with fallback
**Use for:** pull newest attachment from known sender.

```python
import imaplib
from datetime import datetime


def search_email_ids(mail: imaplib.IMAP4_SSL, sender_email: str) -> list[bytes]:
    today = datetime.now().strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'(FROM "{sender_email}" SINCE {today})')
    if status != "OK":
        raise RuntimeError("Email search failed")

    email_ids = messages[0].split()
    if not email_ids:
        status, messages = mail.search(None, f'(FROM "{sender_email}")')
        if status != "OK":
            raise RuntimeError("Fallback email search failed")
        email_ids = messages[0].split()

    if not email_ids:
        raise FileNotFoundError(f"No emails found from {sender_email}")
    return email_ids
```

---

### Snippet 5: Replace sheet with fresh rows (Excel)
**Use for:** daily overwrite tabs, flat data refresh.

```python
from openpyxl import load_workbook


def replace_sheet_rows(target_file: str, source_file: str, sheet_name: str) -> None:
    target_wb = load_workbook(target_file)
    source_wb = load_workbook(source_file)

    source_ws = source_wb.active
    ws = target_wb[sheet_name]

    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)

    for row in source_ws.iter_rows(values_only=True):
        if any(cell is not None for cell in row):
            ws.append(row)

    target_wb.save(target_file)
    source_wb.close()
    target_wb.close()
```

---

### Snippet 6: Send file via SMTP
**Use for:** distribute generated reports or exports.

```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def send_email_with_attachment(
    smtp_server: str,
    smtp_port: int,
    username: str,
    password: str,
    recipients: list[str],
    subject: str,
    body: str,
    attachment_path: str,
) -> None:
    msg = MIMEMultipart()
    msg["From"] = username
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with open(attachment_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={attachment_path.split('/')[-1]}")
    msg.attach(part)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(username, password)
        server.sendmail(username, recipients, msg.as_string())
```

---

## Part C) Reuse Pattern to Teach the Team

For each new task, reuse this structure:

1. **Input** (email/file/system)
2. **Process** (clean/transform/refresh)
3. **Output** (file/report/dashboard)
4. **Notify** (who gets it)
5. **Log** (what happened)

If teammates can describe a task in this format, it is usually automation-ready.
