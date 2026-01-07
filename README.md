# Large Deal Report Automation

This Python script automates the daily generation and distribution of the Large Deal Report with a single click. It handles email retrieval, Excel manipulation, data refresh, and distribution automatically.

## Features

- **Automated Email Retrieval**: Downloads daily worksheet from email automatically (filters by subject "large trade td").
- **Excel Processing**: Updates worksheets, refreshes data connections using Windows COM automation.
- **Data Refresh**: Automatically refreshes all data connections/tables in the workbook (Sheet 2).
- **Values-Only Paste**: Creates a "iphone compatible" version by pasting values only.
- **Email Preview**: Creates a draft email in Outlook and displays it for review before sending.
- **Robustness**: Uses standard Windows automation (`pywin32`) for reliability on Windows 11.

## Architecture

The automation follows this workflow:

1. Download daily "large trade td" worksheet from email.
2. Open Master Workbook (defined in config).
3. Copy daily worksheet data to "large deal report" tab (Cell A1).
4. Refresh "summary" tab (Data Connections/Pivot Tables).
5. Copy "summary" table to "iphone compatible" tab (Values Only).
6. Save report as `large deal report - {today's date}.xlsx`.
7. Create Outlook draft with attachment and display it.

## Prerequisites

- **OS**: Windows 11 (or Windows 10)
- **Software**: Microsoft Excel and Microsoft Outlook (Classic Desktop App)
- **Python**: Python 3.7+

## Installation

### 1. Install Python

Download and install Python from [python.org](https://www.python.org/downloads/). Ensure you check "Add Python to PATH" during installation.

### 2. Install Dependencies

Open Command Prompt or PowerShell in the project folder and run:

```bash
pip install -r requirements.txt
```

(Note: `pywin32` is the key dependency for this script on Windows).

### 3. Configure Settings

Copy the example configuration file:

```cmd
copy config.json.example config.json
```

Edit `config.json` with your details:

```json
{
    "email": {
        "username": "your_email@example.com",
        "password": "your_password_or_app_password",
        "sender_email": "sender@example.com",
        "imap_server": "outlook.office365.com",
        "target_subject_keyword": "large trade td"
    },
    "excel": {
        "master_workbook_path": "C:/path/to/your/Master_Template.xlsx",
        "worksheet_1_name": "large deal report",
        "worksheet_2_name": "summary",
        "worksheet_3_name": "iphone compatible"
    },
    "distribution_list": [
        "person1@example.com",
        "person2@example.com"
    ]
}
```

**Important**: `master_workbook_path` should point to the workbook you use as the base every day (the one that has the queries set up).

## Usage

### One-Click Run

Double-click `run_report.bat`.

The script will:
1. Open a console window showing progress.
2. Automate Excel in the background.
3. Pop up an Outlook email window with the report attached.
4. You can then review and click Send.

## Troubleshooting

- **"win32com not available"**: Run `pip install pywin32`.
- **Excel errors**: Ensure the Master Workbook exists and isn't open when you run the script.
- **Email errors**: Check your IMAP settings in `config.json`. If using Gmail, use an App Password.
- **"large trade td" email not found**: Ensure you have received the email today.

