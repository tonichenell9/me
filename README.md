# Large Deal Report Automation

**One-click automation for the daily Large Deal Report workflow on Windows 11.**

This Python script automates your entire daily reporting process:
1. Downloads the daily worksheet from your "large trade td" email
2. Pastes the data into your workbook
3. Refreshes the summary table
4. Creates the iPhone compatible version (values only)
5. Saves the report with today's date
6. Opens an email preview for you to review before sending

## Quick Start (Windows 11)

### Prerequisites
- Windows 11 with Microsoft Excel installed
- Microsoft Outlook (logged into your email account)
- Python 3.8 or higher

### Installation

1. **Install Python** (if not already installed):
   - Download from [python.org](https://www.python.org/downloads/)
   - **Important:** Check "Add Python to PATH" during installation

2. **Install Dependencies**:
   Open Command Prompt and run:
   ```cmd
   cd path\to\this\folder
   pip install -r requirements.txt
   ```

3. **Create Configuration**:
   ```cmd
   copy config.json.example config.json
   ```
   Then edit `config.json` with your settings (see Configuration section below).

4. **Setup Reports Folder**:
   - Create a `reports` folder in this directory
   - Copy your existing Large Deal Report workbook into it
   - Ensure the workbook has these worksheets:
     - `large deal report` (where daily data goes)
     - `summary` (with your table/pivot)
     - `iphone compatible` (for values-only copy)

### Running the Script

**Double-click** `run_report.bat` or run from Command Prompt:
```cmd
python large_deal_report_automation.py
```

The script will:
1. Find your most recent report in the `reports` folder
2. Download the daily worksheet from your "large trade td" email
3. Update the workbook automatically
4. Open Outlook with the email ready - **just review and click Send**

## Configuration

Edit `config.json` with your settings:

```json
{
  "email": {
    "incoming_subject": "large trade td",
    "use_outlook": true,
    "preview_before_send": true
  },
  "distribution_list": [
    "distribution-list@company.com",
    "manager@company.com"
  ],
  "sender_name": "Your Name",
  "email_body": "Hi all,\n\nPlease find attached today's Large Deal Report.\n\nKind regards,\nYour Name",
  "reports_directory": "./reports",
  "large_deal_report_sheet": "large deal report",
  "summary_sheet": "summary",
  "iphone_compatible_sheet": "iphone compatible"
}
```

### Configuration Options

| Setting | Description | Default |
|---------|-------------|---------|
| `incoming_subject` | Subject line to search for in incoming emails | `"large trade td"` |
| `use_outlook` | Use Outlook for both reading and sending emails | `true` |
| `preview_before_send` | Display email for review instead of auto-sending | `true` |
| `distribution_list` | Array of email addresses to send the report to | Required |
| `sender_name` | Your name for the email signature | Required |
| `email_body` | Custom email body text (optional) | Auto-generated |
| `reports_directory` | Folder containing your report workbooks | `"./reports"` |
| `large_deal_report_sheet` | Name of the data input worksheet | `"large deal report"` |
| `summary_sheet` | Name of the summary worksheet | `"summary"` |
| `iphone_compatible_sheet` | Name of the iPhone output worksheet | `"iphone compatible"` |

## Your Daily Workflow

### Before (Manual Process)
1. Open email with "large trade td" subject
2. Download the attachment
3. Open your Large Deal Report workbook
4. Copy the downloaded data
5. Paste into 'large deal report' sheet at A1
6. Go to 'summary' sheet, right-click table, refresh
7. Select entire summary table
8. Copy and go to 'iphone compatible' sheet
9. Paste Special → Values only
10. Save as "large deal report - {date}.xlsx"
11. Open Outlook, create new email
12. Add recipients, attach file, write message
13. Review and send

### After (Automated)
1. **Double-click `run_report.bat`**
2. Review the email preview that opens
3. **Click Send**

That's it! Everything else is automated.

## File Structure

```
your-folder/
├── large_deal_report_automation.py  # Main script
├── email_handler.py                  # Email handling
├── excel_processor.py                # Excel manipulation
├── config.json                       # Your configuration (create from example)
├── config.json.example               # Configuration template
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── run_report.bat                    # Windows launcher (double-click this!)
├── run_report.sh                     # macOS/Linux launcher
├── reports/                          # Your report files
│   └── large deal report - 06-01-2026.xlsx
└── temp/                             # Temporary files (auto-cleaned)
```

## How It Works

### Step-by-Step Process

1. **Find Latest Report**: Locates the most recent report in your `reports` folder to use as a template

2. **Download Email Attachment**: 
   - Connects to Outlook
   - Searches for emails with subject containing "large trade td"
   - Downloads the Excel attachment from the most recent matching email

3. **Update 'large deal report' Sheet**:
   - Clears existing data
   - Pastes the downloaded data starting at cell A1
   - Preserves formatting

4. **Refresh 'summary' Sheet**:
   - Uses Excel automation (xlwings) to refresh tables/pivots
   - Equivalent to right-clicking the table and selecting "Refresh"

5. **Copy to 'iphone compatible' Sheet**:
   - Copies the summary table
   - Pastes as **values only** (no formulas)
   - Preserves formatting

6. **Save Report**:
   - Saves as `large deal report - DD-MM-YYYY.xlsx`
   - Stored in your `reports` directory

7. **Email Preview**:
   - Creates email with your distribution list
   - Attaches the report
   - **Opens in Outlook for your review** (doesn't auto-send)
   - You click Send when ready

## Troubleshooting

### "No existing reports found"
- Make sure you have at least one report file in the `reports` folder
- The filename should contain "large deal report"

### "No emails found with subject containing 'large trade td'"
- Check that you've received today's email
- Verify the subject line matches (case-insensitive)
- The script will also find the most recent matching email if none from today

### "Worksheet 'xxx' not found"
- Check your workbook has the correct sheet names
- Sheet names are case-insensitive but must match your config
- Available sheets will be listed in the error message

### "Outlook COM error"
- Make sure Outlook is running and you're logged in
- Try opening Outlook manually first
- Ensure pywin32 is installed: `pip install pywin32`

### Table not refreshing properly
- Ensure Excel is installed and licensed
- The script uses Excel automation for refresh
- If refresh fails, you may need to manually refresh after the script runs

### Email preview doesn't open
- Make sure Outlook is your default email client
- Check that pywin32 is installed correctly
- Try running Outlook as administrator once

## Advanced Options

### Auto-Send Without Preview
If you want the email to send automatically without preview:
```json
{
  "email": {
    "preview_before_send": false
  }
}
```
**Warning:** This will send immediately without confirmation!

### Using IMAP Instead of Outlook
For non-Outlook email systems (Gmail, etc.):
```json
{
  "email": {
    "use_outlook": false,
    "username": "your_email@gmail.com",
    "password": "your_app_password",
    "imap_server": "imap.gmail.com",
    "imap_port": 993,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587
  }
}
```
Note: IMAP mode doesn't support email preview - it sends directly.

### Custom Email Message
Set a custom email body in your config:
```json
{
  "email_body": "Hi team,\n\nHere is today's Large Deal Report.\n\nPlease review and let me know if you have questions.\n\nBest regards,\nYour Name"
}
```

### Scheduling (Run Automatically Every Day)
1. Open **Task Scheduler** (search in Start menu)
2. Click **Create Basic Task**
3. Name: "Large Deal Report Automation"
4. Trigger: **Daily** at your preferred time (e.g., 9:00 AM)
5. Action: **Start a program**
6. Program: `python.exe`
7. Arguments: `"C:\full\path\to\large_deal_report_automation.py"`
8. Start in: `C:\full\path\to\folder`

## Dependencies

- `openpyxl` - Excel file reading/writing
- `xlwings` - Excel automation for table refresh
- `pywin32` - Outlook integration (Windows only)
- `pandas` - Data manipulation

All dependencies are listed in `requirements.txt` and installed with:
```cmd
pip install -r requirements.txt
```

## Security Notes

- Never commit `config.json` to version control (it may contain sensitive info)
- The script uses your logged-in Outlook account - no passwords stored
- Email preview mode lets you verify recipients before sending

## Support

If you encounter issues:
1. Check the console output for error messages
2. Verify your configuration file is valid JSON
3. Ensure all prerequisites are installed
4. Check that Outlook is running and logged in

## Version History

- **v2.0**: Complete rewrite for your specific workflow
  - Search by email subject instead of sender
  - Correct worksheet names (large deal report, summary, iphone compatible)
  - Email preview mode (review before sending)
  - DD-MM-YYYY date format
  - Simplified configuration
