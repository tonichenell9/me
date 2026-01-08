# Large Deal Report Automation

**One-click automation for the daily Large Deal Report workflow on Windows 11.**

This Python script automates your entire daily reporting process:
1. Downloads the daily worksheet from your "large trade td" email
2. Pastes the data into your workbook
3. Refreshes the summary table
4. Creates the iPhone compatible version (values only)
5. Saves the report with today's date
6. Opens an email preview for you to review before sending

---

## Step-by-Step Setup Guide

Follow these steps carefully to set up the automation on your Windows 11 PC.

### Step 1: Check Prerequisites

Before starting, make sure you have:

- ✅ **Windows 11** PC
- ✅ **Microsoft Excel** installed (Office 365 or standalone)
- ✅ **Microsoft Outlook** installed and logged into your email account
- ✅ Your existing Large Deal Report workbook with these sheets:
  - `large deal report`
  - `summary`
  - `iphone compatible`

### Step 2: Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/)
2. Click the big yellow **"Download Python 3.x.x"** button
3. Run the downloaded installer
4. **⚠️ IMPORTANT:** On the first screen, tick the box that says **"Add Python to PATH"**
5. Click **"Install Now"**
6. Wait for installation to complete, then click **"Close"**

**To verify Python is installed:**
1. Press `Win + R`, type `cmd`, press Enter
2. Type `python --version` and press Enter
3. You should see something like `Python 3.12.0`

### Step 3: Download/Extract the Script Files

1. Download or copy all the script files to a folder on your PC
2. Recommended location: `C:\Users\YourName\Documents\LargeDealReport\`
3. Make sure you have these files in the folder:
   - `large_deal_report_automation.py`
   - `email_handler.py`
   - `excel_processor.py`
   - `config.json.example`
   - `requirements.txt`
   - `run_report.bat`

### Step 4: Install Python Dependencies

1. Open **Command Prompt**:
   - Press `Win + R`
   - Type `cmd`
   - Press Enter

2. Navigate to your script folder:
   ```cmd
   cd C:\Users\YourName\Documents\LargeDealReport
   ```
   (Replace with your actual folder path)

3. Install the required packages:
   ```cmd
   pip install -r requirements.txt
   ```

4. Wait for installation to complete (may take 1-2 minutes)

### Step 5: Create Your Configuration File

1. In your script folder, find `config.json.example`
2. Make a copy of it and rename the copy to `config.json`
   - Right-click `config.json.example` → Copy
   - Right-click in empty space → Paste
   - Right-click the copy → Rename → `config.json`

3. Open `config.json` with Notepad:
   - Right-click `config.json` → Open with → Notepad

4. Edit the following settings:

```json
{
  "email": {
    "incoming_subject": "large trade td",
    "use_outlook": true,
    "preview_before_send": true
  },
  "distribution_list": [
    "first-distribution-list@yourcompany.com",
    "second-distribution-list@yourcompany.com",
    "manager@yourcompany.com"
  ],
  "sender_name": "Your Actual Name",
  "email_body": "Hi all,\n\nPlease find attached today's Large Deal Report.\n\nKind regards,\nYour Name",
  "reports_directory": "./reports",
  "large_deal_report_sheet": "large deal report",
  "summary_sheet": "summary",
  "iphone_compatible_sheet": "iphone compatible"
}
```

5. **What to change:**
   - `distribution_list`: Replace with your actual email addresses/distribution lists
   - `sender_name`: Replace with your name
   - `email_body`: Customise your email message (use `\n` for new lines)

6. Save the file (Ctrl + S) and close Notepad

### Step 6: Set Up the Reports Folder

1. In your script folder, create a new folder called `reports`
   - Right-click in empty space → New → Folder → name it `reports`

2. Copy your existing Large Deal Report workbook into this `reports` folder
   - This will be used as the template

3. **Important:** Make sure your workbook has these exact worksheet names (tabs):
   - `large deal report` (where the daily data goes)
   - `summary` (with your table that needs refreshing)
   - `iphone compatible` (for the values-only copy)

### Step 7: Test the Setup

1. Make sure **Outlook is open** and you're logged in
2. Make sure you have received an email with subject containing "large trade td"
3. Double-click `run_report.bat`
4. Watch the Command Prompt window - it will show progress
5. When complete, Outlook will open with the email ready to send
6. **Review the email** and click Send when ready

### Step 8: Create a Desktop Shortcut (Optional)

For easy daily access:

1. Right-click on `run_report.bat`
2. Click **"Create shortcut"**
3. Drag the shortcut to your Desktop
4. Rename it to "Large Deal Report" if you like

Now you can double-click the shortcut each day to run the automation!

---

## Daily Usage

Once set up, your daily workflow is simple:

### Every Morning:

1. ✅ Make sure Outlook is open
2. ✅ Make sure you've received the "large trade td" email
3. ✅ Double-click `run_report.bat` (or your desktop shortcut)
4. ✅ Wait for the script to complete (~30 seconds)
5. ✅ Review the email that opens in Outlook
6. ✅ Click **Send**

**That's it!** The script handles everything else automatically.

---

## Configuration Reference

### All Configuration Options

| Setting | Description | Default |
|---------|-------------|---------|
| `incoming_subject` | Subject line to search for in incoming emails (supports date placeholders) | `"large trade td"` |
| `use_outlook` | Use Outlook for both reading and sending emails | `true` |
| `preview_before_send` | Display email for review instead of auto-sending | `true` |
| `distribution_list` | Array of email addresses to send the report to | Required |
| `sender_name` | Your name for the email signature | Required |
| `email_body` | Custom email body text (use `\n` for new lines) | Auto-generated |
| `reports_directory` | Folder containing your report workbooks | `"./reports"` |
| `large_deal_report_sheet` | Name of the data input worksheet | `"large deal report"` |
| `summary_sheet` | Name of the summary worksheet | `"summary"` |
| `iphone_compatible_sheet` | Name of the iPhone output worksheet | `"iphone compatible"` |

### Date Placeholders for incoming_subject

You can include today's date in the email subject search using these placeholders:

**Long Date Formats:**
| Placeholder | Format | Example |
|-------------|--------|---------|
| `{date_long}` | DD Month YYYY | 07 January 2026 |
| `{date_long_day}` | Day, DD Month YYYY | Tuesday, 07 January 2026 |

**Short Date Formats:**
| Placeholder | Format | Example |
|-------------|--------|---------|
| `{date}` | DD/MM/YYYY | 07/01/2026 |
| `{date_dash}` | DD-MM-YYYY | 07-01-2026 |
| `{date_dot}` | DD.MM.YYYY | 07.01.2026 |

**Individual Components:**
| Placeholder | Format | Example |
|-------------|--------|---------|
| `{day_name}` | Day name | Tuesday |
| `{month_name}` | Month name | January |
| `{dd}` | Day (with zero) | 07 |
| `{d}` | Day (no zero) | 7 |
| `{mm}` | Month | 01 |
| `{yyyy}` | Year | 2026 |
| `{yy}` | Short year | 26 |

**Examples:**
```json
"incoming_subject": "large deal report - {date_long}"       // "large deal report - 07 January 2026"
"incoming_subject": "large deal report - {date_long_day}"  // "large deal report - Tuesday, 07 January 2026"
"incoming_subject": "large deal report - {date}"           // "large deal report - 07/01/2026"
"incoming_subject": "large deal report - {date_dash}"      // "large deal report - 07-01-2026"
"incoming_subject": "report {d} {month_name} {yyyy}"       // "report 7 January 2026"
```

### Example config.json

```json
{
  "email": {
    "incoming_subject": "large trade td",
    "use_outlook": true,
    "preview_before_send": true
  },
  "distribution_list": [
    "trading-team@company.com",
    "management@company.com"
  ],
  "sender_name": "John Smith",
  "email_body": "Hi all,\n\nPlease find attached today's Large Deal Report.\n\nKind regards,\nJohn",
  "reports_directory": "./reports",
  "large_deal_report_sheet": "large deal report",
  "summary_sheet": "summary",
  "iphone_compatible_sheet": "iphone compatible"
}
```

---

## What the Script Replaces

### Before (Manual Process - ~10 minutes)
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

### After (Automated - ~30 seconds)
1. **Double-click `run_report.bat`**
2. Review the email preview that opens
3. **Click Send**

That's it! The script does steps 1-12 automatically.

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
