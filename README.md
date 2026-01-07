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
- `pywin32` - Windows COM interface for Outlook integration (Windows only, optional)

**Note:** 
- `xlwings` requires Excel to be installed. If you don't have Excel or prefer not to use xlwings, the script will fall back to openpyxl (though data connections won't be automatically refreshed).
- `pywin32` is required if you want to use Outlook for sending emails instead of SMTP (Windows only).

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

#### Email Sending Options

You have two options for sending emails:

**Option 1: Use Outlook (Recommended for Windows 11)**

Set `"use_outlook_for_sending": true` in your config. This will:
- Use your Outlook application to send emails
- No SMTP credentials needed
- Works with Exchange/Office 365 accounts
- Uses your logged-in Outlook account

**Option 2: Use SMTP (Traditional Method)**

Set `"use_outlook_for_sending": false` (or omit it) and provide:
- `username`: Your email address
- `password`: App Password (for Gmail) or regular password
- `smtp_server`: SMTP server address
- `smtp_port`: SMTP server port

#### Configuration Options

- **email.username**: Your email address for IMAP/SMTP (required if not using Outlook for sending)
- **email.password**: App Password (for Gmail) or regular password (required if not using Outlook for sending)
- **email.sender_email**: The email address that sends you the daily worksheet (required)
- **email.imap_server**: IMAP server address (default: `imap.gmail.com`) - used for receiving emails
- **email.imap_port**: IMAP server port (default: `993`) - used for receiving emails
- **email.smtp_server**: SMTP server address (default: `smtp.gmail.com`) - only needed if not using Outlook
- **email.smtp_port**: SMTP server port (default: `587`) - only needed if not using Outlook
- **email.use_outlook_for_sending**: Set to `true` to use Outlook COM for sending (Windows only, recommended)
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

### Spreadsheet consolidation (case-insensitive column match + grouped/merged output)

If you have two spreadsheets with similar columns (e.g. one uses `SEDOL` and the other uses `Sedol` / `Fund Type`), you can extract a fixed set of columns, merge the rows into one sheet, sort by **ISSUE NAME**, and apply hierarchical “merge & centre” grouping (ISSUE NAME → SEDOL → ISIN → …):

```bash
python3 merge_and_align.py consolidate corporate_actions.xlsx accounts_list.xlsx merged_output.xlsx
```

By default the output columns (uppercase) are:

- `ISSUE NAME`, `SEDOL`, `ISIN`, `ISSUE COUNTRY NAME`, `STRATEGY`, `SUB STRATEGY`, `FUND TYPE`, `FUND NAME`

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
- **Alternative**: Use Outlook for sending by setting `"use_outlook_for_sending": true` (no SMTP credentials needed)

### "Outlook COM requested but win32com not available"

**Problem:** Trying to use Outlook but pywin32 is not installed.

**Solutions:**
- Install pywin32: `pip install pywin32`
- Make sure you're on Windows (Outlook COM only works on Windows)
- Ensure Outlook is installed on your system
- If you don't want to use Outlook, set `"use_outlook_for_sending": false` and provide SMTP credentials

### "Error sending email via Outlook"

**Problem:** Outlook cannot send the email.

**Solutions:**
- Make sure Outlook is installed and you're logged into your account
- Verify Outlook is not blocked by Windows security settings
- Try opening Outlook manually first to ensure it's properly initialized
- Check that your Outlook account has permission to send emails
- Ensure the report file path is accessible

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
