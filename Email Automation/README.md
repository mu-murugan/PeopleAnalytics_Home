# Email Folder Attachment App - Python Version

A simple Windows GUI application that allows users to select a folder and create an Outlook email draft with all supported files from that folder as attachments.

## Features

- **Simple GUI Interface**: Easy-to-use Windows application with tkinter
- **Automatic User Detection**: Automatically detects the current user's email address from Outlook
- **Corporate Environment Ready**: Perfect for shared drives - works with any company laptop
- **Folder Selection**: Browse and select any folder on your system
- **File Filtering**: Automatically filters and attaches only PDF, PPTX, and XLSX files
- **Email Customization**: 
  - Auto-detects your email address
  - "Send to Self" quick button
  - Set recipient email address
  - Customize email subject
  - Edit email body content
- **File Preview**: View list of files that will be attached before creating the draft
- **Outlook Integration**: Creates drafts directly in Microsoft Outlook
- **Threading**: Non-blocking UI during email creation
- **Domain Detection**: Automatically detects company domain for fallback scenarios

## Requirements

- Windows OS
- Python 3.6 or higher
- Microsoft Outlook installed and configured
- Required Python packages (automatically installed):
  - pywin32

## Installation

1. Ensure Python is installed on your system
2. The app will automatically install required packages on first run

## Usage

### Method 1: Double-click the batch file
Simply double-click `run_app.bat` to start the application.

### Method 2: Command line
```bash
python email_folder_app.py
```

## How to Use

1. **Your Email**: The app automatically detects your email address from Outlook
2. **Select Folder**: Click "Browse" to select a folder containing files you want to attach
3. **Set Recipient**: 
   - Type the recipient's email address, OR
   - Click "Send to Self" to email yourself
4. **Set Subject**: Modify the email subject if needed (default: "Documents Attached")
5. **Edit Body**: Customize the email message body
6. **Review Files**: Check the file list to see what will be attached
7. **Create Draft**: Click "Create Email Draft" to generate the draft in Outlook

### Perfect for Corporate Environments
- Put the `.bat` file on a shared drive
- Each user gets their own email automatically detected
- No configuration needed per user

## Supported File Types

The application only attaches files with these extensions:
- `.pdf` - PDF documents
- `.pptx` - PowerPoint presentations  
- `.xlsx` - Excel spreadsheets

## Notes

- The application creates email **drafts** in Outlook - you still need to manually send them
- Make sure Outlook is installed and configured with an email account
- The app runs in a separate thread when creating drafts to keep the UI responsive
- All fields can be cleared using the "Clear All" button

## Troubleshooting

- **"No module named 'win32com'"**: The app should auto-install this, but you can manually run: `pip install pywin32`
- **Outlook not opening**: Ensure Microsoft Outlook is installed and has been run at least once
- **No files found**: Check that your selected folder contains PDF, PPTX, or XLSX files
- **Permission errors**: Run as administrator if you encounter permission issues

## File Structure

```
├── email_folder_app.py    # Main application file
├── send_email.py          # Original batch processing script
├── run_app.bat           # Easy-run batch file
├── requirements.txt       # Python dependencies
└── README.md             # This file
```

# Email Folder Attachment App - PowerShell Version

This is a PowerShell version of the Email Folder Attachment App that works without requiring Python installation.

## Features

- **No Python Required**: Uses built-in Windows PowerShell and Windows Forms
- **Dual Outlook Support**: Works with both Classic Outlook (COM) and New Outlook (web-based)
- **Automatic Email Detection**: Detects your Outlook email address automatically
- **Smart Outlook Detection**: Automatically identifies which version of Outlook you have
- **File Filtering**: Only shows PDF, PPTX, and XLSX files for attachment
- **Adaptive Workflow**: Full automation for Classic Outlook, guided process for New Outlook
- **User-Friendly GUI**: Similar interface to the original Python version with Outlook type indicators

## Requirements

- Windows operating system (Windows 7 or later)
- Microsoft Outlook installed and configured (Classic Outlook or New Outlook)
- PowerShell (built into Windows)
- No additional installations required

## Outlook Version Support

### Classic Outlook (Traditional Desktop)
- **Full Automation**: Creates complete email drafts with attachments automatically
- **Status**: ✓ Classic Outlook detected - Full automation available
- **Experience**: Click "Create Email Draft" → Email appears in Outlook drafts folder

### New Outlook (Modern/Web-based)
- **Guided Process**: Provides instructions and opens Outlook with pre-filled details
- **Status**: ⚠ New Outlook detected - Manual attachment required  
- **Experience**: Click "Prepare Email" → Follow guided instructions to complete email

### Unknown/Mixed Environments
- **Adaptive**: Tries Classic Outlook first, falls back to New Outlook method
- **Status**: ⚠ Outlook type unknown - Will attempt both methods

## How to Use

### Method 1: Double-click the batch file (Recommended)
1. Double-click `Run_EmailApp.bat`
2. The application will start automatically

### Method 2: Run PowerShell script directly
1. Right-click on `EmailFolderApp.ps1`
2. Select "Run with PowerShell"
3. If prompted about execution policy, choose "Yes" or "Yes to All"

### Method 3: Command line
```powershell
powershell -ExecutionPolicy Bypass -File "EmailFolderApp.ps1"
```

## Using the Application

1. **Check Outlook Type**: The app will display which Outlook version is detected at the top
2. **Select Folder**: Click "Browse" to select a folder containing your files
3. **Review Files**: The app will automatically show all PDF, PPTX, and XLSX files found
4. **Set Recipient**: Enter the recipient's email address (or click "Send to Self")
5. **Customize Email**: Modify the subject and body text as needed
6. **Create Email**: 
   - **Classic Outlook**: Click "Create Email Draft" → Email appears in drafts automatically
   - **New Outlook**: Click "Prepare Email" → Follow the guided instructions
7. **Send**: Review and send the email from Outlook

### For New Outlook Users
When using New Outlook, the app will:
1. Copy email details to your clipboard
2. Show detailed instructions
3. Automatically open New Outlook (if possible)
4. Open Windows Explorer to the files location
5. Guide you through manual attachment process

## File Types Supported

- PDF files (.pdf)
- PowerPoint presentations (.pptx)
- Excel spreadsheets (.xlsx)

## Troubleshooting

### "Execution Policy" Error
If you get an execution policy error:
1. Open PowerShell as Administrator
2. Run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
3. Choose "Yes" when prompted
4. Try running the app again

### "Cannot Create COM Object" Error
This usually means Outlook is not installed or not properly configured:
1. Ensure Microsoft Outlook is installed
2. Open Outlook at least once to complete initial setup
3. Make sure Outlook is configured with your email account

### Email Address Not Detected
If your email address shows as "username@aa.com":
1. The app uses a fallback domain
2. Edit line 67 in `EmailFolderApp.ps1` to change "aa.com" to your company domain
3. Or manually enter your correct email and click "Send to Self"

### Files Not Showing
If no files appear in the list:
1. Ensure the selected folder contains PDF, PPTX, or XLSX files
2. Check that the files are not corrupted or locked
3. Verify you have read permissions for the folder

## Customization

You can customize the following in `EmailFolderApp.ps1`:

- **Default Domain** (line 67): Change "aa.com" to your organization's domain
- **File Types** (lines 92, 184): Add or remove file extensions in the `$allowedExts` arrays
- **Default Email Body** (line 429): Modify the default email message template

## Security Notes

- This script only reads files and creates email drafts
- No files are uploaded or sent to external servers
- All operations are performed locally on your machine
- Uses standard Windows and Outlook COM interfaces

## Support

If you encounter issues:
1. Check the troubleshooting section above
2. Ensure all requirements are met
3. Contact your IT department if execution policies prevent running
4. Verify Outlook is properly configured

## Version History

- **v1.0**: Initial PowerShell conversion from Python version
- Features identical functionality to original Python app
- Added automatic email detection and domain fallback
- Improved error handling and user feedback

---

**Note**: This PowerShell version provides the same functionality as the Python version but without requiring Python installation, making it suitable for corporate environments with restrictive IT policies.