# Shared Drive Deployment Guide

## Quick Setup for Shared Drive

### For IT/Admin:

1. **Copy Files to Shared Drive:**
   ```
   \\shared-drive\tools\EmailFolderApp\
   ├── email_folder_app.py
   ├── run_app.bat
   ├── requirements.txt
   └── README.md
   ```

2. **Set Permissions:**
   - Read & Execute permissions for all users
   - No write permissions needed in app folder

3. **Test with Different Users:**
   - Each user should see their own email auto-detected
   - App should work from any company laptop

### For End Users:

1. **Navigate to:** `\\shared-drive\tools\EmailFolderApp\`
2. **Double-click:** `run_app.bat`
3. **First Run:** App will auto-install Python packages if needed

## How It Works

### Automatic Email Detection Priority:
1. **Primary:** Outlook default account SMTP address
2. **Secondary:** Outlook current user address  
3. **Tertiary:** Exchange user primary SMTP
4. **Fallback:** `{username}@{detected-domain}`

### Domain Detection Methods:
1. **USERDNSDOMAIN** environment variable
2. **Computer FQDN** domain extraction
3. **Known corporate patterns** (customizable)
4. **Default:** company.com

## Customization for Your Company

### Update Company Domain Detection:
Edit `email_folder_app.py`, find the `detect_company_domain()` method:

```python
# Common corporate domain patterns - customize this section
if 'yourcompany.com' in str(os.environ.get('LOGONSERVER', '')).lower():
    return 'yourcompany.com'  # Replace with your actual domain

# Default fallback - change this to your company domain
return 'yourcompany.com'  # Replace company.com with your domain
```

### Network Drive Shortcuts:
Create desktop shortcuts pointing to:
```
\\shared-drive\tools\EmailFolderApp\run_app.bat
```

## Troubleshooting

### Common Issues:

1. **"Python not found"**
   - Install Python on user's machine
   - Or use portable Python deployment

2. **"Cannot access Outlook"**
   - Ensure Outlook is installed and configured
   - User must have run Outlook at least once

3. **Wrong email detected**
   - App shows detected email in read-only field
   - User can still manually enter correct recipient

4. **Network drive access**
   - Check network permissions
   - Ensure shared drive is accessible

### Error Logs:
- Errors print to console when running from batch file
- For silent deployment, redirect output to log file

## Security Notes

✅ **Safe for Corporate Use:**
- No data is stored or transmitted outside Outlook
- Only reads user's own email address
- Creates local Outlook drafts only
- No network communication except standard Outlook operations

✅ **Privacy Compliant:**
- Only accesses user's own Outlook profile
- No logging of user data
- No external connections

## Advanced Deployment

### Silent Install with Custom Domain:
```batch
@echo off
set COMPANY_DOMAIN=yourcompany.com
python email_folder_app.py --domain %COMPANY_DOMAIN%
```

### Group Policy Deployment:
- Deploy via software installation policy
- Set registry keys for company domain
- Create Start Menu shortcuts automatically