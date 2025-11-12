# Email Folder Attachment App - PowerShell Version
# This script creates a GUI for selecting folders and creating Outlook email drafts with attachments
# REQUIRES Classic Outlook (New Outlook is not supported due to lack of COM API)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$global:SelectedFolder = ""
$global:UserEmail = ""
$global:Form = $null
$global:Controls = @{}
$global:OutlookType = "Unknown"  # "Classic" or "Unsupported"

# Function to detect Outlook type and get user email
function Get-OutlookTypeAndUserEmail {
    $global:OutlookType = "Unknown"

    # Fast check: Look for running Classic Outlook processes first
    $classicProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue

    if (-not $classicProcesses) {
        Write-Host "Classic Outlook not running - checking if COM API is available..."

        # Quick COM availability test (with timeout)
        try {
            $job = Start-Job -ScriptBlock {
                try {
                    $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
                    $outlook.Quit()
                    return $true
                }
                catch {
                    return $false
                }
            } -Name "OutlookCOMTest"

            # Wait maximum 2 seconds for COM test
            $result = $job | Wait-Job -Timeout 2 | Receive-Job
            Remove-Job $job -ErrorAction SilentlyContinue

            if (-not $result) {
                throw "COM test failed"
            }
        }
        catch {
            Write-Host "Classic Outlook COM API not available"
            $global:OutlookType = "Unsupported"

            # Show error message and exit
            $errorMessage = @"
CLASSIC OUTLOOK REQUIRED

This application requires Classic Outlook to function properly.

REASON: New Outlook (the modern version) does not provide a COM API for automation.
Classic Outlook uses COM (Component Object Model) which allows PowerShell scripts
to programmatically create emails, add attachments, and save drafts.

New Outlook lacks this automation capability, making it impossible to:
1. Automatically create email drafts
2. Programmatically attach files
3. Save emails as drafts without manual intervention

SOLUTION:
Please switch to Classic Outlook:
1. Open Outlook
2. Go to File → Office Account → Outlook Version
3. Switch to "Outlook Classic" or "New Outlook" toggle should be OFF

Then restart this application.
"@

            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Classic Outlook Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            exit 1
        }
    }

    # If we get here, Classic Outlook is available - get email address
    $global:OutlookType = "Classic"
    Write-Host "Classic Outlook detected - getting user email..."

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")

        # Try to get from default account (fastest)
        try {
            $accounts = $namespace.Accounts
            if ($accounts.Count -gt 0) {
                $defaultAccount = $accounts.Item(1)
                if ($defaultAccount.SmtpAddress) {
                    $outlook.Quit()
                    return $defaultAccount.SmtpAddress
                }
            }
        }
        catch {
            # Continue to fallback methods
        }

        # Try current user
        try {
            $currentUser = $namespace.CurrentUser
            if ($currentUser.Address -and $currentUser.Address.Contains("@")) {
                $outlook.Quit()
                return $currentUser.Address
            }
        }
        catch {
            # Continue to fallback
        }

        $outlook.Quit()
    }
    catch {
        Write-Host "Failed to get email from Outlook: $($_.Exception.Message)"
    }

    # Fallback: construct email from username and domain
    $username = $env:USERNAME
    $domain = Get-CompanyDomain
    return "$username@$domain"
}

# Function to get current user's email (wrapper for backward compatibility)
function Get-UserEmail {
    return Get-OutlookTypeAndUserEmail
}
# Function to detect company domain
function Get-CompanyDomain {
    try {
        # Try USERDNSDOMAIN environment variable
        if ($env:USERDNSDOMAIN -and $env:USERDNSDOMAIN.Contains(".")) {
            return $env:USERDNSDOMAIN.ToLower()
        }
        
        # Try to get FQDN
        try {
            $fqdn = [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName
            if ($fqdn.Contains(".") -and -not $fqdn.EndsWith(".local")) {
                $parts = $fqdn.Split(".")
                if ($parts.Length -gt 1) {
                    return ($parts[1..($parts.Length-1)] -join ".")
                }
            }
        }
        catch {
            # Continue with fallback
        }
        
        # Default fallback - customize this for your organization
        return "aa.com"
    }
    catch {
        return "aa.com"
    }
}

# Function to browse for folder
function Browse-Folder {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select folder containing files to attach"
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:SelectedFolder = $folderBrowser.SelectedPath
        $global:Controls.FolderTextBox.Text = $global:SelectedFolder
        Update-FileList
    }
}

# Function to update file list
function Update-FileList {
    $global:Controls.FileListBox.Items.Clear()
    
    if (-not $global:SelectedFolder -or -not (Test-Path $global:SelectedFolder)) {
        return
    }
    
    # Get files with allowed extensions
    $allowedExts = @('.pdf', '.pptx', '.xlsx')
    $files = Get-ChildItem -Path $global:SelectedFolder -File | Where-Object { 
        $allowedExts -contains $_.Extension.ToLower() 
    } | Sort-Object Name
    
    if ($files.Count -gt 0) {
        foreach ($file in $files) {
            $global:Controls.FileListBox.Items.Add($file.Name)
        }
        $global:Controls.StatusLabel.Text = "Found $($files.Count) files to attach"
        $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Blue
    }
    else {
        $global:Controls.FileListBox.Items.Add("No PDF, PPTX, or XLSX files found")
        $global:Controls.StatusLabel.Text = "No attachable files found in selected folder"
        $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Orange
    }
}

# Function to send to self
function Send-ToSelf {
    $global:Controls.RecipientTextBox.Text = $global:UserEmail
}

# Function to create email draft (supports both Classic and New Outlook)
function Create-EmailDraft {
    # Validate inputs
    if (-not $global:SelectedFolder) {
        [System.Windows.Forms.MessageBox]::Show("Please select a folder first", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    if (-not $global:Controls.RecipientTextBox.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Please enter recipient email address", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)  
        return
    }
    
    if (-not $global:Controls.SubjectTextBox.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Please enter email subject", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Disable button and show status
    $global:Controls.CreateDraftButton.Enabled = $false
    $global:Controls.StatusLabel.Text = "Creating email draft..."
    $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Blue
    $global:Form.Refresh()
    
    try {
        $folderPath = $global:SelectedFolder
        $recipient = $global:Controls.RecipientTextBox.Text.Trim()
        $subject = $global:Controls.SubjectTextBox.Text.Trim()
        $bodyText = $global:Controls.EmailBodyTextBox.Text.Trim()
        
        # Get files to attach
        $allowedExts = @('.pdf', '.pptx', '.xlsx')
        $files = Get-ChildItem -Path $folderPath -File | Where-Object { 
            $allowedExts -contains $_.Extension.ToLower() 
        }
        
        if ($files.Count -eq 0) {
            $global:Controls.StatusLabel.Text = "No files to attach found"
            $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Orange
            return
        }
        
        # Try Classic Outlook only
        $success = Create-ClassicOutlookDraft -Recipient $recipient -Subject $subject -Body $bodyText -Files $files

        if ($success) {
            $global:Controls.StatusLabel.Text = "Email draft created successfully with $($files.Count) attachments!"
            $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Green
        } else {
            $global:Controls.StatusLabel.Text = "Failed to create email draft. Please check if Classic Outlook is running."
            $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    }
    catch {
        $errorMsg = "Error creating email: $($_.Exception.Message)"
        $global:Controls.StatusLabel.Text = $errorMsg
        $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $global:Controls.CreateDraftButton.Enabled = $true
    }
}

# Function to create draft in Classic Outlook (COM-based)
function Create-ClassicOutlookDraft {
    param(
        [string]$Recipient,
        [string]$Subject,
        [string]$Body,
        [array]$Files
    )
    
    try {
        # Convert plain text to HTML
        $htmlBody = $Body -replace "`r`n", "<br>" -replace "`n", "<br>"
        $htmlBody = "<div style='font-family: Arial, sans-serif;'>$htmlBody</div>"
        
        # Create Outlook draft
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0) # olMailItem
        $mail.Subject = $Subject
        $mail.HTMLBody = $htmlBody
        
        $recipientObj = $mail.Recipients.Add($Recipient)
        $recipientObj.Type = 1 # olTo
        $mail.Recipients.ResolveAll()
        
        # Add attachments
        foreach ($file in $Files) {
            $mail.Attachments.Add($file.FullName)
        }
        
        # Save as draft
        $mail.Save()
        
        Write-Host "Classic Outlook draft created successfully"
        return $true
    }
    catch {
        Write-Host "Classic Outlook failed: $($_.Exception.Message)"
        return $false
    }
}

# Function to clear all fields
function Clear-All {
    $global:SelectedFolder = ""
    $global:Controls.FolderTextBox.Text = ""
    $global:Controls.RecipientTextBox.Text = ""
    $global:Controls.SubjectTextBox.Text = "Documents Attached"
    $global:Controls.EmailBodyTextBox.Text = "Dear recipient,`r`n`r`nPlease find the attached documents.`r`n`r`nBest regards"
    $global:Controls.FileListBox.Items.Clear()
    $global:Controls.StatusLabel.Text = "Ready"
    $global:Controls.StatusLabel.ForeColor = [System.Drawing.Color]::Green
}

# Function to create the main form
function Create-MainForm {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Email Folder Attachment App"
    $form.Size = New-Object System.Drawing.Size(620, 560)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    
    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Email Folder Attachment App"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(560, 30)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($titleLabel)
    
    # Outlook version info
    $outlookVersionLabel = New-Object System.Windows.Forms.Label
    $outlookVersionText = switch ($global:OutlookType) {
        "Classic" { "[OK] Classic Outlook detected - Full automation available" }
        default { "[ERROR] Classic Outlook not detected - Application will exit" }
    }
    $outlookVersionLabel.Text = $outlookVersionText
    $outlookVersionLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Italic)
    $outlookVersionLabel.Location = New-Object System.Drawing.Point(20, 50)
    $outlookVersionLabel.Size = New-Object System.Drawing.Size(560, 15)
    $outlookVersionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $outlookVersionLabel.ForeColor = switch ($global:OutlookType) {
        "Classic" { [System.Drawing.Color]::Green }
        default { [System.Drawing.Color]::Red }
    }
    $form.Controls.Add($outlookVersionLabel)
    
    # Folder selection
    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Text = "Select Folder:"
    $folderLabel.Location = New-Object System.Drawing.Point(20, 80)
    $folderLabel.Size = New-Object System.Drawing.Size(100, 23)
    $form.Controls.Add($folderLabel)
    
    $folderTextBox = New-Object System.Windows.Forms.TextBox
    $folderTextBox.Location = New-Object System.Drawing.Point(130, 80)
    $folderTextBox.Size = New-Object System.Drawing.Size(380, 23)
    $folderTextBox.ReadOnly = $true
    $form.Controls.Add($folderTextBox)
    $global:Controls.FolderTextBox = $folderTextBox
    
    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = "Browse"
    $browseButton.Location = New-Object System.Drawing.Point(520, 79)
    $browseButton.Size = New-Object System.Drawing.Size(75, 25)
    $browseButton.Add_Click({ Browse-Folder })
    $form.Controls.Add($browseButton)
    
    # Current user email
    $userEmailLabel = New-Object System.Windows.Forms.Label
    $userEmailLabel.Text = "Your Email:"
    $userEmailLabel.Location = New-Object System.Drawing.Point(20, 115)
    $userEmailLabel.Size = New-Object System.Drawing.Size(100, 23)
    $form.Controls.Add($userEmailLabel)
    
    $userEmailTextBox = New-Object System.Windows.Forms.TextBox
    $userEmailTextBox.Location = New-Object System.Drawing.Point(130, 115)
    $userEmailTextBox.Size = New-Object System.Drawing.Size(300, 23)
    $userEmailTextBox.ReadOnly = $true
    $userEmailTextBox.Text = $global:UserEmail
    $form.Controls.Add($userEmailTextBox)
    
    $sendToSelfButton = New-Object System.Windows.Forms.Button
    $sendToSelfButton.Text = "Send to Self"
    $sendToSelfButton.Location = New-Object System.Drawing.Point(440, 114)
    $sendToSelfButton.Size = New-Object System.Drawing.Size(90, 25)
    $sendToSelfButton.Add_Click({ Send-ToSelf })
    $form.Controls.Add($sendToSelfButton)
    
    # Recipient email
    $recipientLabel = New-Object System.Windows.Forms.Label
    $recipientLabel.Text = "Recipient Email:"
    $recipientLabel.Location = New-Object System.Drawing.Point(20, 150)
    $recipientLabel.Size = New-Object System.Drawing.Size(100, 23)
    $form.Controls.Add($recipientLabel)
    
    $recipientTextBox = New-Object System.Windows.Forms.TextBox
    $recipientTextBox.Location = New-Object System.Drawing.Point(130, 150)
    $recipientTextBox.Size = New-Object System.Drawing.Size(465, 23)
    $form.Controls.Add($recipientTextBox)
    $global:Controls.RecipientTextBox = $recipientTextBox
    
    # Email subject
    $subjectLabel = New-Object System.Windows.Forms.Label
    $subjectLabel.Text = "Email Subject:"
    $subjectLabel.Location = New-Object System.Drawing.Point(20, 185)
    $subjectLabel.Size = New-Object System.Drawing.Size(100, 23)
    $form.Controls.Add($subjectLabel)
    
    $subjectTextBox = New-Object System.Windows.Forms.TextBox
    $subjectTextBox.Location = New-Object System.Drawing.Point(130, 185)
    $subjectTextBox.Size = New-Object System.Drawing.Size(465, 23)
    $subjectTextBox.Text = "Documents Attached"
    $form.Controls.Add($subjectTextBox)
    $global:Controls.SubjectTextBox = $subjectTextBox
    
    # Email body
    $bodyLabel = New-Object System.Windows.Forms.Label
    $bodyLabel.Text = "Email Body:"
    $bodyLabel.Location = New-Object System.Drawing.Point(20, 220)
    $bodyLabel.Size = New-Object System.Drawing.Size(100, 23)
    $form.Controls.Add($bodyLabel)
    
    $emailBodyTextBox = New-Object System.Windows.Forms.TextBox
    $emailBodyTextBox.Location = New-Object System.Drawing.Point(130, 220)
    $emailBodyTextBox.Size = New-Object System.Drawing.Size(465, 80)
    $emailBodyTextBox.Multiline = $true
    $emailBodyTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $emailBodyTextBox.Text = "Dear recipient,`r`n`r`nPlease find the attached documents.`r`n`r`nBest regards"
    $form.Controls.Add($emailBodyTextBox)
    $global:Controls.EmailBodyTextBox = $emailBodyTextBox
    
    # File list
    $fileListLabel = New-Object System.Windows.Forms.Label
    $fileListLabel.Text = "Files to attach:"
    $fileListLabel.Location = New-Object System.Drawing.Point(20, 320)
    $fileListLabel.Size = New-Object System.Drawing.Size(100, 23)
    $form.Controls.Add($fileListLabel)
    
    $fileListBox = New-Object System.Windows.Forms.ListBox
    $fileListBox.Location = New-Object System.Drawing.Point(130, 320)
    $fileListBox.Size = New-Object System.Drawing.Size(465, 90)
    $form.Controls.Add($fileListBox)
    $global:Controls.FileListBox = $fileListBox
    
    # Buttons
    $createDraftButton = New-Object System.Windows.Forms.Button
    $createDraftButton.Text = "Create Email Draft"
    $createDraftButton.Location = New-Object System.Drawing.Point(130, 430)
    $createDraftButton.Size = New-Object System.Drawing.Size(130, 30)
    $createDraftButton.BackColor = [System.Drawing.Color]::LightBlue
    $createDraftButton.Add_Click({ Create-EmailDraft })
    $form.Controls.Add($createDraftButton)
    $global:Controls.CreateDraftButton = $createDraftButton
    
    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = "Clear All"
    $clearButton.Location = New-Object System.Drawing.Point(280, 430)
    $clearButton.Size = New-Object System.Drawing.Size(80, 30)
    $clearButton.Add_Click({ Clear-All })
    $form.Controls.Add($clearButton)
    
    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Ready"
    $statusLabel.Location = New-Object System.Drawing.Point(130, 475)
    $statusLabel.Size = New-Object System.Drawing.Size(465, 23)
    $statusLabel.ForeColor = [System.Drawing.Color]::Green
    $form.Controls.Add($statusLabel)
    $global:Controls.StatusLabel = $statusLabel
    
    return $form
}

# Main execution
try {
    Write-Host "Starting Email Folder Attachment App..."
    
    # Get user email and detect Outlook type
    $global:UserEmail = Get-OutlookTypeAndUserEmail
    Write-Host "User email detected: $global:UserEmail"
    Write-Host "Outlook type detected: $global:OutlookType"
    
    # Create and show form
    $global:Form = Create-MainForm
    
    # Show the form
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::Run($global:Form)
}
catch {
    Write-Host "Error starting application: $_"
    [System.Windows.Forms.MessageBox]::Show("Error starting application: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}