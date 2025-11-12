import subprocess
import sys

def ensure_package(pkg_name, import_name=None):
    try:
        __import__(import_name or pkg_name)
    except ImportError:
        print(f"[INFO] Installing missing package: {pkg_name}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name])

# Ensure required packages are installed
ensure_package("pywin32", "win32com")

import sys
import os
import csv
import win32com.client

#def save_draft(from_addr, to_addr, subject, html_body, folder_path):
def save_draft(to_addr, subject, html_body, folder_path):
    allowed_exts = ['.pdf', '.pptx', '.xlsx']
    files = [f for f in os.listdir(folder_path)
             if os.path.isfile(os.path.join(folder_path, f)) and
             os.path.splitext(f)[1].lower() in allowed_exts]

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.HTMLBody = html_body
    # mail.SentOnBehalfOfName = from_addr

    recipient = mail.Recipients.Add(to_addr)
    recipient.Type = 1
    mail.Recipients.ResolveAll()

    for file in files:
        mail.Attachments.Add(os.path.join(folder_path, file))

    mail.Save()

def main(csv_path):
    #from_addr = "jeonghin.chin@aa.com"
    subject = "Documents Attached"
    html_body = "<p>Dear recipient,<br>Please find the attached documents.</p>"

    with open(csv_path, mode='r', encoding='utf-16') as f:
        reader = csv.DictReader(f)
        for row in reader:
            receiver = row['Receiver']
            folder = row['Folder']
            if os.path.isdir(folder):
                print(f"Creating draft for {receiver} from folder {folder}")
                #save_draft(from_addr, receiver, subject, html_body, folder)
                save_draft(receiver, subject, html_body, folder)
            else:
                print(f"Skipped: folder not found for {receiver} -> {folder}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python email_folder_batch.py <csv_file>")
    else:
        main(sys.argv[1])
