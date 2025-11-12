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

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import win32com.client
import threading
import getpass
import socket

class EmailFolderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Folder Attachment App")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_folder = tk.StringVar()
        self.recipient_email = tk.StringVar()
        self.email_subject = tk.StringVar(value="Documents Attached")
        self.user_email = tk.StringVar()
        
        # Get current user's email address
        self.get_user_email()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Email Folder Attachment App", 
                               style="Title.TLabel", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Folder selection
        ttk.Label(main_frame, text="Select Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        folder_frame = ttk.Frame(main_frame)
        folder_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        folder_frame.columnconfigure(0, weight=1)
        
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.selected_folder, 
                                     state="readonly", width=50)
        self.folder_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(folder_frame, text="Browse", 
                  command=self.browse_folder).grid(row=0, column=1)
        
        # Current user info
        ttk.Label(main_frame, text="Your Email:").grid(row=2, column=0, sticky=tk.W, pady=5)
        user_frame = ttk.Frame(main_frame)
        user_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        user_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(user_frame, textvariable=self.user_email, 
                 state="readonly", width=40).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(user_frame, text="Send to Self", 
                  command=self.send_to_self).grid(row=0, column=1)
        
        # Recipient email
        ttk.Label(main_frame, text="Recipient Email:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.recipient_email, 
                 width=50).grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Email subject
        ttk.Label(main_frame, text="Email Subject:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.email_subject, 
                 width=50).grid(row=4, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Email body
        ttk.Label(main_frame, text="Email Body:").grid(row=5, column=0, sticky=(tk.W, tk.N), pady=5)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.email_body = tk.Text(text_frame, height=8, width=50, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.email_body.yview)
        self.email_body.configure(yscrollcommand=scrollbar.set)
        
        self.email_body.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Default email body
        default_body = "Dear recipient,\n\nPlease find the attached documents.\n\nBest regards"
        self.email_body.insert(tk.END, default_body)
        
        # File list display
        ttk.Label(main_frame, text="Files to attach:").grid(row=6, column=0, sticky=(tk.W, tk.N), pady=(20, 5))
        
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=6, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 5))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        self.file_listbox = tk.Listbox(list_frame, height=6)
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=list_scrollbar.set)
        
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=(20, 0))
        
        self.create_draft_btn = ttk.Button(button_frame, text="Create Email Draft", 
                                          command=self.create_email_draft, style="Accent.TButton")
        self.create_draft_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Clear All", 
                  command=self.clear_all).pack(side=tk.LEFT)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready", foreground="green")
        self.status_label.grid(row=8, column=0, columnspan=3, pady=(10, 0))
        
        # Configure grid weights for main_frame
        main_frame.rowconfigure(5, weight=1)  # Email body text widget
        main_frame.rowconfigure(6, weight=1)  # File list
        
    def get_user_email(self):
        """Retrieve the current user's email address from Outlook"""
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Try to get the default account's email address
            accounts = namespace.Accounts
            if accounts.Count > 0:
                # Get the first (usually default) account
                default_account = accounts.Item(1)
                email = default_account.SmtpAddress
                if email:
                    self.user_email.set(email)
                    return email
            
            # Fallback: try to get from current user profile
            current_user = namespace.CurrentUser
            if current_user and hasattr(current_user, 'Address'):
                email = current_user.Address
                if email and '@' in email:
                    self.user_email.set(email)
                    return email
            
            # Try to get email from Exchange properties
            try:
                current_user = namespace.CurrentUser
                if current_user:
                    address_entry = current_user.AddressEntry
                    if address_entry:
                        exchange_user = address_entry.GetExchangeUser()
                        if exchange_user and exchange_user.PrimarySmtpAddress:
                            email = exchange_user.PrimarySmtpAddress
                            self.user_email.set(email)
                            return email
            except:
                pass
                    
            # Another fallback: construct from username and try to detect domain
            username = getpass.getuser()
            domain = self.detect_company_domain()
            fallback_email = f"{username}@{domain}"
            self.user_email.set(fallback_email)
            return fallback_email
            
        except Exception as e:
            print(f"Could not retrieve user email: {e}")
            # Final fallback
            username = getpass.getuser()
            domain = self.detect_company_domain()
            fallback_email = f"{username}@{domain}"
            self.user_email.set(fallback_email)
            return fallback_email
    
    def detect_company_domain(self):
        """Try to detect company domain from environment or return common default"""
        try:
            # Try to get domain from environment variables
            import socket
            
            # Try USERDNSDOMAIN environment variable (common in corporate environments)
            user_domain = os.environ.get('USERDNSDOMAIN', '').lower()
            if user_domain and '.' in user_domain:
                return user_domain
            
            # Try to get from computer's domain
            computer_name = os.environ.get('COMPUTERNAME', '')
            domain_name = os.environ.get('USERDOMAIN', '')
            
            # Try to get FQDN
            try:
                fqdn = socket.getfqdn()
                if '.' in fqdn and not fqdn.endswith('.local'):
                    # Extract domain from FQDN
                    parts = fqdn.split('.')
                    if len(parts) > 1:
                        domain = '.'.join(parts[1:])
                        return domain
            except:
                pass
            
            # Common corporate domain patterns - you can customize this
            if 'aa.com' in str(os.environ.get('LOGONSERVER', '')).lower():
                return 'aa.com'  # American Airlines domain as an example
            
            # Default fallback
            return 'aa.com'
            
        except Exception as e:
            print(f"Could not detect domain: {e}")
            return 'aa.com'
        
    def browse_folder(self):
        folder_path = filedialog.askdirectory(title="Select folder containing files to attach")
        if folder_path:
            self.selected_folder.set(folder_path)
            self.update_file_list()
            
    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        folder_path = self.selected_folder.get()
        
        if not folder_path or not os.path.isdir(folder_path):
            return
            
        # Get files with allowed extensions (same as your original code)
        allowed_exts = ['.pdf', '.pptx', '.xlsx']
        files = [f for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f)) and
                os.path.splitext(f)[1].lower() in allowed_exts]
        
        if files:
            for file in sorted(files):
                self.file_listbox.insert(tk.END, file)
            self.status_label.config(text=f"Found {len(files)} files to attach", foreground="blue")
        else:
            self.file_listbox.insert(tk.END, "No PDF, PPTX, or XLSX files found")
            self.status_label.config(text="No attachable files found in selected folder", foreground="orange")
    
    def send_to_self(self):
        """Set the recipient email to the current user's email"""
        self.recipient_email.set(self.user_email.get())
    
    def create_email_draft(self):
        # Validate inputs
        if not self.selected_folder.get():
            messagebox.showerror("Error", "Please select a folder first")
            return
            
        if not self.recipient_email.get().strip():
            messagebox.showerror("Error", "Please enter recipient email address")
            return
            
        if not self.email_subject.get().strip():
            messagebox.showerror("Error", "Please enter email subject")
            return
        
        # Disable button and show status
        self.create_draft_btn.config(state="disabled")
        self.status_label.config(text="Creating email draft...", foreground="blue")
        self.root.update()
        
        # Run in separate thread to prevent UI freezing
        thread = threading.Thread(target=self._create_draft_worker)
        thread.daemon = True
        thread.start()
    
    def _create_draft_worker(self):
        try:
            folder_path = self.selected_folder.get()
            recipient = self.recipient_email.get().strip()
            subject = self.email_subject.get().strip()
            body_text = self.email_body.get("1.0", tk.END).strip()
            
            # Convert plain text to HTML
            html_body = body_text.replace('\n', '<br>')
            html_body = f"<div style='font-family: Arial, sans-serif;'>{html_body}</div>"
            
            # Get files to attach (same logic as your original code)
            allowed_exts = ['.pdf', '.pptx', '.xlsx']
            files = [f for f in os.listdir(folder_path)
                    if os.path.isfile(os.path.join(folder_path, f)) and
                    os.path.splitext(f)[1].lower() in allowed_exts]
            
            if not files:
                self.root.after(0, lambda: self._update_status("No files to attach found", "orange"))
                return
            
            # Create Outlook draft (same logic as your save_draft function)
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.HTMLBody = html_body
            
            recipient_obj = mail.Recipients.Add(recipient)
            recipient_obj.Type = 1
            mail.Recipients.ResolveAll()
            
            # Add attachments
            for file in files:
                file_path = os.path.join(folder_path, file)
                mail.Attachments.Add(file_path)
            
            # Save as draft
            mail.Save()
            
            # Update UI on main thread
            self.root.after(0, lambda: self._update_status(
                f"Email draft created successfully with {len(files)} attachments!", "green"))
            
        except Exception as e:
            error_msg = f"Error creating email draft: {str(e)}"
            self.root.after(0, lambda: self._update_status(error_msg, "red"))
        finally:
            # Re-enable button on main thread
            self.root.after(0, lambda: self.create_draft_btn.config(state="normal"))
    
    def _update_status(self, message, color):
        self.status_label.config(text=message, foreground=color)
        
    def clear_all(self):
        self.selected_folder.set("")
        self.recipient_email.set("")
        self.email_subject.set("Documents Attached")
        self.email_body.delete("1.0", tk.END)
        self.email_body.insert(tk.END, "Dear recipient,\n\nPlease find the attached documents.\n\nBest regards")
        self.file_listbox.delete(0, tk.END)
        self.status_label.config(text="Ready", foreground="green")
        # Note: We don't clear user_email as it's auto-detected

def main():
    root = tk.Tk()
    app = EmailFolderApp(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()