from bulkTerms_Confluence import (
    verify_rest_connection,
    main,
    export_glossary_to_csv
)

import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import threading
import tkinter.scrolledtext as scrolledtext
import sys
from io import StringIO


# ----- UI Styling -----
BG_COLOR = "#7A57DD"
FG_COLOR = "white"
FONT = ("Segoe UI", 12)
ENTRY_BG = "#F0F0F0"
BUTTON_BG = "#362499"
BUTTON_FG = "white"


# ----- Functions -----
def toggle_cloud_inputs():
    if cloud_var.get():
        email_label.grid()
        email_entry.grid()
        token_label.config(text="API Token:")
        email_entry.config(state="normal")
    else:
        email_label.grid_remove()
        email_entry.grid_remove()
        token_label.config(text="PAT:")
        email_entry.config(state="disabled")

def browse_csv():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    csv_entry.delete(0, tk.END)
    csv_entry.insert(0, file_path)

def test_connection():
    cloud_val = cloud_var.get()
    token = token_entry.get()
    email = email_entry.get()

    if not token or (cloud_val and not email):
        messagebox.showerror("Missing Info", "Please fill in all required fields for the connection test.")
        return

    try:
        result = verify_rest_connection(cloud=cloud_val, email=email, token=token)
        if result:
            messagebox.showinfo("Success", "Connection successful!")
        else:
            messagebox.showerror("Failure", "Connection failed. Check your credentials and try again.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def run_upload_and_show_output():
    cloud_val = cloud_var.get()
    token = token_entry.get()
    email = email_entry.get()
    csv_path = csv_entry.get()

    if not token or not csv_path or (cloud_val and not email):
        messagebox.showerror("Missing Info", "Please fill in all required fields.")
        return

    os.environ["GLOSSARY_CLOUD"] = str(cloud_val)
    os.environ["GLOSSARY_EMAIL"] = email
    os.environ["GLOSSARY_TOKEN"] = token
    os.environ["GLOSSARY_CSV"] = csv_path

    # Disable buttons during upload
    upload_btn.config(state="disabled")
    test_btn.config(state="disabled")

    # Create output window
    output_win = tk.Toplevel(root)
    output_win.title("Upload Output")
    output_text = scrolledtext.ScrolledText(output_win, width=80, height=20)
    output_text.pack(fill=tk.BOTH, expand=True)

    def run_process():
        old_stdout = sys.stdout
        sys.stdout = mystdout = StringIO()

        try:
            main(cloud_val, email, token, csv_path)
        except Exception as e:
            print(f"Error during upload: {e}")
        finally:
            sys.stdout = old_stdout

        output_text.insert(tk.END, mystdout.getvalue())
        output_text.see(tk.END)

        # Re-enable buttons
        upload_btn.config(state="normal")
        test_btn.config(state="normal")

        if "Created page:" in mystdout.getvalue():
            messagebox.showinfo("Upload Complete", "Glossary terms uploaded successfully!")
        else:
            messagebox.showerror("Upload Failed", "Upload failed. See output window for details.")


    threading.Thread(target=run_process, daemon=True).start()

def export_glossary():
    cloud_val = cloud_var.get()
    token = token_entry.get()
    email = email_entry.get()

    if not token or (cloud_val and not email):
        messagebox.showerror("Missing Info", "Please fill in all required fields.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return  # User cancelled

    # Disable buttons during export
    upload_btn.config(state="disabled")
    test_btn.config(state="disabled")
    export_btn.config(state="disabled")

    output_win = tk.Toplevel(root)
    output_win.title("Export Output")
    output_text = scrolledtext.ScrolledText(output_win, width=80, height=20)
    output_text.pack(fill=tk.BOTH, expand=True)

    def run_export():
        old_stdout = sys.stdout
        sys.stdout = mystdout = StringIO()

        try:
            export_glossary_to_csv(cloud_val, email, token, file_path)
        except Exception as e:
            print(f"Error during export: {e}")
        finally:
            sys.stdout = old_stdout

        output_text.insert(tk.END, mystdout.getvalue())
        output_text.see(tk.END)

        # Re-enable buttons
        upload_btn.config(state="normal")
        test_btn.config(state="normal")
        export_btn.config(state="normal")

        if "Export complete." in mystdout.getvalue():
            messagebox.showinfo("Export Complete", f"Glossary exported to:\n{file_path}")
        else:
            messagebox.showerror("Export Failed", "Export failed. See output window for details.")

    threading.Thread(target=run_export, daemon=True).start()


# ----- UI Window -----
root = tk.Tk()
root.title("Glossary Page Uploader")
root.configure(bg=BG_COLOR)
root.geometry("600x400")

cloud_var = tk.BooleanVar(value=True)
tk.Checkbutton(root, text="Use Cloud", variable=cloud_var, command=toggle_cloud_inputs,
               bg=BG_COLOR, fg=FG_COLOR, font=FONT, selectcolor=BG_COLOR).grid(row=0, column=1, sticky="w", pady=(10, 5))

email_label = tk.Label(root, text="Email:", font=FONT, bg=BG_COLOR, fg=FG_COLOR)
email_label.grid(row=1, column=0, sticky="e", padx=10, pady=5)
email_entry = tk.Entry(root, width=40, font=FONT, bg=ENTRY_BG)
email_entry.grid(row=1, column=1, padx=10)

token_label = tk.Label(root, text="API Token:", font=FONT, bg=BG_COLOR, fg=FG_COLOR)
token_label.grid(row=2, column=0, sticky="e", padx=10, pady=5)
token_entry = tk.Entry(root, width=40, font=FONT, bg=ENTRY_BG, show="*")
token_entry.grid(row=2, column=1, padx=10)

tk.Label(root, text="CSV File:", font=FONT, bg=BG_COLOR, fg=FG_COLOR).grid(row=3, column=0, sticky="e", padx=10, pady=5)
csv_entry = tk.Entry(root, width=40, font=FONT, bg=ENTRY_BG)
csv_entry.grid(row=3, column=1, padx=10)
tk.Button(root, text="Browse", command=browse_csv, font=FONT,
          bg=BUTTON_BG, fg=BUTTON_FG).grid(row=3, column=2, padx=5)

test_btn = tk.Button(root, text="Test Credentials", command=test_connection, font=FONT,
                     bg=BUTTON_BG, fg=BUTTON_FG)
test_btn.grid(row=4, column=1, pady=(10, 0))

upload_btn = tk.Button(root, text="Upload Glossary Terms", command=run_upload_and_show_output, font=FONT,
                       bg=BUTTON_BG, fg=BUTTON_FG)
upload_btn.grid(row=5, column=1, pady=(10, 15))

export_btn = tk.Button(root, text="Export Glossary to CSV", command=export_glossary, font=FONT,
                       bg=BUTTON_BG, fg=BUTTON_FG)
export_btn.grid(row=6, column=1, pady=(5, 15))

toggle_cloud_inputs()

root.mainloop()
