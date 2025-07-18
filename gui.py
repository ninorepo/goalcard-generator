import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import subprocess

def browse_input_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        input_file_var.set(filename)

def browse_output_file():
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filename:
        output_file_var.set(filename)

def run_script():
    input_file = input_file_var.get()
    sheet_name = sheet_name_var.get()
    line_code = line_code_var.get()
    product_code = product_code_var.get()
    target = target_var.get()
    date_string = date_var.get()  # Updated to use date_var
    output_file = output_file_var.get()

    if not all([input_file, sheet_name, line_code, product_code, target, date_string, output_file]):
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    try:
        target_int = int(target)
    except ValueError:
        messagebox.showerror("Error", "Target must be an integer.")
        return

    cmd = [
        "python", "goalcard-generator.py",
        input_file, sheet_name, line_code, product_code,
        str(target_int), date_string, output_file
    ]

    try:
        subprocess.run(cmd, check=True)
        messagebox.showinfo("Success", "Script executed successfully.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Script failed:\n{e}")

# === GUI SETUP ===
root = tk.Tk()
root.title("Goalcard Generator by Nino")
root.geometry("600x350")

input_file_var = tk.StringVar()
sheet_name_var = tk.StringVar()
line_code_var = tk.StringVar()
product_code_var = tk.StringVar()
target_var = tk.StringVar()
date_var = tk.StringVar()
output_file_var = tk.StringVar()

# Labels and fields
fields = [
    ("Eng. Sheet", input_file_var, browse_input_file),
    ("Sheet Name", sheet_name_var, None),
    ("Floor", line_code_var, None),
    ("CMT", product_code_var, None),
    ("Target/hr", target_var, None),
    ("Save to", output_file_var, browse_output_file),
]

for i, (label, var, browse_fn) in enumerate(fields):
    tk.Label(root, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=5)
    entry = tk.Entry(root, textvariable=var, width=50)
    entry.grid(row=i, column=1, padx=5)
    if browse_fn:
        tk.Button(root, text="Browse", command=browse_fn).grid(row=i, column=2, padx=5)

# Add calendar field for Date
tk.Label(root, text="Date").grid(row=len(fields), column=0, sticky="w", padx=5, pady=5)
date_entry = DateEntry(root, textvariable=date_var, date_pattern="yyyy-mm-dd", width=47)
date_entry.grid(row=len(fields), column=1, padx=5, pady=5)

# Run button
tk.Button(root, text="Run Script", command=run_script, bg="green", fg="white").grid(
    row=len(fields)+1, column=0, columnspan=3, pady=15
)

root.mainloop()
