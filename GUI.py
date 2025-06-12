import tkinter as tk
from tkinter import filedialog, messagebox

def upload_financial(): pass
def upload_template(): pass

def submit():
    print("Prompt:", entry.get().strip())

root = tk.Tk()
root.title("Financial Analysis")
root.geometry("400x250")

tk.Button(root, text="Upload Financial PDF/Excel", command=upload_financial).pack(pady=8)
tk.Button(root, text="Upload Excel Template",   command=upload_template).pack(pady=8)

# ← this Entry will always show up, even in dark mode
entry = tk.Entry(root, width=50, bd=2, relief="groove")
entry.pack(pady=12)
entry.insert(0, "Type your prompt here…")

tk.Button(root, text="Submit", command=submit).pack(pady=12)

root.mainloop()