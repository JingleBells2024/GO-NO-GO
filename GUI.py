import tkinter as tk

def upload_financial():
    # placeholder for file‐upload logic
    pass

def upload_template():
    # placeholder for file‐upload logic
    pass

def submit():
    # placeholder for submit action
    print("Submit clicked. Text:", text_box.get("1.0", tk.END).strip())

root = tk.Tk()
root.title("Financial Analysis")
root.geometry("400x300")

# Buttons
tk.Button(root, text="Upload Financial PDF/Excel", command=upload_financial).pack(pady=10)
tk.Button(root, text="Upload Excel Template",   command=upload_template).pack(pady=10)

# Text box
text_box = tk.Text(root, height=5, width=40)
text_box.pack(pady=10)

# Submit
tk.Button(root, text="Submit", command=submit).pack(pady=10)

root.mainloop()