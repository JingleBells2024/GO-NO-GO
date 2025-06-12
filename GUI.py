import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import shutil
import os
from openpyxl import load_workbook

class FinancialFillerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Analysis")
        self.root.configure(bg="white")
        self.root.minsize(600, 500)

        # Container so everything shares the white background
        container = tk.Frame(root, bg="white")
        container.pack(fill="both", expand=True)

        # Title
        tk.Label(
            container,
            text="Financial Analysis",
            font=("Arial", 18, "bold"),
            bg="white",
            fg="black"
        ).pack(pady=(10, 15))

        self.input_file = None
        self.template_file = None

        # Upload Financial File
        tk.Button(
            container,
            text="Upload Financial PDF/Excel",
            command=self.upload_input,
            bg="#e0e0e0",
            fg="black"
        ).pack(pady=5)
        self.input_label = tk.Label(
            container,
            text="No file uploaded",
            fg="gray",
            bg="white"
        )
        self.input_label.pack()

        # Upload Template
        tk.Button(
            container,
            text="Upload Excel Template",
            command=self.upload_template,
            bg="#e0e0e0",
            fg="black"
        ).pack(pady=5)
        self.template_label = tk.Label(
            container,
            text="No template uploaded",
            fg="gray",
            bg="white"
        )
        self.template_label.pack()

        # ★ THE VISIBLE TEXT BOX ★
        self.prompt_entry = tk.Text(
            container,
            height=5,
            width=60,
            bg="white",
            fg="black",
            insertbackground="black",
            bd=2,
            relief="solid"            # solid border
        )
        self.prompt_entry.pack(pady=(20, 5))
        self.prompt_entry.insert("1.0", "Enter prompt for ChatGPT here...")

        # Submit
        tk.Button(
            container,
            text="Submit",
            command=self.submit,
            bg="#e0e0e0",
            fg="black"
        ).pack(pady=10)

    def upload_input(self):
        file = filedialog.askopenfilename(title="Choose Financial File")
        if file:
            self.input_file = file
            self.input_label.config(text=os.path.basename(file), fg="black")

    def upload_template(self):
        file = filedialog.askopenfilename(
            title="Choose Excel Template",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file:
            self.template_file = file
            self.template_label.config(text=os.path.basename(file), fg="black")

    def submit(self):
        if not self.input_file or not self.template_file:
            messagebox.showerror("Error", "Please upload both a financial file and a template.")
            return

        prompt = self.prompt_entry.get("1.0", tk.END).strip()
        if not prompt or prompt == "Enter prompt for ChatGPT here...":
            messagebox.showerror("Error", "Please enter a prompt to guide ChatGPT.")
            return

        try:
            # Simulated ChatGPT output
            data = {"Revenue": 964021.78, "COGS": 156873.40, "Net Profit": 169329.63}

            out = os.path.join(os.path.dirname(self.template_file), "filled_output.xlsx")
            shutil.copy(self.template_file, out)
            wb = load_workbook(out)
            ws = wb.active
            ws["J5"], ws["J6"], ws["J13"] = data["Revenue"], data["COGS"], data["Net Profit"]
            wb.save(out)

            if messagebox.askyesno("Done", f"Saved as 'filled_output.xlsx'.\nOpen now?"):
                subprocess.call(["open", out])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fill template:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    FinancialFillerApp(root)
    root.mainloop()