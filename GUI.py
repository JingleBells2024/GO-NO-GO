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
        self.root.minsize(600, 500)

        # ← This Frame is your true black background
        container = tk.Frame(root, bg="black")
        container.pack(fill="both", expand=True)

        # Title at the top
        tk.Label(
            container,
            text="Financial Analysis",
            font=("Arial", 18, "bold"),
            bg="black",
            fg="white"
        ).pack(pady=(10, 15))

        self.input_file = None
        self.template_file = None

        # Upload Financial File
        tk.Button(
            container,
            text="Upload Financial PDF/Excel",
            command=self.upload_input
        ).pack(pady=5)
        self.input_label = tk.Label(
            container, text="No file uploaded", fg="gray", bg="black"
        )
        self.input_label.pack()

        # Upload Template
        tk.Button(
            container,
            text="Upload Excel Template",
            command=self.upload_template
        ).pack(pady=5)
        self.template_label = tk.Label(
            container, text="No template uploaded", fg="gray", bg="black"
        )
        self.template_label.pack()

        # Prompt Input: always white so you can see it
        self.prompt_entry = tk.Text(
            container,
            height=5,
            width=60,
            bg="white",
            fg="black",
            insertbackground="black"
        )
        self.prompt_entry.pack(pady=(20, 5))
        self.prompt_entry.insert("1.0", "Enter prompt for ChatGPT here...")

        # Submit Button
        tk.Button(
            container,
            text="Submit",
            command=self.submit
        ).pack(pady=10)

    def upload_input(self):
        file = filedialog.askopenfilename(title="Choose Financial File")
        if file:
            self.input_file = file
            self.input_label.config(text=os.path.basename(file), fg="white")

    def upload_template(self):
        file = filedialog.askopenfilename(
            title="Choose Excel Template",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file:
            self.template_file = file
            self.template_label.config(text=os.path.basename(file), fg="white")

    def submit(self):
        if not self.input_file or not self.template_file:
            messagebox.showerror("Error", "Please upload both a financial file and a template.")
            return

        prompt = self.prompt_entry.get("1.0", tk.END).strip()
        if not prompt or prompt == "Enter prompt for ChatGPT here...":
            messagebox.showerror("Error", "Please enter a prompt to guide ChatGPT.")
            return

        try:
            # Simulated ChatGPT result
            extracted_data = {
                "Revenue": 964021.78,
                "COGS": 156873.40,
                "Net Profit": 169329.63
            }

            output_path = os.path.join(
                os.path.dirname(self.template_file),
                "filled_output.xlsx"
            )
            shutil.copy(self.template_file, output_path)
            wb = load_workbook(output_path)
            ws = wb.active

            ws["J5"] = extracted_data["Revenue"]
            ws["J6"] = extracted_data["COGS"]
            ws["J13"] = extracted_data["Net Profit"]
            wb.save(output_path)

            open_now = messagebox.askyesno(
                "Done",
                "Template filled and saved as 'filled_output.xlsx'.\nOpen now?"
            )
            if open_now:
                subprocess.call(["open", output_path])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fill template.\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    FinancialFillerApp(root)
    root.mainloop()