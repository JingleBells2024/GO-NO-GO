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

        # Title label at the top
        tk.Label(root, text="Financial Analysis", font=("Arial", 18, "bold")).pack(pady=(10, 15))

        self.input_file = None
        self.template_file = None

        # Upload Financial File
        tk.Button(root, text="Upload Financial PDF/Excel", command=self.upload_input).pack(pady=5)
        self.input_label = tk.Label(root, text="No file uploaded", fg="gray")
        self.input_label.pack()

        # Upload Template
        tk.Button(root, text="Upload Excel Template", command=self.upload_template).pack(pady=5)
        self.template_label = tk.Label(root, text="No template uploaded", fg="gray")
        self.template_label.pack()

        # Prompt Input
        self.prompt_entry = tk.Text(root, height=5, width=60)
        self.prompt_entry.pack(pady=(20, 5))
        self.prompt_entry.insert("1.0", "Enter prompt for ChatGPT here...")

        # Submit Button
        tk.Button(root, text="Submit", command=self.submit).pack(pady=10)

    def upload_input(self):
        file = filedialog.askopenfilename(title="Choose Financial File")
        if file:
            self.input_file = file
            self.input_label.config(text=os.path.basename(file), fg="black")

    def upload_template(self):
        file = filedialog.askopenfilename(title="Choose Excel Template", filetypes=[("Excel files", "*.xlsx")])
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
            # Simulated result from ChatGPT
            extracted_data = {
                "Revenue": 964021.78,
                "COGS": 156873.40,
                "Net Profit": 169329.63
            }

            # Save new Excel
            output_path = os.path.join(os.path.dirname(self.template_file), "filled_output.xlsx")
            shutil.copy(self.template_file, output_path)
            wb = load_workbook(output_path)
            ws = wb.active

            # Example cell filling
            ws["J5"] = extracted_data["Revenue"]
            ws["J6"] = extracted_data["COGS"]
            ws["J13"] = extracted_data["Net Profit"]
            wb.save(output_path)

            # Ask to open the file
            open_now = messagebox.askyesno("Done", "Template filled and saved as 'filled_output.xlsx'.\nOpen now?")
            if open_now:
                subprocess.call(["open", output_path])  # macOS: open with default app

        except Exception as e:
            messagebox.showerror("Error", f"Failed to fill template.\n{str(e)}")

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialFillerApp(root)
    root.mainloop()