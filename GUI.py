import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import shutil
import os

class FinancialFillerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GO-NO-GO Financial Filler")
        self.input_file = None
        self.template_file = None

        # File upload buttons
        tk.Button(root, text="Upload Financial PDF/Excel", command=self.upload_input).pack(pady=5)
        self.input_label = tk.Label(root, text="No file uploaded", fg="gray")
        self.input_label.pack()

        tk.Button(root, text="Upload Excel Template", command=self.upload_template).pack(pady=5)
        self.template_label = tk.Label(root, text="No template uploaded", fg="gray")
        self.template_label.pack()

        # Prompt input
        tk.Label(root, text="Enter GPT prompt to guide data extraction:").pack(pady=(10, 0))
        self.prompt_entry = tk.Text(root, height=5, width=60)
        self.prompt_entry.pack()

        # Fill button
        tk.Button(root, text="Fill Template", command=self.fill_template).pack(pady=10)

        # Status label
        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.pack()

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

    def fill_template(self):
        if not self.input_file or not self.template_file:
            messagebox.showerror("Error", "Please upload both financial and template files.")
            return

        prompt = self.prompt_entry.get("1.0", tk.END).strip()
        if not prompt:
            messagebox.showerror("Error", "Please enter a prompt to guide ChatGPT.")
            return

        try:
            self.status_label.config(text="Processing...")

            # Simulate ChatGPT output (you'll replace this with real GPT call)
            extracted_data = {
                "Revenue": 964021.78,
                "COGS": 156873.40,
                "Net Profit": 169329.63
            }

            # Fill the template
            output_path = os.path.join(os.path.dirname(self.template_file), "filled_output.xlsx")
            shutil.copy(self.template_file, output_path)
            wb = load_workbook(output_path)
            ws = wb.active

            # Example insert – update with your actual cell map
            ws["J5"] = extracted_data["Revenue"]
            ws["J6"] = extracted_data["COGS"]
            ws["J13"] = extracted_data["Net Profit"]
            wb.save(output_path)

            self.status_label.config(text=f"Template filled successfully → {output_path}", fg="green")
            messagebox.showinfo("Success", f"Saved to:\n{output_path}")
        except Exception as e:
            self.status_label.config(text="Failed to fill template.", fg="red")
            messagebox.showerror("Error", str(e))

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialFillerApp(root)
    root.mainloop()