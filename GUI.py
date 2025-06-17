import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QTextEdit, QFileDialog, QMessageBox, QLabel, QLineEdit
)
from compiler import map_to_excel  # Import as a function

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Financial Analysis')
        self.resize(600, 400)

        self.fin_file = None
        self.tpl_file = None
        self.api_key = None
        self.extracted_data_list = []

        layout = QVBoxLayout()

        btn1 = QPushButton('Upload Financial PDF/Excel')
        btn1.clicked.connect(self.upload_financial)
        layout.addWidget(btn1)
        self.fin_label = QLabel('No financial file selected')
        layout.addWidget(self.fin_label)

        btn2 = QPushButton('Upload Excel Template')
        btn2.clicked.connect(self.upload_template)
        layout.addWidget(btn2)
        self.tpl_label = QLabel('No template selected')
        layout.addWidget(self.tpl_label)

        api_label = QLabel('Enter API Key:')
        layout.addWidget(api_label)
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.Password)
        self.api_input.setPlaceholderText('sk-...')
        layout.addWidget(self.api_input)

        prompt_label = QLabel('GPT Prompt:')
        layout.addWidget(prompt_label)
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Include the year in your prompt, e.g., \"Extract all financial fields for 2024...\"")
        layout.addWidget(self.text_edit)

        submit_btn = QPushButton('Submit')
        submit_btn.clicked.connect(self.submit)
        layout.addWidget(submit_btn)

        self.setLayout(layout)

    def upload_financial(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Choose Financial File',
            filter='PDF Files (*.pdf);;Excel Files (*.xlsx *.xlsm)'
        )
        if path:
            self.fin_file = path
            self.fin_label.setText(os.path.basename(path))
            self.fin_label.setStyleSheet('color: green')
        else:
            self.fin_file = None
            self.fin_label.setText('No financial file selected')
            self.fin_label.setStyleSheet('color: red')

    def upload_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Choose Excel Template',
            filter='Excel Files (*.xlsx *.xlsm)'
        )
        if path:
            self.tpl_file = path
            self.tpl_label.setText(os.path.basename(path))
            self.tpl_label.setStyleSheet('color: green')
        else:
            self.tpl_file = None
            self.tpl_label.setText('No template selected')
            self.tpl_label.setStyleSheet('color: red')

    def submit(self):
        prompt = self.text_edit.toPlainText().strip()
        self.api_key = self.api_input.text().strip()

        if not self.api_key:
            QMessageBox.warning(self, 'Error', 'Enter API key.')
            return
        if not self.fin_file:
            QMessageBox.warning(self, 'Error', 'Upload a financial file.')
            return
        if not self.tpl_file:
            QMessageBox.warning(self, 'Error', 'Upload a template.')
            return
        if not prompt:
            QMessageBox.warning(self, 'Error', 'Enter a prompt.')
            return

        # Extraction step: always output to extracted_data.json
        extractor = os.path.join(os.path.dirname(__file__), 'extract.py')
        extracted_json_path = os.path.join(os.getcwd(), 'extracted_data.json')
        try:
            subprocess.run(
                [sys.executable, extractor, self.fin_file, "-o", extracted_json_path],
                check=True, capture_output=True, text=True
            )
            if not os.path.isfile(extracted_json_path):
                raise Exception("Extracted file not created.")
            with open(extracted_json_path) as f:
                extracted_data = json.load(f)
            if not extracted_data:
                raise Exception("Extracted data is empty.")
        except Exception as e:
            QMessageBox.critical(self, 'Extraction Error', str(e))
            return

        # Save prompt as .txt and data as .json, send both to GPT script
        gpt_script = os.path.join(os.path.dirname(__file__), 'GPT.py')
        gpt_output_path = os.path.join(os.getcwd(), 'GPT_output.json')
        prompt_txt_path = os.path.join(os.getcwd(), 'user_prompt.txt')
        with open(prompt_txt_path, "w") as f:
            f.write(prompt)

        try:
            gpt_proc = subprocess.run([
                sys.executable, gpt_script,
                "--data", extracted_json_path,
                "--prompt", prompt_txt_path,
                "--key", self.api_key
            ], capture_output=True, text=True, check=True)
            # GPT.py should output JSON to GPT_output.json
            if not os.path.isfile(gpt_output_path):
                raise Exception("GPT did not produce output.")
            with open(gpt_output_path) as f:
                structured_data = json.load(f)
            self.extracted_data_list.append(structured_data)
        except Exception as e:
            QMessageBox.critical(self, 'GPT Error', str(e))
            return
        except json.JSONDecodeError as e:
            QMessageBox.critical(self, 'GPT Output Error', f'Output is not valid JSON:\n{e}')
            return

        # Prompt user: upload more or open doc
        reply = QMessageBox.question(
            self,
            'Batch Finished',
            "File processed. Would you like to upload more financial files?\n\n"
            "Choose 'Yes' to process another file, or 'No' to export and open the completed Excel document.",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.fin_file = None
            self.fin_label.setText('No financial file selected')
        else:
            print("Extracted data list:", self.extracted_data_list)
            output_excel = self.tpl_file
            try:
                map_to_excel(self.extracted_data_list, self.tpl_file, None)
            except Exception as e:
                QMessageBox.critical(self, 'Compile Error', str(e))
                return
            open_reply = QMessageBox.question(
                self,
                'Processing complete',
                f'Excel file updated: {output_excel}\n\nOpen now?',
                QMessageBox.Yes | QMessageBox.No
            )
            if open_reply == QMessageBox.Yes:
                self.open_file(output_excel)
            self.extracted_data_list = []

    def open_file(self, filepath):
        if platform.system() == 'Darwin':       # macOS
            subprocess.run(['open', filepath])
        elif platform.system() == 'Windows':    # Windows
            os.startfile(filepath)
        else:                                   # Linux and others
            subprocess.run(['xdg-open', filepath])

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = FinancialAnalysis()
    win.show()
    sys.exit(app.exec_())