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
        self.extracted_data_list = []  # Stores all GPT results

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
        self.text_edit.setPlaceholderText('Include the year in your prompt, e.g., \"Extract all financial fields for 2024...\"")
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

        # Extraction step
        extractor = os.path.join(os.path.dirname(__file__), 'extract.py')
        try:
            res = subprocess.run(
                [sys.executable, extractor, self.fin_file],
                capture_output=True, text=True, check=True
            )
            extracted_data = json.loads(res.stdout)
        except Exception as e:
            QMessageBox.critical(self, 'Extraction Error', str(e))
            return

        # GPT step
        gpt_script = os.path.join(os.path.dirname(__file__), 'GPT.py')
        try:
            gpt_proc = subprocess.run([
                sys.executable, gpt_script,
                '--data', '-',  # Pass data via stdin
                '--prompt', prompt,
                '--key', self.api_key
            ], input=json.dumps(extracted_data), capture_output=True, text=True, check=True)
            structured_data = json.loads(gpt_proc.stdout)
            self.extracted_data_list.append(structured_data)  # Store in memory, always append!
        except Exception as e:
            QMessageBox.critical(self, 'GPT Error', str(e))
            return
        except json.JSONDecodeError as e:
            QMessageBox.critical(self, 'GPT Output Error', f'Output is not valid JSON:\n{e}\n{gpt_proc.stdout}')
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
            # Let the user select and submit another file
        else:
            # Debug print: check what is in the extracted data list
            print("Extracted data list:", self.extracted_data_list)
            # Export and open
            output_excel = self.tpl_file  # Overwrite template
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
            # Optionally reset for a new batch
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