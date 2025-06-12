import sys
import os
import json
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QTextEdit, QFileDialog, QMessageBox, QLabel, QLineEdit
)

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Financial Analysis')
        self.resize(600, 550)

        self.fin_file = None
        self.tpl_file = None
        self.api_key = None

        layout = QVBoxLayout()

        # Upload Financial File
        btn1 = QPushButton('Upload Financial PDF/Excel')
        btn1.clicked.connect(self.upload_financial)
        layout.addWidget(btn1)
        self.fin_label = QLabel('No financial file selected')
        layout.addWidget(self.fin_label)

        # Upload Excel Template
        btn2 = QPushButton('Upload Excel Template')
        btn2.clicked.connect(self.upload_template)
        layout.addWidget(btn2)
        self.tpl_label = QLabel('No template selected')
        layout.addWidget(self.tpl_label)

        # API Key input
        api_label = QLabel('Enter API Key:')
        layout.addWidget(api_label)
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.Password)
        self.api_input.setPlaceholderText('sk-...')
        layout.addWidget(self.api_input)

        # Prompt input
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText('Enter prompt for ChatGPT here…')
        layout.addWidget(self.text_edit)

        # Submit
        submit_btn = QPushButton('Submit')
        submit_btn.clicked.connect(self.submit)
        layout.addWidget(submit_btn)

        self.setLayout(layout)

    def upload_financial(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            'Choose Financial File',
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
            self,
            'Choose Excel Template',
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
        # Validate inputs
        prompt = self.text_edit.toPlainText().strip()
        self.api_key = self.api_input.text().strip()
        if not self.api_key:
            QMessageBox.warning(self, 'Error', 'Please enter your API key.')
            return
        if not self.fin_file:
            QMessageBox.warning(self, 'Error', 'Please upload a financial file.')
            return
        if not self.tpl_file:
            QMessageBox.warning(self, 'Error', 'Please upload an Excel template.')
            return
        if not prompt:
            QMessageBox.warning(self, 'Error', 'Please enter a prompt.')
            return

        # Use extractor script to perform extraction
        try:
            result = subprocess.run(
                ['python3', 'extract.py', self.fin_file],
                capture_output=True, text=True, check=True
            )
            extracted_data = json.loads(result.stdout)
        except Exception as e:
            QMessageBox.critical(self, 'Extraction Error', f'Failed to extract data:\n{e}')
            return

        # Save extracted JSON for debugging or logging
        with open('extracted.json', 'w') as f:
            json.dump(extracted_data, f, indent=2)

        # Launch AI processor script in new Terminal
        cmd = (
            f'python3 GPT.py '
            f'--template {self.tpl_file} '
            f'--data extracted.json '
            f'--prompt "{prompt}" '
            f'--key {self.api_key}'
        )
        subprocess.Popen([
            'osascript', '-e',
            f'tell application "Terminal" to do script "{cmd}"'
        ])
        QMessageBox.information(self, 'Started', 'Extraction done; AI script launched.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = FinancialAnalysis()
    win.show()
    sys.exit(app.exec_())