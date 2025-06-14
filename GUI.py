import sys
import os
import json
import subprocess
import re
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QTextEdit, QFileDialog, QMessageBox, QLabel, QLineEdit
)

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Financial Analysis')
        self.resize(600, 400)

        self.fin_file = None
        self.tpl_file = None  # kept for future, but not used here
        self.api_key = None

        layout = QVBoxLayout()

        # Financial File selector
        btn1 = QPushButton('Upload Financial PDF/Excel')
        btn1.clicked.connect(self.upload_financial)
        layout.addWidget(btn1)
        self.fin_label = QLabel('No financial file selected')
        layout.addWidget(self.fin_label)

        # Template selector (optional here)
        btn2 = QPushButton('Upload Excel Template (optional)')
        btn2.clicked.connect(self.upload_template)
        layout.addWidget(btn2)
        self.tpl_label = QLabel('No template selected')
        layout.addWidget(self.tpl_label)

        # API key input
        api_label = QLabel('Enter API Key:')
        layout.addWidget(api_label)
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.Password)
        self.api_input.setPlaceholderText('sk-...')
        layout.addWidget(self.api_input)

        # Prompt input
        prompt_label = QLabel('GPT Prompt:')
        layout.addWidget(prompt_label)
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText('Extract all financial fields for 2024...')
        layout.addWidget(self.text_edit)

        # Submit
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

        # Validate
        if not self.api_key:
            QMessageBox.warning(self, 'Error', 'Enter API key.')
            return
        if not self.fin_file:
            QMessageBox.warning(self, 'Error', 'Upload a financial file.')
            return
        if not prompt:
            QMessageBox.warning(self, 'Error', 'Enter a prompt.')
            return

        # Run extractor script
        extractor = os.path.join(os.path.dirname(__file__), 'extract.py')
        try:
            res = subprocess.run(
                [sys.executable, extractor, self.fin_file],
                capture_output=True, text=True, check=True
            )
            raw_file = os.path.abspath('extracted.json')
            with open(raw_file, 'w') as f:
                f.write(res.stdout)
        except Exception as e:
            QMessageBox.critical(self, 'Extraction Error', str(e))
            return

        # Run GPT.py without --template or --year
        gpt_script = os.path.join(os.path.dirname(__file__), 'GPT.py')
        try:
            subprocess.run([
                sys.executable, gpt_script,
                '--data', raw_file,
                '--prompt', prompt,
                '--key', self.api_key
            ], check=True)
        except Exception as e:
            QMessageBox.critical(self, 'GPT Error', str(e))
            return

        QMessageBox.information(self, 'Done', 'GPT processing complete.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = FinancialAnalysis()
    win.show()
    sys.exit(app.exec_())