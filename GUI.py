import sys
import os
import json
import subprocess
import shlex
import re
import tempfile
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
        self.tpl_file = None
        self.api_key = None

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
        self.text_edit.setPlaceholderText('Extract all financial fields for 2024...')
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

        m = re.search(r'20\d{2}', prompt)
        if not m:
            QMessageBox.warning(self, 'Error', 'Cannot find year in prompt.')
            return
        year = m.group(0)

        gpt_script = os.path.join(os.path.dirname(__file__), 'GPT.py')

        # Write a temporary shell script with the command
        cmd = [
            shlex.quote(sys.executable),
            shlex.quote(gpt_script),
            '--template', shlex.quote(self.tpl_file),
            '--data', shlex.quote(raw_file),
            '--prompt', shlex.quote(prompt),
            '--year', str(year),
            '--key', shlex.quote(self.api_key)
        ]
        full_cmd = " ".join(cmd)

        # Write to a temp file so AppleScript has ZERO quoting issues
        with tempfile.NamedTemporaryFile('w', delete=False, suffix='.sh') as tf:
            tf.write("#!/bin/bash\n")
            tf.write(full_cmd + "\n")
            script_path = tf.name

        os.chmod(script_path, 0o755)

        applescript = f'''
        tell application "Terminal"
            activate
            do script "{script_path}"
        end tell
        '''
        subprocess.Popen(['osascript', '-e', applescript])

        QMessageBox.information(self, 'Started', 'Extraction and GPT scripts launched.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = FinancialAnalysis()
    win.show()
    sys.exit(app.exec_())