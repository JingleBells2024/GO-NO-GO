import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QMessageBox, QLabel, QDialog, QLineEdit, QDialogButtonBox, QDesktopWidget
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from compiler import map_to_excel  # Import as a function

API_KEY_FILE = "api_key.txt"
LOGO_PATH = resource_path("Logo.png")
BUTTON_WIDTH = 340  # Unified width for all elements

class ApiKeyDialog(QDialog):
    def __init__(self, parent=None, current_key=""):
        super().__init__(parent)
        self.setWindowTitle("Enter API Key")
        self.resize(380, 120)
        layout = QVBoxLayout()
        self.label = QLabel("Paste your OpenAI API key below:")
        self.input = QLineEdit()
        self.input.setEchoMode(QLineEdit.Password)
        self.input.setText(current_key)
        layout.addWidget(self.label)
        layout.addWidget(self.input)
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_key(self):
        return self.input.text().strip()

    def resource_path(relative_path):
        # When packaged, sys._MEIPASS is the temp folder where resources are extracted
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.dirname(__file__), relative_path)

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Deal Sheet – Financial Analysis')
        self.resize(BUTTON_WIDTH + 80, 740)
        self.center()

        self.fin_files = []
        self.tpl_file = None
        self.api_key = None
        self.extracted_data_list = []
        self.last_extracted_json = None

        # Load API key if exists
        self.api_key = ""
        if os.path.exists(API_KEY_FILE):
            with open(API_KEY_FILE, 'r') as f:
                self.api_key = f.read().strip()

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)

        # Logo
        if os.path.exists(LOGO_PATH):
            logo_label = QLabel()
            pixmap = QPixmap(LOGO_PATH)
            logo_label.setPixmap(pixmap.scaled(300, 300, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            logo_label.setAlignment(Qt.AlignCenter)
            main_layout.addWidget(logo_label)

        # Upload financial files
        btn1 = QPushButton('Upload Financial Files')
        btn1.setFixedWidth(BUTTON_WIDTH)
        btn1.clicked.connect(self.upload_financial)
        main_layout.addWidget(btn1)
        self.fin_label = QLabel('No financial files selected')
        self.fin_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.fin_label)

        # Upload Excel template
        btn2 = QPushButton('Upload Excel Template')
        btn2.setFixedWidth(BUTTON_WIDTH)
        btn2.clicked.connect(self.upload_template)
        main_layout.addWidget(btn2)
        self.tpl_label = QLabel('No template selected')
        self.tpl_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.tpl_label)

        # API Key Button
        self.api_btn = QPushButton('Enter API Key')
        self.api_btn.setFixedWidth(BUTTON_WIDTH)
        self.api_btn.clicked.connect(self.open_api_dialog)
        main_layout.addWidget(self.api_btn)

        # Show status for API key
        self.api_status = QLabel(self.api_key_status())
        self.api_status.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.api_status)

        # Buttons (all fixed width)
        extract_btn = QPushButton('Extract Data')
        extract_btn.setFixedWidth(BUTTON_WIDTH)
        extract_btn.clicked.connect(self.extract_data)
        main_layout.addWidget(extract_btn)

        open_web_btn = QPushButton('Open ChatGPT Web UI')
        open_web_btn.setFixedWidth(BUTTON_WIDTH)
        open_web_btn.clicked.connect(self.open_web)
        main_layout.addWidget(open_web_btn)

        copy_btn = QPushButton('Copy Extracted JSON to Clipboard')
        copy_btn.setFixedWidth(BUTTON_WIDTH)
        copy_btn.clicked.connect(self.copy_json)
        main_layout.addWidget(copy_btn)

        upload_json_btn = QPushButton('Upload Extracted JSON File (from ChatGPT)')
        upload_json_btn.setFixedWidth(BUTTON_WIDTH)
        upload_json_btn.clicked.connect(self.upload_json)
        main_layout.addWidget(upload_json_btn)

        self.setLayout(main_layout)

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def api_key_status(self):
        if self.api_key and self.api_key.startswith("sk-") and len(self.api_key) > 10:
            return "API Key saved ✔"
        else:
            return "No API Key saved"

    def open_api_dialog(self):
        dlg = ApiKeyDialog(self, current_key=self.api_key)
        if dlg.exec_():
            key = dlg.get_key()
            if key:
                self.api_key = key
                with open(API_KEY_FILE, 'w') as f:
                    f.write(self.api_key)
                self.api_status.setText("API Key saved ✔")
                QMessageBox.information(self, "Saved", "API key saved securely.")
            else:
                self.api_status.setText("No API Key saved")
                QMessageBox.warning(self, "Warning", "API key cannot be empty.")

    def upload_financial(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 'Choose Financial Files',
            filter='All Files (*.pdf *.xlsx *.xlsm *.jpg *.png)'
        )
        if files:
            self.fin_files = files
            self.fin_label.setText(', '.join([os.path.basename(f) for f in files]))
            self.fin_label.setStyleSheet('color: green')
        else:
            self.fin_files = []
            self.fin_label.setText('No financial files selected')
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

    def extract_data(self):
        if not self.api_key or not (self.api_key.startswith("sk-") and len(self.api_key) > 10):
            QMessageBox.warning(self, 'Error', 'Please enter and save a valid API key.')
            return
        if not self.fin_files:
            QMessageBox.warning(self, 'Error', 'Upload at least one financial file.')
            return

        extractor = os.path.join(os.path.dirname(__file__), 'extract.py')
        json_path = os.path.join(os.getcwd(), 'gpt4o_extracted.json')

        args = [
            sys.executable, extractor,
            '--key', self.api_key,
            '--prompt', '',  # blank prompt, as UI prompt is removed
            *self.fin_files
        ]
        try:
            subprocess.run(args, check=True)
        except Exception as e:
            QMessageBox.critical(self, 'Extraction Error', str(e))
            return

        if not os.path.isfile(json_path):
            QMessageBox.critical(self, 'Extraction Error', 'Extraction did not produce output file.')
            return

        self.last_extracted_json = json_path
        QMessageBox.information(self, "Extraction Complete", "Data extraction completed successfully.")

    def copy_json(self):
        json_path = self.last_extracted_json or os.path.join(os.getcwd(), 'gpt4o_extracted.json')
        if not os.path.isfile(json_path):
            QMessageBox.warning(self, "Copy Error", "No extracted JSON to copy.")
            return
        with open(json_path) as f:
            data = f.read()
        cb = QApplication.clipboard()
        cb.setText(data)
        QMessageBox.information(self, "Copied", "Extracted JSON copied to clipboard.")

    def open_web(self):
        import webbrowser
        webbrowser.open('https://chat.openai.com/')

    def upload_json(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select Extracted JSON File',
            filter='JSON Files (*.json)'
        )
        if not path:
            return
        if not self.tpl_file:
            QMessageBox.warning(self, 'Error', 'No Excel template selected.')
            return
        try:
            with open(path) as f:
                data = json.load(f)
            if isinstance(data, dict):
                self.extracted_data_list.append(data)
            elif isinstance(data, list):
                self.extracted_data_list.extend(data)
            else:
                raise Exception("Extracted JSON is not dict or list.")
        except Exception as e:
            QMessageBox.critical(self, 'Data Error', f'Problem loading extracted data: {e}')
            return

        try:
            map_to_excel(self.extracted_data_list, self.tpl_file, None)
        except Exception as e:
            QMessageBox.critical(self, 'Compile Error', str(e))
            return

        open_reply = QMessageBox.question(
            self,
            'Processing complete',
            f'Excel file updated: {self.tpl_file}\n\nOpen now?',
            QMessageBox.Yes | QMessageBox.No
        )
        if open_reply == QMessageBox.Yes:
            self.open_file(self.tpl_file)
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