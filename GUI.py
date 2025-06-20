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
from compiler import map_to_excel

API_KEY_FILE = "api_key.txt"
BUTTON_WIDTH = 340

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)

LOGO_PATH = resource_path("Logo.png")

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

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Deal Sheet – Financial Analysis')
        self.resize(BUTTON_WIDTH + 80, 700)
        self.center()
        self.fin_files = []
        self.tpl_file = None
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
            logo_label.setPixmap(pixmap.scaled(250, 250, Qt.KeepAspectRatio, Qt.SmoothTransformation))
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
        self.api_status = QLabel(self.api_key_status())
        self.api_status.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.api_status)
        # Submit Button
        submit_btn = QPushButton('Submit')
        submit_btn.setFixedWidth(BUTTON_WIDTH)
        submit_btn.clicked.connect(self.submit)
        main_layout.addWidget(submit_btn)
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

    def submit(self):
        if not self.api_key or not (self.api_key.startswith("sk-") and len(self.api_key) > 10):
            QMessageBox.warning(self, 'Error', 'Please enter and save a valid API key.')
            return
        if not self.fin_files:
            QMessageBox.warning(self, 'Error', 'Upload at least one financial file.')
            return
        if not self.tpl_file:
            QMessageBox.warning(self, 'Error', 'Upload an Excel template.')
            return
        extractor = os.path.join(os.path.dirname(__file__), 'extract.py')
        json_path = os.path.join(os.getcwd(), 'gpt4o_extracted.json')
        args = [
            sys.executable, extractor,
            '--key', self.api_key,
            '--prompt', '',  # Prompt left blank for hardcoded
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
        # Now fill the Excel template
        try:
            with open(json_path) as f:
                data = json.load(f)
            map_to_excel(data, self.tpl_file, None)
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