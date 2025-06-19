import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QMessageBox, QLabel, QLineEdit, QSizePolicy
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from compiler import map_to_excel

LOGO_PATH = os.path.join(os.path.dirname(__file__), "Logo.png")
API_KEY_FILE = os.path.join(os.path.expanduser("~"), ".dealsheet_api_key")

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Deal Sheet – Financial Analysis')
        self.resize(470, 750)
        self.setStyleSheet("background: #28292e;")
        self.fin_files = []
        self.tpl_file = None
        self.api_key = None
        self.extracted_data_list = []
        self.json_data = None

        outer = QVBoxLayout()
        outer.setAlignment(Qt.AlignTop)
        outer.setSpacing(12)

        logo_label = QLabel()
        pixmap = QPixmap(LOGO_PATH)
        pixmap = pixmap.scaled(280, 280, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignHCenter)
        outer.addWidget(logo_label, alignment=Qt.AlignHCenter)

        card = QVBoxLayout()
        card.setAlignment(Qt.AlignTop)
        card.setSpacing(10)
        card.setContentsMargins(0, 0, 0, 0)

        def add_card_widget(w):
            w.setFixedWidth(330)
            w.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            card.addWidget(w, alignment=Qt.AlignHCenter)

        # Financial files
        btn1 = QPushButton('Upload Financial Files')
        btn1.clicked.connect(self.upload_financial)
        add_card_widget(btn1)
        self.fin_label = QLabel('No financial files selected')
        self.fin_label.setStyleSheet('color: #eee')
        add_card_widget(self.fin_label)

        btn2 = QPushButton('Upload Excel Template')
        btn2.clicked.connect(self.upload_template)
        add_card_widget(btn2)
        self.tpl_label = QLabel('No template selected')
        self.tpl_label.setStyleSheet('color: #eee')
        add_card_widget(self.tpl_label)

        api_label = QLabel('Enter API Key:')
        api_label.setStyleSheet('color: #eee')
        add_card_widget(api_label)
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.Password)
        self.api_input.setPlaceholderText('sk-...')
        add_card_widget(self.api_input)

        # Load API key if saved
        self.load_api_key()

        # Save API key button
        save_key_btn = QPushButton('Save API Key')
        save_key_btn.clicked.connect(self.save_api_key)
        add_card_widget(save_key_btn)

        extract_btn = QPushButton('Extract Data')
        extract_btn.clicked.connect(self.extract_data)
        add_card_widget(extract_btn)

        chatgpt_btn = QPushButton('Open ChatGPT Web UI')
        chatgpt_btn.clicked.connect(lambda: self.open_url("https://chat.openai.com"))
        add_card_widget(chatgpt_btn)

        copy_btn = QPushButton('Copy Extracted JSON to Clipboard')
        copy_btn.clicked.connect(self.copy_json_to_clipboard)
        add_card_widget(copy_btn)

        upload_btn = QPushButton('Upload Extracted JSON File (from ChatGPT)')
        upload_btn.clicked.connect(self.upload_json_from_chatgpt)
        add_card_widget(upload_btn)

        outer.addLayout(card)
        self.setLayout(outer)

    def load_api_key(self):
        """Loads API key from file if available."""
        if os.path.exists(API_KEY_FILE):
            try:
                with open(API_KEY_FILE, 'r') as f:
                    key = f.read().strip()
                self.api_input.setText(key)
            except Exception:
                pass

    def save_api_key(self):
        """Saves API key to local file."""
        key = self.api_input.text().strip()
        if not key:
            QMessageBox.warning(self, 'Error', 'API key is empty.')
            return
        try:
            with open(API_KEY_FILE, 'w') as f:
                f.write(key)
            os.chmod(API_KEY_FILE, 0o600)  # Only readable/writable by user
            QMessageBox.information(self, 'Success', 'API key saved securely.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to save API key: {e}')

    def upload_financial(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 'Choose Financial Files',
            filter='PDF Files (*.pdf);;All Files (*)'
        )
        if files:
            self.fin_files = files
            self.fin_label.setText(', '.join([os.path.basename(f) for f in files]))
            self.fin_label.setStyleSheet('color: #7ee57a')
        else:
            self.fin_files = []
            self.fin_label.setText('No financial files selected')
            self.fin_label.setStyleSheet('color: #ff6961')

    def upload_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Choose Excel Template',
            filter='Excel Files (*.xlsx *.xlsm)'
        )
        if path:
            self.tpl_file = path
            self.tpl_label.setText(os.path.basename(path))
            self.tpl_label.setStyleSheet('color: #7ee57a')
        else:
            self.tpl_file = None
            self.tpl_label.setText('No template selected')
            self.tpl_label.setStyleSheet('color: #ff6961')

    def extract_data(self):
        self.api_key = self.api_input.text().strip()
        if not self.api_key:
            QMessageBox.warning(self, 'Error', 'Enter API key.')
            return
        if not self.fin_files:
            QMessageBox.warning(self, 'Error', 'Upload at least one financial file.')
            return

        extractor = os.path.join(os.path.dirname(__file__), 'extract.py')
        json_path = os.path.join(os.getcwd(), 'gpt4o_extracted.json')

        args = [
            sys.executable, extractor,
            '--key', self.api_key,
            '--prompt', '',  # No user prompt needed now
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

        with open(json_path) as f:
            self.json_data = json.load(f)
        QMessageBox.information(self, 'Extraction Complete', 'Data extracted and ready.')

    def copy_json_to_clipboard(self):
        if self.json_data is None:
            QMessageBox.warning(self, 'Copy Error', 'No extracted JSON loaded.')
            return
        cb = QApplication.clipboard()
        cb.setText(json.dumps(self.json_data, indent=2))
        QMessageBox.information(self, 'Copied', 'Extracted JSON copied to clipboard!')

    def upload_json_from_chatgpt(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Select Extracted JSON File',
            filter='JSON Files (*.json);;All Files (*)'
        )
        if not path:
            return
        try:
            with open(path) as f:
                data = json.load(f)
            self.json_data = data
        except Exception as e:
            QMessageBox.critical(self, 'Load Error', f"Failed to load JSON: {e}")
            return

        if not self.tpl_file:
            QMessageBox.warning(self, 'Error', 'No Excel template selected.')
            return
        try:
            map_to_excel(self.json_data, self.tpl_file, None)
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
        if platform.system() == 'Darwin':
            subprocess.run(['open', filepath])
        elif platform.system() == 'Windows':
            os.startfile(filepath)
        else:
            subprocess.run(['xdg-open', filepath])

    def open_url(self, url):
        import webbrowser
        webbrowser.open(url)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = FinancialAnalysis()
    win.show()
    sys.exit(app.exec_())