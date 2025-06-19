import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QMessageBox, QLabel, QLineEdit, QHBoxLayout
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from compiler import map_to_excel  # Import as a function

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Deal Sheet – Financial Analysis')
        self.resize(700, 560)

        self.fin_files = []
        self.tpl_file = None
        self.api_key = None
        self.extracted_data_list = []
        self.extracted_json_path = os.path.join(os.getcwd(), 'gpt4o_extracted.json')

        main_layout = QVBoxLayout()

        # Logo at the top
        logo_path = os.path.join(os.path.dirname(__file__), "Logo.png")
        if os.path.isfile(logo_path):
            logo_label = QLabel()
            pixmap = QPixmap(logo_path)
            pixmap = pixmap.scaledToWidth(220, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            logo_label.setAlignment(Qt.AlignCenter)
            main_layout.addWidget(logo_label)
        else:
            main_layout.addWidget(QLabel("(Logo not found)"))

        # Upload Financial Files
        btn1 = QPushButton('Upload Financial Files')
        btn1.clicked.connect(self.upload_financial)
        main_layout.addWidget(btn1)
        self.fin_label = QLabel('No financial files selected')
        main_layout.addWidget(self.fin_label)

        # Upload Excel Template
        btn2 = QPushButton('Upload Excel Template')
        btn2.clicked.connect(self.upload_template)
        main_layout.addWidget(btn2)
        self.tpl_label = QLabel('No template selected')
        main_layout.addWidget(self.tpl_label)

        # API Key input
        api_label = QLabel('Enter API Key:')
        main_layout.addWidget(api_label)
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.Password)
        self.api_input.setPlaceholderText('sk-...')
        main_layout.addWidget(self.api_input)

        # Extract Data button
        extract_btn = QPushButton('Extract Data')
        extract_btn.clicked.connect(self.extract_data)
        main_layout.addWidget(extract_btn)

        # Open ChatGPT Web UI button
        open_web_btn = QPushButton('Open ChatGPT Web UI')
        open_web_btn.clicked.connect(self.open_web_ui)
        main_layout.addWidget(open_web_btn)

        # Copy extracted JSON to clipboard
        copy_json_btn = QPushButton('Copy Extracted JSON to Clipboard')
        copy_json_btn.clicked.connect(self.copy_json_to_clipboard)
        main_layout.addWidget(copy_json_btn)

        # Upload extracted JSON file from web UI
        upload_json_btn = QPushButton('Upload Extracted JSON File (from ChatGPT)')
        upload_json_btn.clicked.connect(self.upload_json_file)
        main_layout.addWidget(upload_json_btn)

        self.setLayout(main_layout)

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
        self.api_key = self.api_input.text().strip()
        if not self.api_key:
            QMessageBox.warning(self, 'Error', 'Enter API key.')
            return
        if not self.fin_files:
            QMessageBox.warning(self, 'Error', 'Upload at least one financial file.')
            return

        extractor = os.path.join(os.path.dirname(__file__), 'extract.py')
        json_path = self.extracted_json_path

        args = [
            sys.executable, extractor,
            '--key', self.api_key,
            '--prompt', '',  # No user prompt, just pass empty string
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

        QMessageBox.information(self, 'Extraction Complete', "Data extraction complete.\nYou can now copy the JSON, open it in ChatGPT, or upload the final JSON for Excel processing.")

    def open_web_ui(self):
        import webbrowser
        webbrowser.open('https://chat.openai.com/')  # or use your preferred web UI

    def copy_json_to_clipboard(self):
        if not os.path.isfile(self.extracted_json_path):
            QMessageBox.warning(self, 'Error', 'No extracted JSON file found.')
            return
        with open(self.extracted_json_path) as f:
            data = f.read()
        cb = QApplication.clipboard()
        cb.setText(data)
        QMessageBox.information(self, 'Copied', 'Extracted JSON copied to clipboard!')

    def upload_json_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Choose Extracted JSON File',
            filter='JSON Files (*.json);;All Files (*)'
        )
        if not path:
            return
        if not self.tpl_file:
            QMessageBox.warning(self, 'Error', 'No Excel template selected.')
            return
        try:
            with open(path) as f:
                data = json.load(f)
            # Accept both dict or list
            extracted_data_list = []
            if isinstance(data, dict):
                extracted_data_list.append(data)
            elif isinstance(data, list):
                extracted_data_list.extend(data)
            else:
                raise Exception("Extracted JSON is not dict or list.")
        except Exception as e:
            QMessageBox.critical(self, 'Data Error', f'Problem loading extracted data: {e}')
            return

        try:
            map_to_excel(extracted_data_list, self.tpl_file, None)
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