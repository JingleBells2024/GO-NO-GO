import sys
import os
import json
import subprocess
import platform
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QMessageBox, QLabel, QLineEdit
)
from compiler import map_to_excel  # Import as a function

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Financial Analysis')
        self.resize(600, 400)

        self.fin_files = []
        self.tpl_file = None
        self.api_key = None
        self.extracted_data_list = []

        layout = QVBoxLayout()

        btn1 = QPushButton('Upload Financial Files')
        btn1.clicked.connect(self.upload_financial)
        layout.addWidget(btn1)
        self.fin_label = QLabel('No financial files selected')
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

        extract_btn = QPushButton('Extract Data')
        extract_btn.clicked.connect(self.extract_data)
        layout.addWidget(extract_btn)

        self.setLayout(layout)

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
        txt_path = os.path.join(os.getcwd(), 'gpt4o_extracted.txt')
        json_path = os.path.join(os.getcwd(), 'gpt4o_extracted.json')

        # Corrected: no --files flag
        args = [
            sys.executable, extractor,
            '--key', self.api_key,
            *self.fin_files
        ]
        try:
            subprocess.run(args, check=True)
        except Exception as e:
            QMessageBox.critical(self, 'Extraction Error', str(e))
            return

        # Check for the output files
        if not os.path.isfile(txt_path) or not os.path.isfile(json_path):
            QMessageBox.critical(self, 'Extraction Error', 'Extraction did not produce output files.')
            return

        reply = QMessageBox.question(
            self,
            'Extraction Complete',
            "Data extracted.\n\nWould you like to download the summary as a text file, or submit for Excel processing?",
            QMessageBox.Save | QMessageBox.Apply,
            QMessageBox.Apply
        )
        if reply == QMessageBox.Save:
            self.open_file(txt_path)
        else:
            if not self.tpl_file:
                QMessageBox.warning(self, 'Error', 'No Excel template selected.')
                return
            try:
                with open(json_path) as f:
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