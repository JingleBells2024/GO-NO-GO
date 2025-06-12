import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QTextEdit, QFileDialog, QMessageBox
)

class FinancialAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Financial Analysis')
        self.resize(400, 350)

        self.fin_file = None
        self.tpl_file = None

        layout = QVBoxLayout()

        # Upload buttons
        btn1 = QPushButton('Upload Financial PDF/Excel')
        btn1.clicked.connect(self.upload_financial)
        layout.addWidget(btn1)

        btn2 = QPushButton('Upload Excel Template')
        btn2.clicked.connect(self.upload_template)
        layout.addWidget(btn2)

        # ← Visible, multi-line text box
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText('Enter prompt for ChatGPT here…')
        layout.addWidget(self.text_edit)

        # Submit
        submit = QPushButton('Submit')
        submit.clicked.connect(self.submit)
        layout.addWidget(submit)

        self.setLayout(layout)

    def upload_financial(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Choose Financial File',
            filter='PDF Files (*.pdf);;Excel Files (*.xlsx)'
        )
        if path:
            self.fin_file = path

    def upload_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Choose Excel Template',
            filter='Excel Files (*.xlsx)'
        )
        if path:
            self.tpl_file = path

    def submit(self):
        prompt = self.text_edit.toPlainText().strip()
        if not self.fin_file or not self.tpl_file:
            QMessageBox.warning(self, 'Error', 'Please upload both files first.')
            return
        if not prompt:
            QMessageBox.warning(self, 'Error', 'Please enter a prompt.')
            return

        # For demo, just echo it:
        QMessageBox.information(self, 'Submitted Prompt', prompt)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = FinancialAnalysis()
    win.show()
    sys.exit(app.exec_())