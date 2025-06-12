# GO-NO-GO Financial Filler (Desktop App)

A lightweight Python desktop application that allows users to:

- Upload a financial document (PDF or Excel)
- Upload a pre-formatted Excel template
- Input a custom prompt to guide ChatGPT
- Automatically extract, process, and fill in key financial fields
- Save and optionally open the filled template in Excel or Numbers

---

## ✅ Features

- Upload and parse financial files
- Guide extraction with a prompt using ChatGPT API
- Preserve formatting in pre-styled Excel templates
- Auto-fill and save output as `filled_output.xlsx`
- Option to open the result in your default spreadsheet app (macOS)

---

## 💻 Requirements

- Python 3.7+
- macOS (for now)
- `openpyxl`
- `tkinter` (bundled with Python)
- [OpenAI API Key](https://platform.openai.com/account/api-keys)

Install dependencies:

```bash
pip3 install openpyxl
