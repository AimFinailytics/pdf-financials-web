# PDF Financial Statement Converter

Local web app for converting uploaded annual or quarterly report PDFs into Excel workbooks.

## Run

```powershell
python -m pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5050`.

## Output

- One Excel workbook per company.
- A master workbook when two or more PDFs are uploaded.
- Each workbook has three sheets: `Income Statement`, `Balance Sheet`, and `Cash Flow Statement`.
