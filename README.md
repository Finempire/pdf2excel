# pdf2excel

Convert bank statement PDFs to Excel using Streamlit or the CLI tools in `app.py`.

## Running locally
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Launch the Streamlit app:
   ```bash
   streamlit run streamlit_app.py
   ```

The dependency list ensures packages like `pdfplumber` are available so the app can import and process PDFs.

## Google Vision OCR
You can enable Google Vision OCR (CLI `--vision` or Streamlit checkbox) for scanned statements. Set
`GOOGLE_APPLICATION_CREDENTIALS` to your service account JSON, or keep the provided JSON file in the repo
root so the app can pick it up automatically.
