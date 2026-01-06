import os
import tempfile
from typing import Optional, Tuple

import streamlit as st
import pandas as pd

from app import (
    auto_parse,
    extract_lines_pdfplumber,
    parse_lines_method,
    parse_with_tables,
    txns_to_dataframe,
    write_excel,
)


def convert_pdf_to_excel(
    pdf_bytes: bytes,
    method: str,
    use_ocr: bool,
    include_raw: bool,
) -> Tuple[bytes, dict, int, pd.DataFrame]:
    """
    Run the existing conversion pipeline against uploaded bytes and return Excel bytes.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, "input.pdf")
        out_path = os.path.join(tmpdir, "output.xlsx")

        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        try:
            lines_for_meta = extract_lines_pdfplumber(pdf_path)
            full_text = "\n".join(lines_for_meta)
        except Exception:
            lines_for_meta = []
            full_text = ""

        # Metadata is extracted inside auto_parse/parse_lines in the CLI tool,
        # but here we only care about transactions + diag.
        meta = {}

        if method == "auto":
            txns, diag, raw_lines = auto_parse(pdf_path, ocr=use_ocr)
        elif method in ("camelot", "tabula"):
            txns, diag, raw_lines = parse_with_tables(pdf_path, method)
            if len(txns) < 3:
                txns, diag, raw_lines = parse_lines_method(pdf_path)
        elif method == "lines":
            txns, diag, raw_lines = parse_lines_method(pdf_path)
        else:
            txns, diag, raw_lines = parse_lines_method(pdf_path)

        df_txn = txns_to_dataframe(txns)
        raw_lines = raw_lines if (include_raw and raw_lines) else None

        write_excel(out_path, meta, diag, df_txn, raw_lines=raw_lines)

        with open(out_path, "rb") as f:
            excel_bytes = f.read()

    return excel_bytes, diag, len(df_txn), df_txn


def main() -> None:
    st.set_page_config(page_title="PDF Statement to Excel", page_icon="📄")
    st.title("PDF Statement to Excel")
    st.write("Upload a bank statement PDF, choose a parsing method, preview rows, and download Excel.")

    uploaded = st.file_uploader("Upload PDF", type=["pdf"])
    method = st.selectbox("Method", ["auto", "camelot", "tabula", "lines"], index=0)
    use_ocr = st.checkbox("Enable OCR fallback (slower)", value=False)
    include_raw = st.checkbox("Include Raw sheet (debug)", value=False)

    if st.button("Convert") and uploaded is not None:
        with st.spinner("Processing..."):
            try:
                excel_bytes, diag, rows, df_txn = convert_pdf_to_excel(
                    uploaded.read(),
                    method=method,
                    use_ocr=use_ocr,
                    include_raw=include_raw,
                )
            except Exception as exc:  # pragma: no cover - interactive error path
                st.error("Conversion failed. Please verify the PDF and method, then try again.")
                st.exception(exc)
                return

        st.success(f"Done. Rows: {rows}. Method used: {diag.get('method', method)}")
        st.write("Diagnostics:", diag)

        if not df_txn.empty:
            st.subheader("Preview")
            st.dataframe(df_txn.head(50))
        else:
            st.warning("No transactions were detected in the uploaded PDF.")

        st.download_button(
            label="Download Excel",
            data=excel_bytes,
            file_name="statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
