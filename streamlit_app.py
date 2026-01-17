import os
import tempfile
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

from app import (
    auto_parse,
    extract_lines_pdfplumber,
    parse_lines_method,
    parse_with_tables,
    parse_vision_method,
    txns_to_dataframe,
    write_excel,
)


def convert_pdf_to_excel(
    pdf_bytes: bytes,
    method: str,
    use_ocr: bool,
    use_vision: bool,
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
            txns, diag, raw_lines = auto_parse(pdf_path, ocr=use_ocr, vision=use_vision)
        elif method in ("camelot", "tabula"):
            txns, diag, raw_lines = parse_with_tables(pdf_path, method)
            if len(txns) < 3:
                txns, diag, raw_lines = parse_lines_method(pdf_path)
        elif method == "lines":
            txns, diag, raw_lines = parse_lines_method(pdf_path)
        elif method == "vision":
            txns, diag, raw_lines = parse_vision_method(pdf_path)
        else:
            txns, diag, raw_lines = parse_lines_method(pdf_path)

        df_txn = txns_to_dataframe(txns)
        raw_lines = raw_lines if (include_raw and raw_lines) else None

        write_excel(out_path, meta, diag, df_txn, raw_lines=raw_lines)

        with open(out_path, "rb") as f:
            excel_bytes = f.read()

    return excel_bytes, diag, len(df_txn), df_txn


def human_filesize(num_bytes: int) -> str:
    """Return a human-readable file size string."""
    step_unit = 1024.0
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < step_unit:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= step_unit
    return f"{num_bytes:.1f} TB"


def main() -> None:
    st.set_page_config(page_title="PDF Statement to Excel", page_icon="📄", layout="wide")
    st.title("PDF Statement to Excel")
    st.write(
        "A faster, friendlier dashboard to turn bank statement PDFs into clean Excel files."
    )

    st.markdown(
        """
        - 📥 Upload a statement, pick your parsing strategy, and launch the conversion.
        - 👀 Instantly review a large-screen preview covering 95% of the workspace.
        - 📤 Download the Excel output once you're satisfied with the preview.
        """
    )

    if "last_result" not in st.session_state:
        st.session_state["last_result"] = None

    st.markdown("---")
    input_col, info_col = st.columns([1.1, 0.9])

    with input_col:
        st.subheader("1) Upload and configure")
        uploaded = st.file_uploader("Upload PDF", type=["pdf"], help="Supported: text PDFs and scanned (with OCR)")
        if uploaded is not None:
            st.info(
                f"File: **{uploaded.name}** · Size: {human_filesize(uploaded.size)}",
                icon="📂",
            )

        method = st.selectbox(
            "Parsing method",
            ["auto", "camelot", "tabula", "lines", "vision"],
            index=0,
            help="Choose how to extract tables and lines from the PDF.",
        )
        use_ocr = st.checkbox(
            "Enable OCR fallback (Tesseract, slower)",
            value=False,
            help="Use OCR when text is not directly extractable.",
        )
        use_vision = st.checkbox(
            "Enable Google Vision OCR fallback",
            value=False,
            help="Use Google Vision OCR when text is not directly extractable.",
        )
        include_raw = st.checkbox(
            "Include Raw sheet (debug)", value=False, help="Add raw parsed lines to the Excel output."
        )

        convert_pressed = st.button("🚀 Convert now", use_container_width=True, type="primary")

    with info_col:
        st.subheader("Helpful tips")
        st.write(
            "- `auto` works best for most PDFs, falling back to line parsing when needed.\n"
            "- Enable OCR if the statement is scanned or if other methods miss data.\n"
            "- Vision OCR can be more accurate for noisy scans if credentials are available.\n"
            "- The preview pops over the page so you can validate results before downloading."
        )
        st.info(
            "Upload a PDF and hit convert to see a nearly full-screen preview of detected transactions.",
            icon="✨",
        )

    if convert_pressed:
        if uploaded is None:
            st.warning("Please upload a PDF before converting.")
        else:
            with st.spinner("Processing your statement..."):
                try:
                    excel_bytes, diag, rows, df_txn = convert_pdf_to_excel(
                        uploaded.read(),
                        method=method,
                        use_ocr=use_ocr,
                        use_vision=use_vision,
                        include_raw=include_raw,
                    )
                except Exception as exc:  # pragma: no cover - interactive error path
                    st.error(
                        "Conversion failed. Please verify the PDF and method, then try again.",
                        icon="🚨",
                    )
                    st.exception(exc)
                else:
                    st.session_state["last_result"] = {
                        "excel": excel_bytes,
                        "diag": diag,
                        "rows": rows,
                        "df": df_txn,
                    }
    last_result = st.session_state.get("last_result")

    if last_result:
        df_txn = last_result["df"]
        diag = last_result["diag"]
        rows = last_result["rows"]
        excel_bytes = last_result["excel"]

        st.markdown("---")
        st.subheader("2) Review output")
        metric_col1, metric_col2, metric_col3 = st.columns(3)
        metric_col1.metric("Rows detected", rows)
        metric_col2.metric("Method used", diag.get("method", method))
        metric_col3.metric("Pages parsed", diag.get("pages", "n/a"))

        if not df_txn.empty:
            st.caption("Conversion preview (showing up to 200 rows)")
            st.dataframe(df_txn.head(200), use_container_width=True, height=500)
            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name="statement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.warning("No transactions were detected in the uploaded PDF.")


if __name__ == "__main__":
    main()
