import os
import tempfile
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

from app import parse_ocr_method, parse_vision_method, txns_to_dataframe, write_excel


def convert_pdf_to_excel(
    pdf_bytes: bytes,
) -> Tuple[bytes, dict, int, pd.DataFrame]:
    """
    Run the existing conversion pipeline against uploaded bytes and return Excel bytes.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, "input.pdf")
        out_path = os.path.join(tmpdir, "output.xlsx")

        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        # Metadata is extracted inside the CLI tool, but here we only care about
        # transactions + diag.
        meta = {}

        try:
            txns, diag, raw_lines = parse_vision_method(pdf_path)
        except RuntimeError as exc:
            vision_error = str(exc)
            if "Vision" in vision_error or "Unauthenticated" in vision_error or "credentials" in vision_error.lower():
                txns, diag, raw_lines = parse_ocr_method(pdf_path)
                diag["warning"] = (
                    "Google Vision OCR failed; falling back to local OCR. "
                    f"Reason: {vision_error}"
                )
            else:
                raise

        df_txn = txns_to_dataframe(txns)
        raw_lines = None

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
        - 📥 Upload a statement and launch the conversion.
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

        convert_pressed = st.button("🚀 Convert now", use_container_width=True, type="primary")

    with info_col:
        st.subheader("Helpful tips")
        st.write(
            "- Upload a bank statement PDF and the backend will parse it with Google Vision.\n"
            "- Make sure the Vision credentials are configured on the server.\n"
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
        metric_col2.metric("Method used", diag.get("method", "vision"))
        metric_col3.metric("Pages parsed", diag.get("pages", "n/a"))
        warning = diag.get("warning")
        if warning:
            st.warning(warning, icon="⚠️")

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
