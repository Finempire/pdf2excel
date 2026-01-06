import os
import tempfile
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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


def human_filesize(num_bytes: int) -> str:
    """Return a human-readable file size string."""
    step_unit = 1024.0
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < step_unit:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= step_unit
    return f"{num_bytes:.1f} TB"


def render_preview_overlay(df: pd.DataFrame, diag: dict, rows: int) -> None:
    """Render a large, nearly full-screen preview using a custom HTML overlay."""

    method_used = diag.get("method") or "auto"
    preview_table = df.head(200).to_html(index=False, classes="preview-table")
    html = f"""
    <style>
        .preview-scrim {{
            position: fixed;
            inset: 0;
            background: rgba(12, 18, 28, 0.75);
            backdrop-filter: blur(2px);
            z-index: 998;
        }}
        .preview-panel {{
            position: fixed;
            inset: 2vh 2vw;
            width: 96vw;
            height: 96vh;
            background: #0b1224;
            border-radius: 16px;
            border: 1px solid #1e293b;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.55);
            z-index: 999;
            display: flex;
            flex-direction: column;
            padding: 18px 22px;
            color: #e2e8f0;
            overflow: hidden;
        }}
        .preview-top {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 10px;
        }}
        .preview-title {{
            margin: 0;
            font-size: 1.25rem;
            font-weight: 700;
            color: #c7d2fe;
        }}
        .preview-sub {{
            margin: 0;
            color: #94a3b8;
            font-size: 0.95rem;
        }}
        .preview-close {{
            background: #1e293b;
            color: #e2e8f0;
            border: 1px solid #334155;
            padding: 8px 12px;
            border-radius: 10px;
            cursor: pointer;
            font-weight: 600;
        }}
        .preview-close:hover {{
            background: #334155;
        }}
        .preview-meta {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 12px;
            margin: 12px 0 8px 0;
        }}
        .preview-meta .item {{
            background: #0f172a;
            border: 1px solid #1e293b;
            border-radius: 12px;
            padding: 10px 12px;
        }}
        .preview-meta .label {{
            color: #94a3b8;
            font-size: 0.8rem;
            margin-bottom: 4px;
        }}
        .preview-meta .value {{
            color: #e2e8f0;
            font-weight: 600;
            font-size: 1rem;
        }}
        .preview-body {{
            flex: 1;
            overflow: auto;
            background: #0f172a;
            border: 1px solid #1e293b;
            border-radius: 12px;
            padding: 12px;
        }}
        .preview-table {{
            width: 100%;
            border-collapse: collapse;
            color: #e2e8f0;
            font-size: 0.9rem;
        }}
        .preview-table th,
        .preview-table td {{
            border: 1px solid #1e293b;
            padding: 6px 8px;
        }}
        .preview-table th {{
            background: #111827;
            color: #c7d2fe;
            position: sticky;
            top: 0;
            z-index: 1;
        }}
        .preview-table tbody tr:nth-child(odd) {{
            background: #0b1224;
        }}
    </style>
    <div class="preview-scrim"></div>
    <div class="preview-panel">
        <div class="preview-top">
            <div>
                <p class="preview-title">Conversion preview</p>
                <p class="preview-sub">Showing up to 200 recent rows with parsing details.</p>
            </div>
            <button class="preview-close" onclick="document.querySelector('.preview-panel').style.display='none';document.querySelector('.preview-scrim').style.display='none';">Close</button>
        </div>
        <div class="preview-meta">
            <div class="item"><div class="label">Rows detected</div><div class="value">{rows}</div></div>
            <div class="item"><div class="label">Method</div><div class="value">{method_used}</div></div>
            <div class="item"><div class="label">Pages parsed</div><div class="value">{diag.get('pages', 'n/a')}</div></div>
            <div class="item"><div class="label">Parser hints</div><div class="value">{diag.get('notes', 'No additional notes')}</div></div>
        </div>
        <div class="preview-body">{preview_table}</div>
    </div>
    """

    components.html(html, height=900, scrolling=False)


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
    if "show_preview" not in st.session_state:
        st.session_state["show_preview"] = False

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
            ["auto", "camelot", "tabula", "lines"],
            index=0,
            help="Choose how to extract tables and lines from the PDF.",
        )
        use_ocr = st.checkbox(
            "Enable OCR fallback (slower)", value=False, help="Use OCR when text is not directly extractable."
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
                    st.session_state["show_preview"] = True

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
            st.dataframe(df_txn.head(25), use_container_width=True)
            preview_col1, preview_col2 = st.columns([1.5, 1])
            with preview_col1:
                if st.button("Open large preview", key="open_preview"):
                    st.session_state["show_preview"] = True
            with preview_col2:
                st.download_button(
                    label="Download Excel",
                    data=excel_bytes,
                    file_name="statement.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.warning("No transactions were detected in the uploaded PDF.")

    if st.session_state.get("show_preview") and last_result and not last_result["df"].empty:
        render_preview_overlay(last_result["df"], last_result["diag"], last_result["rows"])


if __name__ == "__main__":
    main()
