from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

import streamlit as st

from excel_translator.processor import process_excel_file

LANGUAGES = {
    "English": "en",
    "French": "fr",
    "German": "de",
    "Spanish": "es",
    "Arabic": "ar",
    "Chinese (Simplified)": "zh-Hans",
    "Japanese": "ja",
}


def _extract_excel_files(uploaded_files: list) -> list[tuple[str, bytes]]:
    extracted: list[tuple[str, bytes]] = []
    for up in uploaded_files:
        name = up.name
        data = up.read()
        if name.lower().endswith(".xlsx"):
            extracted.append((name, data))
        elif name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                for member in zf.infolist():
                    if member.filename.lower().endswith(".xlsx"):
                        extracted.append((Path(member.filename).name, zf.read(member.filename)))
    return extracted


st.set_page_config(page_title="Excel Translator", layout="wide")
st.title("Enterprise Excel Translation Automation")

uploaded_files = st.file_uploader(
    "Upload one or more .xlsx files or a .zip containing .xlsx files",
    type=["xlsx", "zip"],
    accept_multiple_files=True,
)
source_lang_label = st.selectbox("Source language", list(LANGUAGES.keys()), index=0)
target_lang_label = st.selectbox("Target language", list(LANGUAGES.keys()), index=1)
engine = st.radio("Translation engine", ["azure", "local"], help="Azure auto-falls back to local on failure")
st.caption("Translate cells, sheet names, chart/drawing text (titles, labels, text boxes, shapes), comments, and notes while preserving workbook formatting.")

if st.button("Translate", type="primary"):
    files = _extract_excel_files(uploaded_files or [])
    if not files:
        st.warning("No Excel files found in upload.")
        st.stop()

    source_lang = LANGUAGES[source_lang_label]
    target_lang = LANGUAGES[target_lang_label]

    all_outputs: list[tuple[str, bytes]] = []
    all_logs = []

    progress = st.progress(0.0)
    status = st.empty()

    for idx, (name, payload) in enumerate(files, start=1):
        status.info(f"Processing {idx}/{len(files)}: {name}")
        result = process_excel_file(
            file_name=name,
            file_bytes=payload,
            source_lang=source_lang,
            target_lang=target_lang,
            selected_engine=engine,
        )
        all_outputs.append((result.output_filename, result.output_bytes))
        all_logs.extend([entry.__dict__ for entry in result.logs])
        progress.progress(idx / len(files))

    status.success("Translation completed.")

    st.subheader("Logs")
    st.dataframe(all_logs, use_container_width=True)

    st.subheader("Downloads")
    for name, payload in all_outputs:
        st.download_button(
            label=f"Download {name}",
            data=payload,
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if len(all_outputs) > 1:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, payload in all_outputs:
                zout.writestr(name, payload)
            zout.writestr("translation_logs.json", json.dumps(all_logs, ensure_ascii=False, indent=2))
        st.download_button(
            label="Download all translated files (ZIP)",
            data=zip_buf.getvalue(),
            file_name=f"translated_{target_lang}.zip",
            mime="application/zip",
        )
