# app.py

import streamlit as st

from logic.allele_compare import run_batch_comparison, timestamp_str

st.set_page_config(page_title="Batch Allele Comparison Tool (Excel Files)", layout="centered")

st.title("🧬 Batch Allele Comparison Tool (Excel Files)")
st.markdown("Upload one or more Excel files (`.xls` or `.xlsx`) and download a ZIP of results.")

uploaded_files = st.file_uploader(
    "Drag and drop multiple Excel files here, or browse to select",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

run_btn = st.button("Run Batch Comparison", type="primary", disabled=not uploaded_files)

progress_text = st.empty()
progress_bar = st.progress(0)

if run_btn and uploaded_files:
    try:
        # Show live progress text while processing (simple UX: text only, bar fills at end)
        total = len(uploaded_files)
        progress_text.markdown(f"Processing {total} files...")

        # Run all comparisons (processing progress is internal; we update UI after)
        zip_bytes, count = run_batch_comparison(uploaded_files)

        # Update UI on success
        progress_bar.progress(100)
        progress_text.markdown(f"✅ Processing complete — {count} result file(s) added to ZIP.")

        # Timestamped ZIP name
        zip_name = f"All_Allele_Comparison_Results_{timestamp_str()}.zip"
        st.download_button(
            label="📦 Download ZIP",
            data=zip_bytes.getvalue(),
            file_name=zip_name,
            mime="application/zip"
        )

    except Exception as e:
        progress_bar.progress(0)
        progress_text.markdown("")
        st.error(f"An error occurred: {e}")
