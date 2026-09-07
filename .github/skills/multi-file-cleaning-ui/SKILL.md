---
name: multi-file-cleaning-ui
description: "Use when extending the Streamlit data-cleaning UI from single-file processing to multiple-file workflows such as merge, chained join, and automatic detection."
argument-hint: "Describe the multi-file UI phase or workflow to implement."
---

# Multi-File Cleaning UI

## Purpose

Guide incremental changes to the Streamlit UI so multiple-file workflows are introduced without breaking the existing cleaning configuration or confusing the user.

## Procedure

1. Inspect the current uploader and identify whether downstream code expects one uploaded object or a list.
2. Change the uploader to `accept_multiple_files=True` and treat the result as a list.
3. Add one mutually exclusive work-mode selector with these labels:
   - `Single File`
   - `Gabungkan (Merge)`
   - `Join Berantai`
   - `Deteksi Otomatis`
4. Validate the file count before selecting a processing path:
   - In `Single File`, warn when more than one file is uploaded and process only the first file until multi-file processing is implemented.
   - In every multi-file mode, require at least two files and explain how to recover when fewer are uploaded.
5. Keep the sidebar configuration outside mode-specific branches. Theme, currency, missing-value strategy, text rules, and export settings must be shared by every mode.
6. Render only the controls and result surface for the selected mode. Do not show all mode-specific workflows at once.
7. Preserve the existing single-file preview, audit trail, and download behavior while the other modes are being introduced incrementally.

## Quality Checks

- The app compiles with `python -m py_compile app.py`.
- The uploader accepts multiple files.
- A single-file selection with multiple uploads shows an explicit warning naming the count and the recommended alternative modes.
- A multi-file mode with fewer than two uploads shows an actionable warning.
- The sidebar settings are defined once and remain available regardless of the selected mode.
- The sample-data button remains limited to `Single File` until sample handling is defined for multi-file modes.