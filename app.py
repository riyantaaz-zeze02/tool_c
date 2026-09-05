"""
app.py — ExcelCleaner Pro Web Studio (Streamlit)
================================================
Dashboard web interaktif untuk pembersihan data dan formatting Excel.
Jalankan dengan: streamlit run app.py
"""

import io
import os
import pandas as pd
import streamlit as st

from cleaner.engine import CleaningEngine
from cleaner.formatter.themes import THEMES, get_theme
from cleaner.reader import read_csv_with_fallback


# Konfigurasi Halaman
st.set_page_config(
    page_title="ExcelCleaner Pro",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #595959;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 16px;
        border-left: 4px solid #1F4E79;
        margin-bottom: 12px;
    }
    .stDownloadButton button {
        background-color: #217346 !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        padding: 0.5rem 1.2rem !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🧹 ExcelCleaner Pro — Web Studio</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Bersihkan data kotor, standarisasi format, dan hasilkan spreadsheet Excel siap pakai berstandar korporat.</div>', unsafe_allow_html=True)

# SIDEBAR: Panel Kontrol
with st.sidebar:
    st.header("⚙️ Pengaturan Studio")

    st.subheader("🎨 Tema & Tampilan Excel")
    theme_choice = st.selectbox(
        "Pilih Tema Warna:",
        options=list(THEMES.keys()),
        format_func=lambda x: THEMES[x]["name"],
        index=0,
    )
    currency_choice = st.selectbox(
        "Mata Uang Default:",
        options=["IDR", "USD"],
        index=0,
    )
    add_total_row = st.checkbox("Tambahkan Baris TOTAL (=SUM)", value=True)
    include_summary_sheet = st.checkbox("Sertakan Sheet 'Cleaning Summary'", value=True)

    st.markdown("---")
    st.subheader("🧹 Aturan Pembersihan")
    num_strat = st.selectbox(
        "Pengisian Nilai Kosong (Angka):",
        options=["median", "mean", "mode", "constant"],
        index=0,
    )
    text_case = st.selectbox(
        "Kapitalisasi Teks:",
        options=["title", "upper", "lower", "sentence"],
        format_func=lambda x: {
            "title": "Title Case (Budi Santoso)",
            "upper": "UPPERCASE (BUDI SANTOSO)",
            "lower": "lowercase (budi santoso)",
            "sentence": "Sentence case (Budi santoso)",
        }[x],
        index=0,
    )
    only_name_cols = st.checkbox(
        "Hanya Kolom Nama/Judul untuk Title Case",
        value=True,
        help="Hanya terapkan Title Case ke kolom nama/judul (misal: Artist, Tour title, Nama, Kota). Menghindari deskripsi/catatan terkapitalisasi semua.",
    )
    preserve_brackets = st.checkbox(
        "Pertahankan Kurung Siku [LIVE] & Akronim (VIP)",
        value=True,
        help="Menjaga kata dalam [kurung siku] dan akronim all-caps (VIP, USA, IDR, dll.) agar tidak berubah.",
    )
    strip_spaces = st.checkbox("Hapus Spasi Berlebih & Unicode", value=True)

# UPLOAD FILE ATAU GUNAKAN CONTOH
col_upload, col_sample = st.columns([3, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "Unggah File Data (.csv atau .xlsx)",
        type=["csv", "xlsx", "xls"],
        help="Mendukung format CSV dan Excel",
    )

with col_sample:
    st.write("")
    st.write("")
    use_sample = st.button("📂 Gunakan Data Contoh")

# Baca Data
df_raw = None
file_label = ""

if uploaded_file is not None:
    file_label = uploaded_file.name
    try:
        if uploaded_file.name.endswith(".csv"):
            df_raw = read_csv_with_fallback(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")

elif use_sample:
    sample_path = os.path.join("data", "input", "contoh_data.csv")
    if os.path.exists(sample_path):
        df_raw = read_csv_with_fallback(sample_path)
        file_label = "contoh_data.csv"

    else:
        st.warning("File contoh_data.csv tidak ditemukan di data/input/")

# PROSES CLEANING & PREVIEW
if df_raw is not None:
    # Buat Konfigurasi Engine Dinamis
    config = {
        "missing": {
            "numeric_strategy": num_strat,
            "text_strategy": "constant",
            "text_constant": "Unknown",
            "datetime_strategy": "keep_nat",
        },
        "duplicate": {
            "subset": None,
            "keep": "first",
        },
        "number": {
            "auto_detect": True,
        },
        "date": {
            "auto_detect": True,
        },
        "text": {
            "case_format": text_case,
            "only_name_columns": only_name_cols,
            "preserve_brackets": preserve_brackets,
            "preserve_acronyms": preserve_brackets,
            "strip_whitespace": strip_spaces,
            "normalize_unicode": strip_spaces,
            "kolom_exclude": ["email", "e-mail"],
        },
        "formatting": {
            "theme": theme_choice,
            "currency": currency_choice,
            "add_total_row": add_total_row,
            "include_summary_sheet": include_summary_sheet,
        },
    }

    engine = CleaningEngine(config=config)
    df_clean, report = engine.clean(df_raw)

    # METRICS DISPLAY
    st.markdown("---")
    st.subheader(f"📊 Hasil Pembersihan: {file_label}")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Baris Awal", report["baris_awal"])
    with m2:
        st.metric("Baris Akhir (Bersih)", report["baris_akhir"])
    with m3:
        st.metric("Duplikat Dihapus", report["total_duplikat_dihapus"])
    with m4:
        missing_sebelum = sum(report["langkah"]["missing"]["missing_sebelum"].values())
        st.metric("Missing Values Diperbaiki", missing_sebelum)

    # TABS: Preview Sebelum vs Sesudah & Audit Trail
    tab1, tab2, tab3 = st.tabs(["✨ Data Bersih", "🔍 Sebelum vs Sesudah", "📋 Lembar Audit Trail"])

    with tab1:
        st.dataframe(df_clean, use_container_width=True)

        # Generate file Excel ke in-memory buffer untuk tombol download
        temp_out = os.path.join("data", "output", "_temp_web_export.xlsx")
        engine.export_excel(
            df_clean,
            temp_out,
            report=report,
            theme=theme_choice,
            currency=currency_choice,
            add_total_row=add_total_row,
            include_summary_sheet=include_summary_sheet,
        )

        with open(temp_out, "rb") as f:
            excel_bytes = f.read()

        base_clean_name = os.path.splitext(file_label)[0] + "_bersih.xlsx"
        st.download_button(
            label=f"📥 Unduh Excel Rapi ({base_clean_name})",
            data=excel_bytes,
            file_name=base_clean_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with tab2:
        c_kiri, c_kanan = st.columns(2)
        with c_kiri:
            st.markdown("**Data Asli (Kotor):**")
            st.dataframe(df_raw.head(15), use_container_width=True)
        with c_kanan:
            st.markdown("**Data Hasil Pembersihan:**")
            st.dataframe(df_clean.head(15), use_container_width=True)

    with tab3:
        st.markdown("**Aksi yang Dilakukan pada Setiap Kolom:**")
        audit_rows = []
        for step in ["number", "date", "missing", "text"]:
            step_actions = report["langkah"].get(step, {}).get("aksi", [])
            for a in step_actions:
                audit_rows.append({"Tahap": step.capitalize(), "Aksi": a})

        if audit_rows:
            st.table(pd.DataFrame(audit_rows))
        else:
            st.info("Data input sudah sangat bersih, tidak ada transformasi agresif.")

else:
    st.info("👆 Silakan upload file data (.csv / .xlsx) di atas atau klik 'Gunakan Data Contoh' untuk memulai.")
