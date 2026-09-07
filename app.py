"""
app.py — ExcelCleaner Pro Web Studio (Streamlit)
================================================
Dashboard web interaktif untuk pembersihan data dan formatting Excel.
Jalankan dengan: streamlit run app.py
"""

import io
import os
import shutil
import tempfile

import pandas as pd
import streamlit as st
import yaml

from cli import analyze_auto_files
from cleaner.engine import CleaningEngine
from cleaner.formatter.themes import THEMES
from cleaner.reader import get_sheet_names, read_csv_with_fallback


def read_uploaded_dataframe(uploaded_file):
    """Read an uploaded file without depending on its current stream position."""
    file_buffer = io.BytesIO(uploaded_file.getvalue())
    if uploaded_file.name.lower().endswith(".csv"):
        return read_csv_with_fallback(file_buffer)
    return pd.read_excel(file_buffer)


def save_uploaded_files(uploaded_files, selected_sheets=None):
    """Materialize Streamlit uploads for engine APIs that require file paths."""
    temp_dir = tempfile.mkdtemp(prefix="excelcleaner_merge_")
    file_paths = []
    try:
        for index, uploaded_file in enumerate(uploaded_files):
            suffix = os.path.splitext(uploaded_file.name)[1].lower()
            selected_sheet = selected_sheets[index] if selected_sheets else None
            if selected_sheets is not None and suffix in {".xlsx", ".xls"}:
                selected_df = pd.read_excel(io.BytesIO(uploaded_file.getvalue()), sheet_name=selected_sheet)
                temp_path = os.path.join(temp_dir, f"{os.path.splitext(uploaded_file.name)[0]}.csv")
                selected_df.to_csv(temp_path, index=False)
            else:
                temp_path = os.path.join(temp_dir, f"{len(file_paths)}{suffix}")
                with open(temp_path, "wb") as output_file:
                    output_file.write(uploaded_file.getvalue())
            file_paths.append(temp_path)
        return temp_dir, file_paths
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise


def get_uploaded_sheet_names(uploaded_file):
    """Return Excel sheet names from an uploaded file, or an empty list for CSV."""
    if not uploaded_file.name.lower().endswith((".xlsx", ".xls")):
        return []
    return list(pd.ExcelFile(io.BytesIO(uploaded_file.getvalue())).sheet_names)


def load_uploaded_config(uploaded_config):
    """Parse an optional YAML upload into a configuration mapping."""
    if uploaded_config is None:
        return {}
    parsed = yaml.safe_load(uploaded_config.getvalue().decode("utf-8")) or {}
    if not isinstance(parsed, dict):
        raise ValueError("Config YAML harus berisi mapping/object di level teratas.")
    return parsed


def set_work_mode(mode):
    """Update the mode radio safely from an action button callback."""
    st.session_state["work_mode"] = mode


def build_config(theme_choice, currency_choice, add_total_row, include_summary_sheet,
                 num_strat, text_case, only_name_cols, preserve_brackets, strip_spaces):
    return {
        "missing": {
            "numeric_strategy": num_strat,
            "text_strategy": "constant",
            "text_constant": "Unknown",
            "datetime_strategy": "keep_nat",
        },
        "duplicate": {"subset": None, "keep": "first"},
        "number": {"auto_detect": True},
        "date": {"auto_detect": True},
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


st.set_page_config(page_title="ExcelCleaner Pro", page_icon="🧹", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1F4E79; margin-bottom: 0.2rem; }
    .sub-header { font-size: 1.05rem; color: #595959; margin-bottom: 1.5rem; }
    .metric-card { background-color: #F8F9FA; border-radius: 8px; padding: 16px; border-left: 4px solid #1F4E79; margin-bottom: 12px; }
    .stDownloadButton button { background-color: #217346 !important; color: white !important; font-weight: 600 !important; border-radius: 6px !important; padding: 0.5rem 1.2rem !important; }
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="main-header">🧹 ExcelCleaner Pro — Web Studio</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Bersihkan data kotor, standarisasi format, dan hasilkan spreadsheet Excel siap pakai berstandar korporat.</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Pengaturan Studio")
    st.subheader("🎨 Tema & Tampilan Excel")
    theme_choice = st.selectbox("Pilih Tema Warna:", options=list(THEMES.keys()), format_func=lambda x: THEMES[x]["name"], index=0)
    currency_choice = st.selectbox("Mata Uang Default:", options=["IDR", "USD"], index=0)
    add_total_row = st.checkbox("Tambahkan Baris TOTAL (=SUM)", value=True)
    include_summary_sheet = st.checkbox("Sertakan Sheet 'Cleaning Summary'", value=True)
    st.markdown("---")
    st.subheader("🧹 Aturan Pembersihan")
    num_strat = st.selectbox("Pengisian Nilai Kosong (Angka):", options=["median", "mean", "mode", "constant"], index=0)
    text_case = st.selectbox(
        "Kapitalisasi Teks:",
        options=["title", "upper", "lower", "sentence"],
        format_func=lambda x: {"title": "Title Case (Budi Santoso)", "upper": "UPPERCASE (BUDI SANTOSO)", "lower": "lowercase (budi santoso)", "sentence": "Sentence case (Budi santoso)"}[x],
        index=0,
    )
    only_name_cols = st.checkbox("Hanya Kolom Nama/Judul untuk Title Case", value=True)
    preserve_brackets = st.checkbox("Pertahankan Kurung Siku [LIVE] & Akronim (VIP)", value=True)
    strip_spaces = st.checkbox("Hapus Spasi Berlebih & Unicode", value=True)
    output_format = st.radio("Format Output", options=["Excel", "CSV"], horizontal=True)
    uploaded_config = st.file_uploader("Config YAML (opsional)", type=["yaml", "yml"], key="config_upload")

mode_options = ["Single File", "Gabungkan (Merge)", "Join Berantai", "Deteksi Otomatis"]
work_mode = st.radio(
    "Mode Kerja",
    options=mode_options,
    horizontal=True,
    key="work_mode",
    help="Pilih satu alur kerja sebelum memproses file.",
)
col_upload, col_sample = st.columns([3, 1])
with col_upload:
    uploaded_files = st.file_uploader("Unggah File Data (.csv atau .xlsx)", type=["csv", "xlsx", "xls"], help="Mendukung format CSV dan Excel", accept_multiple_files=True)
with col_sample:
    st.write("")
    st.write("")
    use_sample = st.button("📂 Gunakan Data Contoh")


df_raw = None
file_label = ""
merge_temp_dir = None
merge_file_paths = []
merge_sheet_names = []
join_temp_dir = None
join_file_paths = []
join_keys = []
join_sheet_names = []
auto_temp_dir = None
auto_file_paths = []
single_temp_dir = None
single_file_path = None
single_selected_sheets = None
single_all_sheets = False
cleaned_sheet_results = None
cleaned_sheet_reports = None

if work_mode == "Single File" and len(uploaded_files) > 1:
    st.warning(f"Kamu pilih mode Single File tapi upload {len(uploaded_files)} file — hanya file pertama yang akan diproses, atau ganti mode ke Merge/Join.")
elif work_mode != "Single File" and uploaded_files and len(uploaded_files) < 2:
    st.warning(f"Mode {work_mode} membutuhkan minimal 2 file. Upload file tambahan atau pilih mode Single File.")

if uploaded_files and work_mode == "Single File":
    single_file_index = 0
    if len(uploaded_files) > 1:
        single_file_index = st.selectbox(
            "File yang diproses sekarang",
            options=range(len(uploaded_files)),
            format_func=lambda index: uploaded_files[index].name,
        )
    uploaded_file = uploaded_files[single_file_index]
    file_label = uploaded_file.name
    try:
        sheet_names = get_uploaded_sheet_names(uploaded_file)
        if len(sheet_names) > 1:
            st.subheader(f"Pilih Sheet: {uploaded_file.name}")
            single_all_sheets = st.checkbox("Pilih Semua Sheet", value=True, key="single_all_sheets")
            single_selected_sheets = st.multiselect(
                "Sheet yang dibersihkan",
                options=sheet_names,
                default=sheet_names if single_all_sheets else sheet_names[:1],
                key="single_selected_sheets",
            )
            if not single_selected_sheets:
                st.error("Pilih minimal satu sheet untuk diproses.")
            else:
                single_temp_dir, single_paths = save_uploaded_files([uploaded_file])
                single_file_path = single_paths[0]
                df_raw = pd.read_excel(single_file_path, sheet_name=single_selected_sheets[0])
        else:
            df_raw = read_uploaded_dataframe(uploaded_file)
    except Exception as error:
        st.error(f"Gagal membaca file: {error}")
elif uploaded_files and work_mode == "Gabungkan (Merge)" and len(uploaded_files) >= 2:
    try:
        preview_rows = []
        preview_frames = []
        for uploaded_file in uploaded_files:
            available_sheets = get_uploaded_sheet_names(uploaded_file)
            selected_sheet = available_sheets[0] if available_sheets else 0
            if len(available_sheets) > 1:
                selected_sheets = st.multiselect(
                    f"Sheet untuk {uploaded_file.name}",
                    options=available_sheets,
                    default=available_sheets[:1],
                    key=f"merge_sheets_{uploaded_file.name}",
                )
                if not selected_sheets:
                    st.error(f"Pilih minimal satu sheet untuk {uploaded_file.name}.")
                    continue
                selected_sheet = selected_sheets[0]
            merge_sheet_names.append(selected_sheet)
            preview_df = read_uploaded_dataframe(uploaded_file) if not available_sheets else pd.read_excel(io.BytesIO(uploaded_file.getvalue()), sheet_name=selected_sheet)
            preview_frames.append((uploaded_file.name, preview_df))
            preview_rows.append({"Nama File": uploaded_file.name, "Jumlah Baris": len(preview_df), "Jumlah Kolom": len(preview_df.columns)})
        st.subheader("Preview File yang Akan Digabung")
        st.dataframe(pd.DataFrame(preview_rows), hide_index=True, use_container_width=True)
        if any(row["Jumlah Kolom"] <= 1 for row in preview_rows):
            st.warning("Ada file dengan satu kolom atau kurang. Periksa kembali file tersebut sebelum melanjutkan.")
        merge_temp_dir, merge_file_paths = save_uploaded_files(uploaded_files, merge_sheet_names)
        df_raw = pd.concat([df.assign(**{"Sumber File": name}) for name, df in preview_frames], ignore_index=True)
        file_label = f"gabungan_{len(uploaded_files)}file"
    except Exception as error:
        st.error(f"Gagal membaca file untuk merge: {error}")
elif uploaded_files and work_mode == "Join Berantai" and len(uploaded_files) >= 2:
    try:
        join_preview_frames = []
        for uploaded_file in uploaded_files:
            available_sheets = get_uploaded_sheet_names(uploaded_file)
            selected_sheet = available_sheets[0] if available_sheets else 0
            if len(available_sheets) > 1:
                selected_sheets = st.multiselect(
                    f"Sheet untuk {uploaded_file.name}",
                    options=available_sheets,
                    default=available_sheets[:1],
                    key=f"join_sheets_{uploaded_file.name}",
                )
                if not selected_sheets:
                    st.error(f"Pilih minimal satu sheet untuk {uploaded_file.name}.")
                    continue
                selected_sheet = selected_sheets[0]
            join_sheet_names.append(selected_sheet)
            preview_df = read_uploaded_dataframe(uploaded_file) if not available_sheets else pd.read_excel(io.BytesIO(uploaded_file.getvalue()), sheet_name=selected_sheet)
            join_preview_frames.append((uploaded_file.name, preview_df))

        st.subheader("Urutan File Join")
        for index, (file_name, preview_df) in enumerate(join_preview_frames, start=1):
            st.write(f"{index}. **{file_name}** ({len(preview_df)} baris, {len(preview_df.columns)} kolom)")

        selected_join_keys = []
        join_keys_valid = True
        for index, ((left_name, left_df), (right_name, right_df)) in enumerate(
            zip(join_preview_frames, join_preview_frames[1:]),
            start=1,
        ):
            shared_columns = [column for column in left_df.columns if column in right_df.columns]
            st.markdown(f"**Tahap {index}: {left_name} → {right_name}**")
            if not shared_columns:
                st.error(f"Tidak ada kolom kunci yang sama antara {left_name} dan {right_name}.")
                join_keys_valid = False
                selected_join_keys.append("")
                continue
            selected_key = st.selectbox(
                f"Pilih key tahap {index}",
                options=["Pilih key..."] + shared_columns,
                key=f"join_key_{index}",
            )
            if selected_key == "Pilih key...":
                st.error(f"Key untuk tahap {index} belum dipilih.")
                join_keys_valid = False
                selected_join_keys.append("")
            else:
                selected_join_keys.append(selected_key)

        join_process = st.button("Proses Join", disabled=not join_keys_valid)
        if join_process:
            for index, key in enumerate(selected_join_keys):
                left_columns = join_preview_frames[index][1].columns
                right_columns = join_preview_frames[index + 1][1].columns
                if not key or key not in left_columns or key not in right_columns:
                    st.error(f"Key untuk tahap {index + 1} harus tersedia di kedua file pasangan.")
                    join_keys_valid = False
            if join_keys_valid:
                join_temp_dir, join_file_paths = save_uploaded_files(uploaded_files, join_sheet_names)
                join_keys = selected_join_keys
                df_raw = join_preview_frames[0][1]
                file_label = f"join_{len(uploaded_files)}file"
    except Exception as error:
        st.error(f"Gagal menyiapkan file untuk join: {error}")
elif uploaded_files and work_mode == "Deteksi Otomatis" and len(uploaded_files) >= 2:
    try:
        auto_temp_dir, auto_file_paths = save_uploaded_files(uploaded_files)
        analysis = analyze_auto_files(auto_file_paths)
        header_lines = [
            f"- **{uploaded_file.name}**: {', '.join(analysis['headers'][path]) or '(tanpa kolom)'}"
            for uploaded_file, path in zip(uploaded_files, auto_file_paths)
        ]
        candidate_lines = []
        auto_name_by_path = {
            path: uploaded_file.name
            for uploaded_file, path in zip(uploaded_files, auto_file_paths)
        }
        for key, paths in analysis["candidate_keys"].items():
            names = ", ".join(auto_name_by_path.get(path, os.path.basename(path)) for path in paths)
            candidate_lines.append(f"- `{key}`: {names}")
        recommendation = "Merge" if analysis["mode"] == "merge" else "Join" if analysis["mode"] == "join" else "Belum ada rekomendasi aman"
        auto_message = "\n".join([
            "**Kolom per file:**",
            *header_lines,
            "",
            f"**Rekomendasi:** {recommendation}",
            f"**Alasan:** {analysis['reason']}",
            "",
            "**Kandidat key:**",
            *(candidate_lines or ["- Tidak ada kandidat key bersama."]),
        ])
        st.info(auto_message)
        if analysis["mode"] in {"merge", "join"}:
            recommended_mode = "Gabungkan (Merge)" if analysis["mode"] == "merge" else "Join Berantai"
            st.button(
                "✅ Lanjutkan dengan rekomendasi ini",
                on_click=set_work_mode,
                args=(recommended_mode,),
            )
            st.button(
                "✏️ Pilih manual",
                on_click=set_work_mode,
                args=("Single File",),
            )
        else:
            st.warning("Struktur file ambigu. Pilih mode Merge atau Join secara manual setelah meninjau kandidat key.")
    except Exception as error:
        st.error(f"Gagal menganalisis file: {error}")
    finally:
        if auto_temp_dir:
            shutil.rmtree(auto_temp_dir, ignore_errors=True)
elif uploaded_files and len(uploaded_files) >= 2:
    st.info(f"{work_mode} dipilih untuk {len(uploaded_files)} file. Alur pemrosesan mode ini akan tersedia pada fase berikutnya. Semua konfigurasi sidebar tetap aktif untuk mode ini.")
elif use_sample and work_mode == "Single File":
    sample_path = os.path.join("data", "input", "contoh_data.csv")
    if os.path.exists(sample_path):
        df_raw = read_csv_with_fallback(sample_path)
        file_label = "contoh_data.csv"
    else:
        st.warning("File contoh_data.csv tidak ditemukan di data/input/")

if df_raw is not None:
    config = build_config(theme_choice, currency_choice, add_total_row, include_summary_sheet, num_strat, text_case, only_name_cols, preserve_brackets, strip_spaces)
    try:
        custom_config = load_uploaded_config(uploaded_config)
        for section, values in custom_config.items():
            if isinstance(values, dict) and isinstance(config.get(section), dict):
                config[section].update(values)
            else:
                config[section] = values
    except Exception as error:
        st.error(f"Gagal membaca config YAML: {error}")
        st.stop()
    formatting_config = config.get("formatting", {})
    effective_theme = formatting_config.get("theme", theme_choice)
    effective_currency = formatting_config.get("currency", currency_choice)
    effective_add_total = formatting_config.get("add_total_row", add_total_row)
    effective_include_summary = formatting_config.get("include_summary_sheet", include_summary_sheet)
    engine = CleaningEngine(config=config)
    if work_mode == "Single File" and single_file_path:
        cleaned_sheet_results = engine.clean_sheets(
            single_file_path,
            sheets=single_selected_sheets,
            all_sheets=single_all_sheets,
        )
        cleaned_sheet_reports = {name: result["report"] for name, result in cleaned_sheet_results.items()}
        first_sheet = next(iter(cleaned_sheet_results.values()))
        df_clean, report = first_sheet["df"], first_sheet["report"]
    elif work_mode == "Gabungkan (Merge)":
        try:
            engine.validate_merged_files(merge_file_paths)
            df_clean, report = engine.clean_merged_files(
                merge_file_paths,
                sheet_name=0,
            )
        except Exception as error:
            st.error(str(error))
            st.stop()
        finally:
            if merge_temp_dir:
                shutil.rmtree(merge_temp_dir, ignore_errors=True)
    elif work_mode == "Join Berantai":
        try:
            df_clean, report = engine.clean_joined_files(
                join_file_paths,
                join_keys,
                sheet_name=0,
            )
        except Exception as error:
            st.error(str(error))
            st.stop()
        finally:
            if join_temp_dir:
                shutil.rmtree(join_temp_dir, ignore_errors=True)
    else:
        df_clean, report = engine.clean(df_raw)

    st.markdown("---")
    st.subheader(f"📊 Hasil Pembersihan: {file_label}")
    metric_columns = st.columns(5 if report.get("is_merge") or report.get("is_join") else 4)
    with metric_columns[0]:
        st.metric("Baris Awal", report["baris_awal"])
    with metric_columns[1]:
        st.metric("Baris Akhir (Bersih)", report["baris_akhir"])
    if report.get("is_merge"):
        with metric_columns[2]:
            st.metric("Duplikat Antar File", report["duplikat_antar_file"])
        with metric_columns[3]:
            st.metric("Duplikat Dalam File", report["duplikat_internal"])
        missing_metric_column = metric_columns[4]
    elif report.get("is_join"):
        with metric_columns[2]:
            st.metric("Baris Ter-drop", report["total_orphan_rows"])
        with metric_columns[3]:
            st.metric("Duplikat Dihapus", report["total_duplikat_dihapus"])
        missing_metric_column = metric_columns[4]
    else:
        with metric_columns[2]:
            st.metric("Duplikat Dihapus", report["total_duplikat_dihapus"])
        missing_metric_column = metric_columns[3]
    with missing_metric_column:
        missing_sebelum = sum(report["langkah"]["missing"]["missing_sebelum"].values())
        st.metric("Missing Values Diperbaiki", missing_sebelum)

    tab_names = ["✨ Data Bersih", "🔍 Sebelum vs Sesudah", "📋 Lembar Audit Trail"]
    if report.get("is_join"):
        tab_names.append("⚠️ Baris Ter-drop")
    tabs = st.tabs(tab_names)
    tab1, tab2, tab3 = tabs[:3]
    with tab1:
        if report.get("is_join"):
            st.subheader("Match Rate per Tahap")
            match_rate_rows = [
                {
                    "Stage": stage["tahap"],
                    "Left": stage["left"],
                    "Right": stage["right"],
                    "Key": stage["key"],
                    "Match Rate": f"{stage['match_rate']:.2f}%",
                }
                for stage in report["join_stages"]
            ]
            st.dataframe(pd.DataFrame(match_rate_rows), hide_index=True, use_container_width=True)
        display_df = df_clean
        if report.get("is_merge"):
            source_options = ["Semua"] + sorted(df_clean["Sumber File"].dropna().unique().tolist())
            source_filter = st.selectbox("Filter Sumber File", source_options)
            if source_filter != "Semua":
                display_df = df_clean[df_clean["Sumber File"] == source_filter]
        st.dataframe(display_df, use_container_width=True)
        temp_out = os.path.join("data", "output", "_temp_web_export.xlsx")
        if output_format == "Excel":
            engine.export_excel(
                cleaned_sheet_results if cleaned_sheet_results else df_clean,
                temp_out,
                report=cleaned_sheet_reports if cleaned_sheet_reports else report,
                theme=effective_theme,
                currency=effective_currency,
                add_total_row=effective_add_total,
                include_summary_sheet=effective_include_summary,
                original_file_path=single_file_path,
            )
        else:
            df_clean.to_csv(temp_out.replace(".xlsx", ".csv"), index=False)
            temp_out = temp_out.replace(".xlsx", ".csv")
        with open(temp_out, "rb") as output_file:
            excel_bytes = output_file.read()
        if output_format == "CSV":
            download_mime = "text/csv"
            download_extension = ".csv"
        else:
            download_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            download_extension = ".xlsx"
        if report.get("is_merge"):
            base_clean_name = f"gabungan_{report['total_files']}file_bersih.xlsx"
        elif report.get("is_join"):
            base_clean_name = f"join_{report['total_files']}file_bersih.xlsx"
        else:
            base_clean_name = os.path.splitext(file_label)[0] + "_bersih" + download_extension
        if output_format == "CSV" and report.get("is_merge"):
            base_clean_name = f"gabungan_{report['total_files']}file_bersih.csv"
        elif output_format == "CSV" and report.get("is_join"):
            base_clean_name = f"join_{report['total_files']}file_bersih.csv"
        st.download_button(label=f"📥 Unduh Hasil ({base_clean_name})", data=excel_bytes, file_name=base_clean_name, mime=download_mime)
        if single_temp_dir:
            shutil.rmtree(single_temp_dir, ignore_errors=True)

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
            for action in report["langkah"].get(step, {}).get("aksi", []):
                audit_rows.append({"Tahap": step.capitalize(), "Aksi": action})
        if audit_rows:
            st.table(pd.DataFrame(audit_rows))
        else:
            st.info("Data input sudah sangat bersih, tidak ada transformasi agresif.")

    if report.get("is_join"):
        with tabs[3]:
            orphan_rows = report["orphan_rows"]
            if orphan_rows.empty:
                st.success("Semua baris berhasil di-join, tidak ada yang hilang.")
            else:
                st.dataframe(orphan_rows, hide_index=True, use_container_width=True)
elif use_sample:
    st.info("Data contoh hanya tersedia untuk mode Single File. Pilih mode tersebut untuk memprosesnya.")
elif work_mode == "Join Berantai" and uploaded_files:
    st.info("Pilih key untuk setiap pasangan file, lalu klik 'Proses Join'.")
else:
    st.info("👆 Silakan upload file data (.csv / .xlsx) di atas atau klik 'Gunakan Data Contoh' untuk memulai.")
