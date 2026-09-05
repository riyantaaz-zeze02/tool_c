"""
engine.py — Pipeline Orkestrator Data Cleaning
===============================================
Mengatur urutan pembersihan data, mengelola konfigurasi,
dan mencatat audit trail (laporan statistik pembersihan).
"""

import pandas as pd
import os
from cleaner.reader import baca_data
from cleaner.rules.missing_handler import tangani_missing
from cleaner.rules.duplicate_handler import hapus_duplikat
from cleaner.rules.text_cleaner import bersihkan_teks
from cleaner.rules.number_cleaner import bersihkan_angka
from cleaner.rules.date_cleaner import bersihkan_tanggal
from cleaner.formatter.excel_styler import ExcelStyler
from cleaner.reporter import CleaningReporter
import openpyxl


class CleaningEngine:
    """
    Orkestrator pipeline pembersihan data.
    """

    DEFAULT_CONFIG = {
        "missing": {
            "numeric_strategy": "median",
            "text_strategy": "constant",
            "numeric_constant": 0,
            "text_constant": "Unknown",
            "drop_threshold": None,
        },
        "duplicate": {
            "subset": None,
            "keep": "first",
        },
        "number": {
            "kolom_angka": [],
            "kolom_currency": [],
            "kolom_persen": [],
            "auto_detect": True,
        },
        "date": {
            "kolom_tanggal": [],
            "auto_detect": True,
            "output_format": None,
        },
        "text": {
            "case_format": "title",
            "only_name_columns": True,
            "kolom_case": [],
            "preserve_brackets": True,
            "preserve_acronyms": True,
            "custom_acronyms": [],
            "strip_whitespace": True,
            "normalize_unicode": True,
            "remove_non_ascii": False,
            "kolom_exclude": ["email", "e-mail"],
        },
    }

    def __init__(self, config=None):
        """
        Inisialisasi CleaningEngine dengan konfigurasi kustom atau default.
        """
        self.config = {}
        # Merge default config dengan custom config
        for step, defaults in self.DEFAULT_CONFIG.items():
            self.config[step] = defaults.copy()
            if config and step in config:
                self.config[step].update(config[step])

        self.last_report = {}

    def clean(self, data_or_path, sheet_name=0):
        """
        Menjalankan seluruh pipeline pembersihan pada file atau DataFrame.

        Parameter:
            data_or_path (str | pd.DataFrame): Path file atau objek DataFrame.
            sheet_name (str | int): Sheet Excel yang dibaca jika berupa file Excel.

        Return:
            tuple: (pd.DataFrame bersih, dict laporan_audit)
        """
        source_name = "DataFrame"
        if isinstance(data_or_path, str):
            source_name = os.path.basename(data_or_path)
            df = baca_data(data_or_path, sheet_name=sheet_name)
        elif isinstance(data_or_path, pd.DataFrame):
            df = data_or_path.copy()
        else:
            raise TypeError("data_or_path harus berupa string path file atau pd.DataFrame")

        if df is None:
            raise ValueError("Gagal membaca data.")

        initial_rows = len(df)
        initial_cols = len(df.columns)

        report = {
            "sumber": source_name,
            "baris_awal": initial_rows,
            "kolom_awal": initial_cols,
            "langkah": {},
        }

        # 1. Parsing Angka & Mata Uang (sebelum teks distrip/dimodifikasi)
        df, log_number = bersihkan_angka(df, config=self.config.get("number", {}))
        report["langkah"]["number"] = log_number

        # 2. Parsing Tanggal
        df, log_date = bersihkan_tanggal(df, config=self.config.get("date", {}))
        report["langkah"]["date"] = log_date

        # 3. Tangani Missing Value
        df, log_missing = tangani_missing(df, config=self.config.get("missing", {}))
        report["langkah"]["missing"] = log_missing

        # 4. Hapus Duplikat
        df, log_duplicate = hapus_duplikat(df, config=self.config.get("duplicate", {}))
        report["langkah"]["duplicate"] = log_duplicate

        # 5. Standarisasi Teks
        df, log_text = bersihkan_teks(df, config=self.config.get("text", {}))
        report["langkah"]["text"] = log_text

        # Ringkasan Akhir
        report["baris_akhir"] = len(df)
        report["kolom_akhir"] = len(df.columns)
        report["total_duplikat_dihapus"] = log_duplicate.get("jumlah_duplikat", 0)
        report["total_baris_didrop"] = log_missing.get("baris_didrop", 0)

        self.last_report = report
        return df, report

    def export_excel(
        self,
        df,
        output_path,
        report=None,
        theme="corporate_blue",
        currency="IDR",
        sheet_name="Data Bersih",
        add_total_row=True,
        include_summary_sheet=True,
    ):
        """
        Mengekspor DataFrame bersih ke file Excel dengan styling profesional dan audit trail.

        Parameter:
            df (pd.DataFrame): DataFrame yang sudah dibersihkan.
            output_path (str): Lokasi file output (.xlsx).
            report (dict|None): Dict audit report dari self.clean().
            theme (str): Tema warna Excel ('corporate_blue', 'modern_slate', 'emerald_green', dll).
            currency (str): Mata uang ('IDR', 'USD').
            sheet_name (str): Nama worksheet data utama.
            add_total_row (bool): Apakah menyertakan baris Total dengan formula =SUM.
            include_summary_sheet (bool): Apakah menyertakan sheet kedua 'Cleaning Summary'.
        """
        # Pastikan direktori output ada
        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.views.sheetView[0].showGridLines = True

        # Tulis Header
        for col_idx, col_name in enumerate(df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=col_name)

        # Tulis Baris Data
        for row_idx, row_values in enumerate(df.itertuples(index=False), start=2):
            for col_idx, val in enumerate(row_values, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if pd.isna(val):
                    cell.value = None
                elif hasattr(val, "to_pydatetime"):
                    # Konversi Timestamp pandas ke datetime native python
                    cell.value = val.to_pydatetime()
                else:
                    cell.value = val

        # Terapkan Styling Profesional
        styler = ExcelStyler(theme=theme, currency=currency)
        styler.style_worksheet(ws, df, add_total_row=add_total_row)

        # Tambahkan Sheet Cleaning Summary jika diminta
        if include_summary_sheet and (report or self.last_report):
            active_report = report if report is not None else self.last_report
            reporter = CleaningReporter(theme=theme)
            reporter.add_summary_sheet(wb, active_report, df)

        wb.save(output_path)
        print(f"💾 File Excel profesional berhasil disimpan ke: {output_path}")
        return output_path
