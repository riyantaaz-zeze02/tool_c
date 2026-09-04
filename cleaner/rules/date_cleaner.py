"""
date_cleaner.py — Parsing & Normalisasi Tanggal
=================================================
Parsing otomatis berbagai format tanggal ke datetime standar.
"""

import pandas as pd
import re
import warnings


# Daftar format tanggal yang sering ditemui
COMMON_DATE_FORMATS = [
    "%Y-%m-%d",           # 2024-01-15
    "%d/%m/%Y",           # 15/01/2024
    "%m/%d/%Y",           # 01/15/2024
    "%d-%m-%Y",           # 15-01-2024
    "%d %B %Y",           # 15 January 2024
    "%d %b %Y",           # 15 Jan 2024
    "%Y/%m/%d",           # 2024/01/15
    "%d.%m.%Y",           # 15.01.2024
    "%B %d, %Y",          # January 15, 2024
    "%b %d, %Y",          # Jan 15, 2024
    "%Y-%m-%d %H:%M:%S",  # 2024-01-15 10:30:00
    "%d/%m/%Y %H:%M",     # 15/01/2024 10:30
]

# Kata kunci nama kolom yang kemungkinan berisi tanggal
DATE_COLUMN_KEYWORDS = [
    "tanggal", "date", "tgl", "waktu", "time",
    "created", "updated", "modified", "timestamp",
    "lahir", "birth", "expired", "deadline",
    "mulai", "selesai", "start", "end",
]


def bersihkan_tanggal(df, config=None):
    """
    Mendeteksi dan mengkonversi kolom tanggal ke datetime standar.

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan.
        config (dict|None): Konfigurasi.
            - kolom_tanggal: list kolom yang harus diparse sebagai tanggal
            - auto_detect: bool (True = deteksi otomatis berdasarkan nama kolom & isi)
            - output_format: str format output (default: None = pandas datetime)

    Return:
        tuple: (pd.DataFrame bersih, dict log perubahan)
    """
    if config is None:
        config = {}

    df = df.copy()
    kolom_tanggal = list(config.get("kolom_tanggal", []))
    auto_detect = config.get("auto_detect", True)
    # output_format disimpan di log untuk referensi formatter nanti
    output_format = config.get("output_format", None)

    log = {"kolom_diproses": [], "aksi": [], "output_format": output_format}

    print("📅 Membersihkan format tanggal...")

    # Auto-detect kolom tanggal
    if auto_detect:
        for kolom in df.columns:
            if kolom in kolom_tanggal:
                continue

            nama_lower = kolom.lower()

            # Cek berdasarkan nama kolom
            if any(kw in nama_lower for kw in DATE_COLUMN_KEYWORDS):
                kolom_tanggal.append(kolom)
                continue

            # Cek berdasarkan isi (sampling)
            if df[kolom].dtype == "object" or pd.api.types.is_string_dtype(df[kolom]):
                sample = df[kolom].dropna().head(10).astype(str)
                if len(sample) > 0 and _looks_like_date(sample):
                    kolom_tanggal.append(kolom)

    # Proses setiap kolom tanggal
    for kolom in kolom_tanggal:
        if kolom not in df.columns:
            continue

        # Skip jika sudah datetime
        if pd.api.types.is_datetime64_any_dtype(df[kolom]):
            continue

        converted = _convert_to_datetime(df[kolom])
        if converted is not None:
            df[kolom] = converted
            log["kolom_diproses"].append(kolom)
            log["aksi"].append(f"'{kolom}' → datetime")
            print(f"   📅 Kolom '{kolom}' → dikonversi ke datetime")

    if not log["kolom_diproses"]:
        print("   ✅ Tidak ada kolom tanggal yang perlu dikonversi.")
    else:
        print(
            f"   ✅ Selesai! {len(log['kolom_diproses'])} kolom tanggal dikonversi."
        )
    print()

    return df, log


def _looks_like_date(series):
    """Heuristik cek apakah suatu Series kemungkinan berisi tanggal."""
    date_patterns = [
        r"\d{4}[-/]\d{1,2}[-/]\d{1,2}",  # YYYY-MM-DD atau YYYY/MM/DD
        r"\d{1,2}[-/]\d{1,2}[-/]\d{4}",  # DD/MM/YYYY atau MM/DD/YYYY
        r"\d{1,2}\s+\w+\s+\d{4}",         # 15 January 2024
        r"\d{1,2}\.\d{1,2}\.\d{4}",       # 15.01.2024
    ]

    match_count = 0
    for val in series:
        val_str = str(val).strip()
        for pattern in date_patterns:
            if re.search(pattern, val_str):
                match_count += 1
                break

    return match_count / len(series) > 0.5


def _convert_to_datetime(series):
    """
    Coba konversi Series ke datetime dengan berbagai format.
    Return None jika gagal total.
    """
    # Coba pandas auto-parse dengan format='mixed' dan dayfirst=True
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            result = pd.to_datetime(series, errors="coerce", format="mixed", dayfirst=True)
        valid_count = series.dropna().count()
        if valid_count > 0 and (result.notna().sum() / valid_count) >= 0.5:
            return result
    except Exception:
        pass

    # Coba format-format yang umum satu per satu
    for fmt in COMMON_DATE_FORMATS:
        try:
            result = pd.to_datetime(series, format=fmt, errors="coerce")
            valid_count = series.dropna().count()
            if valid_count > 0 and (result.notna().sum() / valid_count) >= 0.5:
                return result
        except Exception:
            continue

    return None
