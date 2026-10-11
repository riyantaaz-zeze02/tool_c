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

        converted, n_ambigu = _convert_with_report(df[kolom])
        if converted is not None:
            df[kolom] = converted
            log["kolom_diproses"].append(kolom)
            log["aksi"].append(f"'{kolom}' → datetime")
            print(f"   📅 Kolom '{kolom}' → dikonversi ke datetime")
            if n_ambigu:
                pesan = (
                    f"'{kolom}': {n_ambigu} tanggal ambigu (hari/bulan sama-sama ≤12, "
                    f"tidak ada bukti urutan) — diasumsikan DD/MM/YYYY, mohon dicek"
                )
                log.setdefault("peringatan", []).append(pesan)
                log["aksi"].append("⚠️ " + pesan)
                print(f"   ⚠️  {pesan}")

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


_NUMERIC_DATE_RE = r"^(\d{1,4})([-/.])(\d{1,2})\2(\d{1,4})$"
_ISO_DATETIME_RE = r"^\d{4}-\d{1,2}-\d{1,2}[T ]\d"


def _convert_with_report(series):
    """
    Konversi Series ke datetime TANPA menebak urutan hari/bulan secara membabi buta.

    Aturan:
      - Tahun di depan (YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD) -> selalu tahun-bulan-hari
      - Nama bulan ("04 Jun 2025") -> tidak ambigu
      - Numerik hari/bulan di depan -> urutan disimpulkan per kelompok separator
        dari bukti nyata (ada angka >12 di posisi pertama = DD/MM, di posisi kedua = MM/DD)
      - Jika tidak ada bukti -> default DD/MM/YYYY dan baris tsb DIHITUNG sebagai ambigu

    Return: (Series datetime | None, jumlah_baris_ambigu)
    """
    import numpy as np

    valid = series.notna()
    s = series.astype(str).str.strip()
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    handled = pd.Series(False, index=series.index)
    n_ambigu = 0

    parts = s.str.extract(_NUMERIC_DATE_RE)
    numeric = valid & parts[0].notna()
    a_len = parts[0].str.len()
    c_len = parts[3].str.len()

    # 1) Tahun di depan -> ISO, tidak pernah ditukar
    ymd = numeric & (a_len == 4)
    if ymd.any():
        iso = parts.loc[ymd, 0] + "-" + parts.loc[ymd, 2] + "-" + parts.loc[ymd, 3]
        result.loc[ymd] = pd.to_datetime(iso, format="%Y-%m-%d", errors="coerce")
        handled |= ymd

    # 2) Hari/bulan di depan, tahun 4 digit -> simpulkan urutan per separator
    dm = numeric & (a_len <= 2) & (c_len == 4)
    for sep in parts.loc[dm, 1].dropna().unique():
        grp = dm & (parts[1] == sep)
        first = parts.loc[grp, 0].astype(int)
        second = parts.loc[grp, 2].astype(int)
        day_evidence = (first > 12).any()
        month_evidence = (second > 12).any()
        if day_evidence and not month_evidence:
            default_dayfirst = True
            has_clear_evidence = True
        elif month_evidence and not day_evidence:
            default_dayfirst = False
            has_clear_evidence = True
        else:
            default_dayfirst = True  # default Indonesia: DD/MM/YYYY
            has_clear_evidence = False

        dayfirst = np.where(first > 12, True, np.where(second > 12, False, default_dayfirst))
        day = np.where(dayfirst, first, second)
        month = np.where(dayfirst, second, first)
        year = parts.loc[grp, 3].astype(int)
        result.loc[grp] = pd.to_datetime(
            pd.DataFrame({"year": year, "month": month, "day": day}, index=first.index),
            errors="coerce",
        )
        handled |= grp

        if not has_clear_evidence:
            n_ambigu += int(((first <= 12) & (second <= 12) & (first != second)).sum())

    # 3) ISO dengan jam -> parse sebagai ISO8601 (jangan kena dayfirst)
    iso_dt = valid & ~handled & s.str.match(_ISO_DATETIME_RE, na=False)
    if iso_dt.any():
        result.loc[iso_dt] = pd.to_datetime(s[iso_dt], format="ISO8601", errors="coerce")
        handled |= iso_dt

    # 4) Sisanya (nama bulan, dll.) -> pandas mixed
    rest = valid & ~handled
    if rest.any():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            result.loc[rest] = pd.to_datetime(
                s[rest], errors="coerce", format="mixed", dayfirst=True
            )

    valid_count = int(valid.sum())
    if valid_count > 0 and (result.notna().sum() / valid_count) >= 0.5:
        return result, n_ambigu

    # Fallback lama: coba format umum satu per satu
    for fmt in COMMON_DATE_FORMATS:
        try:
            alt = pd.to_datetime(series, format=fmt, errors="coerce")
            if valid_count > 0 and (alt.notna().sum() / valid_count) >= 0.5:
                return alt, 0
        except Exception:
            continue

    return None, 0


def _convert_to_datetime(series):
    """Kompatibilitas: hanya mengembalikan Series (atau None)."""
    converted, _ = _convert_with_report(series)
    return converted
