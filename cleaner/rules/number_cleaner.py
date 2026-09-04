"""
number_cleaner.py — Pembersihan Angka & Mata Uang
===================================================
Konversi string uang/angka ke numerik: "Rp 1.500.000,00" → 1500000.0
"""

import pandas as pd
import re


def bersihkan_angka(df, config=None):
    """
    Membersihkan kolom yang berisi angka/mata uang dalam format string.

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan.
        config (dict|None): Konfigurasi.
            - kolom_angka: list kolom yang harus dikonversi ke numerik
            - kolom_currency: list kolom mata uang (Rp, $, €, dll)
            - kolom_persen: list kolom persentase ("15%" → 0.15)
            - auto_detect: bool (True = deteksi otomatis kolom angka dari string)

    Return:
        tuple: (pd.DataFrame bersih, dict log perubahan)
    """
    if config is None:
        config = {}

    df = df.copy()

    kolom_angka = config.get("kolom_angka", [])
    kolom_currency = config.get("kolom_currency", [])
    kolom_persen = config.get("kolom_persen", [])
    auto_detect = config.get("auto_detect", True)

    log = {"kolom_diproses": [], "aksi": []}

    print("🔢 Membersihkan format angka...")

    if auto_detect:
        for kolom in df.columns:
            if kolom in kolom_angka or kolom in kolom_currency or kolom in kolom_persen:
                continue

            if df[kolom].dtype == "object" or pd.api.types.is_string_dtype(df[kolom]):
                sample = df[kolom].dropna().head(20).astype(str)
                if len(sample) == 0:
                    continue

                # Cek apakah mengandung simbol mata uang
                currency_pattern = r"^[\s]*(Rp\.?|USD|\$|€|£|¥)\s*[\d.,]+$"
                if sample.str.match(currency_pattern, na=False).mean() > 0.5:
                    kolom_currency.append(kolom)
                    continue

                # Cek apakah mengandung simbol persen
                persen_pattern = r"^[\d.,\s]+%$"
                if sample.str.match(persen_pattern, na=False).mean() > 0.5:
                    kolom_persen.append(kolom)
                    continue

                # Cek apakah string angka dengan separator
                angka_pattern = r"^[\d.,\s-]+$"
                if sample.str.match(angka_pattern, na=False).mean() > 0.7:
                    # Pastikan bukan nomor telepon / ID
                    nama_lower = kolom.lower()
                    skip_keywords = [
                        "telp", "telepon", "phone", "hp", "id",
                        "nik", "nip", "kode", "code", "no_", "nomor",
                    ]
                    if not any(kw in nama_lower for kw in skip_keywords):
                        kolom_angka.append(kolom)

    # Proses kolom currency
    for kolom in kolom_currency:
        if kolom not in df.columns:
            continue
        df[kolom] = df[kolom].astype(str).apply(_parse_currency)
        log["kolom_diproses"].append(kolom)
        log["aksi"].append(f"'{kolom}': currency → float")
        print(f"   💰 Kolom '{kolom}' → konversi currency ke float")

    # Proses kolom persen
    for kolom in kolom_persen:
        if kolom not in df.columns:
            continue
        df[kolom] = df[kolom].astype(str).apply(_parse_persen)
        log["kolom_diproses"].append(kolom)
        log["aksi"].append(f"'{kolom}': persen → float (desimal)")
        print(f"   📊 Kolom '{kolom}' → konversi persen ke desimal")

    # Proses kolom angka biasa
    for kolom in kolom_angka:
        if kolom not in df.columns:
            continue
        df[kolom] = df[kolom].astype(str).apply(_parse_angka)
        log["kolom_diproses"].append(kolom)
        log["aksi"].append(f"'{kolom}': string angka → float")
        print(f"   🔢 Kolom '{kolom}' → konversi string angka ke float")

    if not log["kolom_diproses"]:
        print("   ✅ Tidak ada kolom angka yang perlu dibersihkan.")
    else:
        print(
            f"   ✅ Selesai! {len(log['kolom_diproses'])} kolom angka diproses."
        )
    print()

    return df, log


def _parse_currency(value):
    """
    Parse string mata uang ke float.
    "Rp 1.500.000,00" → 1500000.0
    "$1,500.50" → 1500.50
    """
    if pd.isna(value) or value in ("nan", "None", ""):
        return None

    value = str(value).strip()

    # Hapus simbol mata uang
    value = re.sub(r"(Rp\.?\s*|USD\s*|\$|€|£|¥)", "", value).strip()

    # Deteksi format Indonesia (titik = ribuan, koma = desimal)
    # vs format US/Intl (koma = ribuan, titik = desimal)
    if re.match(r"^[\d]+\.[\d]{3}", value):
        # Format Indonesia: 1.500.000,00
        value = value.replace(".", "")
        value = value.replace(",", ".")
    else:
        # Format US: 1,500.00
        value = value.replace(",", "")

    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _parse_persen(value):
    """Parse string persentase ke float desimal. "15%" → 0.15"""
    if pd.isna(value) or value in ("nan", "None", ""):
        return None

    value = str(value).strip().replace("%", "").replace(",", ".").strip()
    try:
        return float(value) / 100
    except (ValueError, TypeError):
        return None


def _parse_angka(value):
    """Parse string angka dengan separator ke float."""
    if pd.isna(value) or value in ("nan", "None", ""):
        return None

    value = str(value).strip()

    # Deteksi format Indonesia vs US
    if re.match(r"^[\d]+\.[\d]{3}", value):
        value = value.replace(".", "")
        value = value.replace(",", ".")
    else:
        value = value.replace(",", "")

    try:
        return float(value)
    except (ValueError, TypeError):
        return None
