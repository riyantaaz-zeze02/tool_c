"""
duplicate_handler.py — Deduplikasi Tingkat Lanjut
==================================================
Mendukung: semua kolom, subset kolom kunci, keep first/last.
"""

import pandas as pd


def hapus_duplikat(df, config=None):
    """
    Mendeteksi dan menghapus baris duplikat dari DataFrame.

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan.
        config (dict|None): Konfigurasi deduplikasi.
            - subset: list kolom untuk cek duplikat (None = semua kolom)
            - keep: 'first' | 'last' | False (hapus semua duplikat)

    Return:
        tuple: (pd.DataFrame tanpa duplikat, dict log perubahan)
    """
    if config is None:
        config = {}

    df = df.copy()

    subset = config.get("subset", None)
    keep = config.get("keep", "first")

    log = {
        "baris_sebelum": len(df),
        "jumlah_duplikat": 0,
        "baris_sesudah": len(df),
        "subset_kolom": subset,
        "keep": keep,
    }

    print("🗑️  Menghapus duplikat...")

    baris_sebelum = len(df)
    jumlah_duplikat = df.duplicated(subset=subset, keep=keep).sum()

    if jumlah_duplikat == 0:
        print("   ✅ Tidak ada duplikat. Data sudah unik!")
        print()
        return df, log

    # Hapus duplikat
    df = df.drop_duplicates(subset=subset, keep=keep).reset_index(drop=True)
    baris_sesudah = len(df)

    log["jumlah_duplikat"] = jumlah_duplikat
    log["baris_sesudah"] = baris_sesudah

    subset_info = f" (kolom: {subset})" if subset else " (semua kolom)"
    print(f"   📋 Baris sebelum : {baris_sebelum}")
    print(f"   📋 Duplikat      : {jumlah_duplikat}{subset_info}")
    print(f"   📋 Keep          : {keep}")
    print(f"   ✅ Baris sesudah : {baris_sesudah}")
    print()

    return df, log
