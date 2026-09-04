"""
reader.py — Pembaca Multi-Format (CSV, XLSX, XLS)
===================================================
Modul untuk membaca file data dari berbagai format dan mengembalikan DataFrame.
"""

import pandas as pd
import os


def baca_data(file_path, sheet_name=0):
    """
    Membaca file CSV atau Excel dan mengembalikan DataFrame.

    Parameter:
        file_path (str): Lokasi file yang mau dibaca.
        sheet_name (str|int): Nama atau indeks sheet Excel (default: sheet pertama).

    Return:
        pd.DataFrame: Data dalam bentuk tabel DataFrame.

    Raises:
        FileNotFoundError: Jika file tidak ditemukan.
        ValueError: Jika format file tidak didukung.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' tidak ditemukan!")

    ekstensi = os.path.splitext(file_path)[1].lower()

    if ekstensi == ".csv":
        df = pd.read_csv(file_path)
        print(f"📖 Berhasil membaca file CSV: {file_path}")
    elif ekstensi in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"📖 Berhasil membaca file Excel: {file_path}")
    else:
        raise ValueError(
            f"Format '{ekstensi}' tidak didukung! Gunakan .csv, .xlsx, atau .xls"
        )

    print(f"   📊 Jumlah baris  : {len(df)}")
    print(f"   📊 Jumlah kolom  : {len(df.columns)}")
    print(f"   📊 Nama kolom    : {list(df.columns)}")
    print()

    return df
