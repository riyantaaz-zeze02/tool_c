"""
reader.py — Pembaca Multi-Format (CSV, XLSX, XLS)
===================================================
Modul untuk membaca file data dari berbagai format dan mengembalikan DataFrame.
"""

import pandas as pd
import os


def read_csv_with_fallback(file_path_or_buffer, **kwargs):
    """
    Membaca file CSV dengan mencoba beberapa encoding berurutan:
    utf-8 -> utf-8-sig -> windows-1252 -> latin-1.

    Parameter:
        file_path_or_buffer (str | IO): Path file atau buffer/file-like object.
        **kwargs: Parameter tambahan untuk pd.read_csv.

    Return:
        pd.DataFrame: Data dalam bentuk tabel DataFrame.
    """
    encodings = ["utf-8", "utf-8-sig", "windows-1252", "latin-1"]
    last_err = None

    for enc in encodings:
        try:
            if hasattr(file_path_or_buffer, "seek"):
                file_path_or_buffer.seek(0)
            return pd.read_csv(file_path_or_buffer, encoding=enc, **kwargs)
        except UnicodeDecodeError as e:
            last_err = e
            continue

    if last_err is not None:
        raise last_err
    return pd.read_csv(file_path_or_buffer, **kwargs)


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
        df = read_csv_with_fallback(file_path)
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

