"""
DataCleaner — Tool Pembersih Data Sederhana
============================================
Tool ini membersihkan file CSV/Excel dengan 4 langkah:
1. Baca data dari file
2. Tangani missing value (nilai kosong)
3. Hapus baris duplikat
4. Standarisasi format teks

Cara pakai:
    python data_cleaner.py data/input/contoh_data.csv
"""

import pandas as pd
import sys
import os
from cleaner.rules.text_cleaner import smart_title_case, is_name_or_title_column


# ============================================================
# FUNGSI 1: Baca Data
# ============================================================
def baca_data(file_path):
    """
    Membaca file CSV atau Excel dan mengembalikan DataFrame.

    Parameter:
        file_path (str): Lokasi file yang mau dibaca

    Return:
        pd.DataFrame: Data dalam bentuk tabel DataFrame
    """
    # Cek apakah file ada
    if not os.path.exists(file_path):
        print(f"❌ Error: File '{file_path}' tidak ditemukan!")
        return None

    # Ambil ekstensi file
    ekstensi = os.path.splitext(file_path)[1].lower()

    # Baca sesuai format
    if ekstensi == ".csv":
        df = pd.read_csv(file_path)
        print(f"📖 Berhasil membaca file CSV: {file_path}")
    elif ekstensi in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path)
        print(f"📖 Berhasil membaca file Excel: {file_path}")
    else:
        print(f"❌ Error: Format '{ekstensi}' tidak didukung! Gunakan .csv atau .xlsx")
        return None

    # Tampilkan info dasar
    print(f"   📊 Jumlah baris  : {len(df)}")
    print(f"   📊 Jumlah kolom  : {len(df.columns)}")
    print(f"   📊 Nama kolom    : {list(df.columns)}")
    print()

    return df


# ============================================================
# FUNGSI 2: Tangani Missing Value
# ============================================================
def tangani_missing(df):
    """
    Mendeteksi dan mengisi nilai kosong (NaN) dalam DataFrame.

    Strategi:
    - Kolom numerik (angka)  → diisi dengan MEDIAN
    - Kolom teks (string)    → diisi dengan "Unknown"

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan

    Return:
        pd.DataFrame: DataFrame tanpa missing value
    """
    print("🩹 Menangani missing value...")

    # Hitung missing per kolom
    missing_sebelum = df.isnull().sum()
    total_missing = missing_sebelum.sum()

    if total_missing == 0:
        print("   ✅ Tidak ada missing value. Data sudah lengkap!")
        print()
        return df

    # Tampilkan missing sebelum
    print("   📋 Missing value per kolom (SEBELUM):")
    for kolom, jumlah in missing_sebelum.items():
        if jumlah > 0:
            print(f"      - {kolom}: {jumlah} missing")

    # Isi missing value
    for kolom in df.columns:
        if df[kolom].isnull().sum() > 0:
            if df[kolom].dtype in ["int64", "float64"]:
                # Kolom numerik → isi dengan median
                nilai_median = df[kolom].median()
                df[kolom] = df[kolom].fillna(nilai_median)
                print(f"   🔧 Kolom '{kolom}' (numerik) → diisi median: {nilai_median}")
            else:
                # Kolom teks → isi dengan "Unknown"
                df[kolom] = df[kolom].fillna("Unknown")
                print(f"   🔧 Kolom '{kolom}' (teks) → diisi: 'Unknown'")

    # Verifikasi
    missing_sesudah = df.isnull().sum().sum()
    print(f"   ✅ Selesai! Missing value: {total_missing} → {missing_sesudah}")
    print()

    return df


# ============================================================
# FUNGSI 3: Hapus Duplikat
# ============================================================
def hapus_duplikat(df):
    """
    Mendeteksi dan menghapus baris duplikat dari DataFrame.
    Baris pertama yang muncul akan disimpan (keep='first').

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan

    Return:
        pd.DataFrame: DataFrame tanpa baris duplikat
    """
    print("🗑️  Menghapus duplikat...")

    baris_sebelum = len(df)
    jumlah_duplikat = df.duplicated().sum()

    if jumlah_duplikat == 0:
        print("   ✅ Tidak ada duplikat. Data sudah unik!")
        print()
        return df

    # Hapus duplikat, simpan kemunculan pertama
    df = df.drop_duplicates(keep="first")
    baris_sesudah = len(df)

    print(f"   📋 Baris sebelum : {baris_sebelum}")
    print(f"   📋 Duplikat      : {jumlah_duplikat}")
    print(f"   ✅ Baris sesudah : {baris_sesudah}")
    print()

    return df


# ============================================================
# FUNGSI 4: Standarisasi Format
# ============================================================
def standarisasi_format(df):
    """
    Merapikan format data dalam DataFrame:
    - Kolom teks: hapus spasi berlebih (strip) + ubah ke Title Case
    - Kolom tanggal: konversi ke format datetime

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan

    Return:
        pd.DataFrame: DataFrame dengan format standar
    """
    print("✨ Standarisasi format...")

    kolom_diproses = []

    for kolom in df.columns:
        if df[kolom].dtype == "object" or pd.api.types.is_string_dtype(df[kolom]):
            # Trim spasi di awal dan akhir
            df[kolom] = df[kolom].astype(str).str.strip()

            # Cek apakah kolom ini kemungkinan tanggal
            if "tanggal" in kolom.lower() or "date" in kolom.lower():
                try:
                    df[kolom] = pd.to_datetime(df[kolom], errors="coerce")
                    print(f"   📅 Kolom '{kolom}' → dikonversi ke datetime")
                    kolom_diproses.append(kolom)
                    continue
                except Exception:
                    pass

            # Hanya terapkan Title Case pada kolom yang murni nama/judul (misal: artist, tour, nama, kota)
            if is_name_or_title_column(kolom):
                df[kolom] = df[kolom].apply(lambda x: smart_title_case(x) if pd.notna(x) else x)
                print(f"   🔤 Kolom '{kolom}' → strip spasi + Smart Title Case (nama/judul)")
            else:
                print(f"   🔤 Kolom '{kolom}' → strip spasi (kapitalisasi dipertahankan)")
            kolom_diproses.append(kolom)

    if not kolom_diproses:
        print("   ✅ Tidak ada kolom yang perlu distandarisasi.")
    else:
        print(f"   ✅ Selesai! {len(kolom_diproses)} kolom distandarisasi.")
    print()

    return df


# ============================================================
# FUNGSI UTAMA: Main (Orkestrator)
# ============================================================
def main():
    """
    Fungsi utama yang menjalankan seluruh pipeline cleaning.
    Memanggil 4 fungsi secara berurutan dan menyimpan hasilnya.
    """
    print("=" * 55)
    print("   🧹 DataCleaner — Tool Pembersih Data Sederhana")
    print("=" * 55)
    print()

    # Ambil path file dari argumen command line
    if len(sys.argv) < 2:
        print("❌ Cara pakai: python data_cleaner.py <path_file>")
        print("   Contoh   : python data_cleaner.py data/input/contoh_data.csv")
        return

    file_input = sys.argv[1]

    # === LANGKAH 1: Baca Data ===
    print("━" * 55)
    print("LANGKAH 1/4 — Membaca Data")
    print("━" * 55)
    df = baca_data(file_input)
    if df is None:
        return

    # === LANGKAH 2: Tangani Missing ===
    print("━" * 55)
    print("LANGKAH 2/4 — Menangani Missing Value")
    print("━" * 55)
    df = tangani_missing(df)

    # === LANGKAH 3: Hapus Duplikat ===
    print("━" * 55)
    print("LANGKAH 3/4 — Menghapus Duplikat")
    print("━" * 55)
    df = hapus_duplikat(df)

    # === LANGKAH 4: Standarisasi Format ===
    print("━" * 55)
    print("LANGKAH 4/4 — Standarisasi Format")
    print("━" * 55)
    df = standarisasi_format(df)

    # === SIMPAN HASIL ===
    print("━" * 55)
    print("HASIL AKHIR")
    print("━" * 55)

    # Buat nama file output
    nama_file = os.path.splitext(os.path.basename(file_input))[0]
    file_output = os.path.join("data", "output", f"{nama_file}_bersih.csv")

    # Pastikan folder output ada
    os.makedirs(os.path.dirname(file_output), exist_ok=True)

    # Simpan ke CSV
    df.to_csv(file_output, index=False)
    print(f"   💾 File CSV disimpan ke  : {file_output}")

    # Simpan juga ke Excel berformat profesional melalui CleaningEngine
    try:
        from cleaner.engine import CleaningEngine
        file_excel = os.path.join("data", "output", f"{nama_file}_bersih.xlsx")
        engine = CleaningEngine()
        engine.export_excel(df, file_excel, theme="corporate_blue", sheet_name="Data Bersih")
        print(f"   📊 File Excel disimpan ke: {file_excel}")
    except Exception as e:
        pass

    print(f"   📊 Total baris akhir     : {len(df)}")
    print()
    print("=" * 55)
    print("   ✅ SELESAI! Data berhasil dibersihkan.")
    print("=" * 55)


# Jalankan main() saat file ini dieksekusi langsung
if __name__ == "__main__":
    main()
