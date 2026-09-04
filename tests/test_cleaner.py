"""
test_cleaner.py — Unit Tests untuk Rules Pembersihan Data
"""

import pandas as pd
import numpy as np
import pytest

from cleaner.rules.missing_handler import tangani_missing
from cleaner.rules.duplicate_handler import hapus_duplikat
from cleaner.rules.text_cleaner import bersihkan_teks
from cleaner.rules.number_cleaner import bersihkan_angka, _parse_currency, _parse_angka
from cleaner.rules.date_cleaner import bersihkan_tanggal


def test_missing_handler_numeric():
    df = pd.DataFrame({
        "nilai": [10.0, 20.0, np.nan, 40.0],
        "kategori": ["A", "B", np.nan, "D"],
    })
    cleaned, log = tangani_missing(df, {"numeric_strategy": "median", "text_constant": "Unknown"})
    # median dari 10, 20, 40 adalah 20.0
    assert cleaned["nilai"].iloc[2] == 20.0
    assert cleaned["kategori"].iloc[2] == "Unknown"
    assert cleaned.isnull().sum().sum() == 0


def test_duplicate_handler():
    df = pd.DataFrame({
        "id": [1, 2, 2, 3],
        "nama": ["Andi", "Budi", "Budi", "Cici"],
    })
    cleaned, log = hapus_duplikat(df, {"keep": "first"})
    assert len(cleaned) == 3
    assert log["jumlah_duplikat"] == 1


def test_text_cleaner_nan_preservation():
    df = pd.DataFrame({
        "kota": ["  jakarta  ", "BANDUNG", np.nan],
    })
    cleaned, log = bersihkan_teks(df, {"case_format": "title"})
    assert cleaned["kota"].iloc[0] == "Jakarta"
    assert cleaned["kota"].iloc[1] == "Bandung"
    # Pastikan NaN tidak berubah menjadi string 'Nan'
    assert pd.isna(cleaned["kota"].iloc[2])


def test_number_cleaner_currency():
    df = pd.DataFrame({
        "gaji": ["Rp 5.000.000", "Rp 7.500.000,50", "1.200.000"],
        "diskon": ["10%", "25%", "5%"],
    })
    cleaned, log = bersihkan_angka(df, {"auto_detect": True})
    assert cleaned["gaji"].iloc[0] == 5000000.0
    assert cleaned["gaji"].iloc[1] == 7500000.50
    assert cleaned["diskon"].iloc[0] == 0.10


def test_date_cleaner():
    df = pd.DataFrame({
        "tanggal_lahir": ["2024-01-15", "15/02/2024", "2024-03-10"],
    })
    cleaned, log = bersihkan_tanggal(df, {"auto_detect": True})
    assert pd.api.types.is_datetime64_any_dtype(cleaned["tanggal_lahir"])
    assert len(cleaned["tanggal_lahir"].dropna()) == 3


def test_smart_title_case_brackets_and_acronyms():
    from cleaner.rules.text_cleaner import smart_title_case

    # Kurung siku dipertahankan apa adanya
    assert smart_title_case("The Eras Tour [LIVE]") == "The Eras Tour [LIVE]"
    assert smart_title_case("Coldplay [OFFICIAL VIP]") == "Coldplay [OFFICIAL VIP]"
    assert smart_title_case("taylor swift [acoustic version]") == "Taylor Swift [acoustic version]"

    # Akronim all-caps tetap kapital
    assert smart_title_case("tiket VIP konser di JKT") == "Tiket VIP Konser Di JKT"
    assert smart_title_case("world tour USA 2024") == "World Tour USA 2024"
    assert smart_title_case("PT TELKOM INDONESIA TBK") == "PT Telkom Indonesia TBK"

    # String murni all-caps diubah Title Case kecuali akronim terdaftar
    assert smart_title_case("BANDUNG") == "Bandung"
    assert smart_title_case("COLDPLAY MUSIC OF THE SPHERES [VIP]") == "Coldplay Music Of The Spheres [VIP]"


def test_title_case_only_name_columns():
    df = pd.DataFrame({
        "Artist": ["taylor swift", "COLDPLAY"],
        "Tour title": ["the eras tour [LIVE]", "MUSIC OF THE SPHERES [VIP]"],
        "Description": [
            "tiket presale sold out, silakan cek website resmi",
            "gate dibuka pukul 17:00 WIB untuk pemegang tiket VIP",
        ],
        "catatan_tambahan": [
            "dilarang membawa kamera profesional",
            "harap membawa kartu identitas asli",
        ],
    })

    cleaned, log = bersihkan_teks(df, {"case_format": "title", "only_name_columns": True})

    # Kolom nama (Artist, Tour title) mendapatkan Title Case cerdas
    assert cleaned["Artist"].iloc[0] == "Taylor Swift"
    assert cleaned["Artist"].iloc[1] == "Coldplay"
    assert cleaned["Tour title"].iloc[0] == "The Eras Tour [LIVE]"
    assert cleaned["Tour title"].iloc[1] == "Music Of The Spheres [VIP]"

    # Kolom non-nama (Description, catatan_tambahan) TIDAK disamaratakan ke Title Case
    assert cleaned["Description"].iloc[0] == "tiket presale sold out, silakan cek website resmi"
    assert cleaned["Description"].iloc[1] == "gate dibuka pukul 17:00 WIB untuk pemegang tiket VIP"
    assert cleaned["catatan_tambahan"].iloc[0] == "dilarang membawa kamera profesional"


def test_kolom_case_explicit_override():
    df = pd.DataFrame({
        "Artist": ["taylor swift", "coldplay"],
        "Tour title": ["the eras tour", "music of the spheres"],
    })

    # Hanya kolom 'Artist' yang diformat
    cleaned, log = bersihkan_teks(df, {"case_format": "title", "kolom_case": ["Artist"]})
    assert cleaned["Artist"].iloc[0] == "Taylor Swift"
    assert cleaned["Tour title"].iloc[0] == "the eras tour"


def test_parse_currency_with_footnotes():
    assert _parse_currency("$229,100,000[b]") == 229100000.0
    assert _parse_currency("$167,700,000[e]") == 167700000.0
    assert _parse_currency("$500,000[a][b]") == 500000.0
    assert _parse_currency("$1,000,000[17]") == 1000000.0
    assert _parse_currency("Rp 2.500.000,00[1]") == 2500000.0


def test_parse_angka_with_footnotes():
    assert _parse_angka("100,000[b]") == 100000.0
    assert _parse_angka("1,000,000[17]") == 1000000.0
    assert _parse_angka("229,100,000[a][b]") == 229100000.0
    assert _parse_angka("1.500.000[3]") == 1500000.0


def test_number_cleaner_currency_dataframe_with_footnotes():
    df = pd.DataFrame({
        "Gross": ["$229,100,000[b]", "$167,700,000[e]", "$500,000[a][b]"]
    })
    cleaned, log = bersihkan_angka(df, {"kolom_currency": ["Gross"]})
    assert cleaned["Gross"].iloc[0] == 229100000.0
    assert cleaned["Gross"].iloc[1] == 167700000.0
    assert cleaned["Gross"].iloc[2] == 500000.0
    assert not cleaned["Gross"].isnull().any()


