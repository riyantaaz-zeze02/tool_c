"""
missing_handler.py — Penanganan Missing Value Fleksibel
========================================================
Strategi: median, mean, mode, constant, interpolasi, ffill, bfill, drop.
"""

import pandas as pd


def tangani_missing(df, config=None):
    """
    Mendeteksi dan mengisi nilai kosong (NaN) dalam DataFrame.

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan.
        config (dict|None): Konfigurasi strategi pengisian per tipe data.
            - numeric_strategy: 'median' | 'mean' | 'mode' | 'constant' | 'interpolate'
            - text_strategy: 'constant' | 'ffill' | 'bfill'
            - numeric_constant: nilai konstan untuk numerik (jika strategy='constant')
            - text_constant: nilai konstan untuk teks (default: 'Unknown')
            - drop_threshold: float 0.0-1.0, drop baris jika % missing > threshold

    Return:
        tuple: (pd.DataFrame yang sudah bersih, dict log perubahan)
    """
    if config is None:
        config = {}

    df = df.copy()

    numeric_strategy = config.get("numeric_strategy", "median")
    text_strategy = config.get("text_strategy", "constant")
    numeric_constant = config.get("numeric_constant", 0)
    text_constant = config.get("text_constant", "Unknown")
    drop_threshold = config.get("drop_threshold", None)

    log = {
        "missing_sebelum": {},
        "missing_sesudah": {},
        "aksi": [],
        "baris_didrop": 0,
    }

    print("🩹 Menangani missing value...")

    # Hitung missing per kolom
    missing_sebelum = df.isnull().sum()
    total_missing = missing_sebelum.sum()
    log["missing_sebelum"] = missing_sebelum.to_dict()

    if total_missing == 0:
        print("   ✅ Tidak ada missing value. Data sudah lengkap!")
        print()
        log["missing_sesudah"] = df.isnull().sum().to_dict()
        return df, log

    # Tampilkan missing sebelum
    print("   📋 Missing value per kolom (SEBELUM):")
    for kolom, jumlah in missing_sebelum.items():
        if jumlah > 0:
            print(f"      - {kolom}: {jumlah} missing")

    # Drop baris jika threshold diset
    if drop_threshold is not None and 0 < drop_threshold <= 1:
        baris_sebelum = len(df)
        jumlah_kolom = len(df.columns)
        # Hitung persentase missing per baris
        missing_per_baris = df.isnull().sum(axis=1) / jumlah_kolom
        df = df[missing_per_baris <= drop_threshold].copy()
        baris_didrop = baris_sebelum - len(df)
        if baris_didrop > 0:
            log["baris_didrop"] = baris_didrop
            log["aksi"].append(
                f"Drop {baris_didrop} baris (missing > {drop_threshold*100:.0f}%)"
            )
            print(
                f"   🗑️  {baris_didrop} baris di-drop (missing > {drop_threshold*100:.0f}%)"
            )

    # Isi missing value per kolom
    for kolom in df.columns:
        jumlah_null = df[kolom].isnull().sum()
        if jumlah_null == 0:
            continue

        if pd.api.types.is_numeric_dtype(df[kolom]):
            # Kolom numerik
            if numeric_strategy == "median":
                nilai = df[kolom].median()
                df[kolom] = df[kolom].fillna(nilai)
                label = f"median ({nilai})"
            elif numeric_strategy == "mean":
                nilai = df[kolom].mean()
                df[kolom] = df[kolom].fillna(nilai)
                label = f"mean ({nilai:.2f})"
            elif numeric_strategy == "mode":
                mode_vals = df[kolom].mode()
                nilai = mode_vals.iloc[0] if not mode_vals.empty else 0
                df[kolom] = df[kolom].fillna(nilai)
                label = f"mode ({nilai})"
            elif numeric_strategy == "constant":
                df[kolom] = df[kolom].fillna(numeric_constant)
                label = f"constant ({numeric_constant})"
            elif numeric_strategy == "interpolate":
                df[kolom] = df[kolom].interpolate(method="linear")
                df[kolom] = df[kolom].bfill().ffill()  # Handle edges
                label = "interpolasi linear"
            else:
                nilai = df[kolom].median()
                df[kolom] = df[kolom].fillna(nilai)
                label = f"median ({nilai})"

            print(f"   🔧 Kolom '{kolom}' (numerik) → diisi {label}")
            log["aksi"].append(f"'{kolom}' (numerik): {jumlah_null} missing → {label}")

        elif pd.api.types.is_datetime64_any_dtype(df[kolom]):
            # Kolom datetime
            datetime_strategy = config.get("datetime_strategy", "keep_nat")
            if datetime_strategy == "mode":
                mode_vals = df[kolom].mode()
                if not mode_vals.empty:
                    nilai = mode_vals.iloc[0]
                    df[kolom] = df[kolom].fillna(nilai)
                    label = f"mode ({nilai.strftime('%Y-%m-%d')})"
                else:
                    label = "NaT (tidak ada mode)"
            elif datetime_strategy == "ffill":
                df[kolom] = df[kolom].ffill()
                label = "forward fill"
            elif datetime_strategy == "bfill":
                df[kolom] = df[kolom].bfill()
                label = "backward fill"
            else:
                # default: biarkan NaT agar tetap valid datetime
                label = "dibiarkan NaT (valid datetime)"

            print(f"   🔧 Kolom '{kolom}' (datetime) → {label}")
            log["aksi"].append(f"'{kolom}' (datetime): {jumlah_null} missing → {label}")

        else:
            # Kolom teks/object
            if text_strategy == "constant":
                df[kolom] = df[kolom].fillna(text_constant)
                label = f"'{text_constant}'"
            elif text_strategy == "ffill":
                df[kolom] = df[kolom].ffill()
                df[kolom] = df[kolom].fillna(text_constant)  # Fallback
                label = "forward fill"
            elif text_strategy == "bfill":
                df[kolom] = df[kolom].bfill()
                df[kolom] = df[kolom].fillna(text_constant)  # Fallback
                label = "backward fill"
            else:
                df[kolom] = df[kolom].fillna(text_constant)
                label = f"'{text_constant}'"

            print(f"   🔧 Kolom '{kolom}' (teks) → diisi {label}")
            log["aksi"].append(f"'{kolom}' (teks): {jumlah_null} missing → {label}")

    # Verifikasi
    missing_sesudah = df.isnull().sum().sum()
    log["missing_sesudah"] = df.isnull().sum().to_dict()
    print(f"   ✅ Selesai! Missing value: {total_missing} → {missing_sesudah}")
    print()

    return df, log
