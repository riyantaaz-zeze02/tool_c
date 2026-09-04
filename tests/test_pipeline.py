"""
test_pipeline.py — Integration Tests untuk Pipeline Cleaning & Export Excel
"""

import os
import openpyxl
import pandas as pd
from cleaner.engine import CleaningEngine


def test_full_pipeline_on_sample_csv(tmp_path):
    sample_csv = os.path.join("data", "input", "contoh_data.csv")
    assert os.path.exists(sample_csv)

    engine = CleaningEngine()
    df_clean, report = engine.clean(sample_csv)

    # Validasi hasil cleaning
    assert len(df_clean) == 11
    assert report["baris_awal"] == 15
    assert report["total_duplikat_dihapus"] == 4
    assert df_clean["nama"].iloc[0] == "Hani Wijaya"

    # Validasi ekspor Excel
    out_file = str(tmp_path / "output_test.xlsx")
    engine.export_excel(df_clean, out_file, report=report, theme="emerald_green")
    assert os.path.exists(out_file)

    # Validasi isi file Excel
    wb = openpyxl.load_workbook(out_file)
    assert "Data Bersih" in wb.sheetnames
    assert "Ringkasan Cleaning" in wb.sheetnames

    ws_data = wb["Data Bersih"]
    assert ws_data["A1"].value == "nama"
    assert ws_data["E13"].value == "=SUM(E2:E12)"
