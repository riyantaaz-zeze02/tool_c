import pandas as pd

from cleaner.engine import CleaningEngine


def _write_csv(path, df):
    df.to_csv(path, index=False)


def test_append_keeps_base_untouched_and_adds_only_new_rows(tmp_path):
    base = pd.DataFrame({
        "Order ID": ["A1", "A2"],
        "Nama": ["Budi", "Siti"],
        "Harga": ["Rp 100.000", "Rp 200.000"],
        "Tanggal": ["15/01/2025", "20/01/2025"],
    })
    new = pd.DataFrame({
        "Order ID": ["A2", "A3"],
        "Nama": ["Siti", "Andi"],
        "Harga": ["Rp 200.000", "Rp 300.000"],
        "Tanggal": ["20/01/2025", "05/02/2025"],
    })
    base_path = tmp_path / "base.csv"
    new_path = tmp_path / "new.csv"
    _write_csv(base_path, base)
    _write_csv(new_path, new)

    df_final, report = CleaningEngine().append_files(str(base_path), [str(new_path)])

    assert len(df_final) == 3
    assert report["total_ditambahkan"] == 1
    assert report["total_dilewati"] == 1
    assert list(df_final["Order ID"]) == ["A1", "A2", "A3"]


def test_append_reorders_columns_and_skips_total_row(tmp_path):
    base = pd.DataFrame({
        "Order ID": ["A1", "A2"],
        "Nama": ["Budi", "Siti"],
        "Harga": ["Rp 100.000", "Rp 200.000"],
        "Tanggal": ["15/01/2025", "20/01/2025"],
    })
    base.loc[len(base)] = ["TOTAL", "", "", ""]
    new = pd.DataFrame({
        "Nama": ["Andi"],
        "Tanggal": ["05/02/2025"],
        "Harga": ["Rp 300.000"],
        "Order ID": ["A3"],
    })
    base_path = tmp_path / "base.csv"
    new_path = tmp_path / "new.csv"
    _write_csv(base_path, base)
    _write_csv(new_path, new)

    df_final, report = CleaningEngine().append_files(str(base_path), [str(new_path)])

    assert df_final["Order ID"].tolist() == ["A1", "A2", "A3"]
    assert report["base_total_row_dibuang"] == 1
    assert report["total_ditambahkan"] == 1
    assert report["baris_akhir"] == 3


def test_append_rejects_different_column_set(tmp_path):
    base = pd.DataFrame({"A": [1], "B": [2]})
    new = pd.DataFrame({"A": [3], "C": [4]})
    base_path = tmp_path / "base.csv"
    new_path = tmp_path / "new.csv"
    _write_csv(base_path, base)
    _write_csv(new_path, new)

    try:
        CleaningEngine().append_files(str(base_path), [str(new_path)])
        assert False, "Expected ValueError"
    except ValueError as exc:
        assert "Kolom hilang" in str(exc) or "kolom" in str(exc).lower()


def test_append_uses_dedupe_keys_to_skip_rows(tmp_path):
    base = pd.DataFrame({
        "Order ID": ["A1"],
        "Nama": ["Budi"],
        "Harga": ["Rp 100.000"],
    })
    new = pd.DataFrame({
        "Order ID": ["A1", "A3", "A3", "A2"],
        "Nama": ["Budi", "Cici", "Cici", "Ayu"],
        "Harga": ["Rp 100.000", "Rp 999.000", "Rp 900.000", "Rp 150.000"],
    })
    base_path = tmp_path / "base.csv"
    new_path = tmp_path / "new.csv"
    _write_csv(base_path, base)
    _write_csv(new_path, new)

    df_final, report = CleaningEngine().append_files(str(base_path), [str(new_path)], dedupe_keys=["Order ID"])

    assert list(df_final["Order ID"]) == ["A1", "A3", "A2"]
    assert report["total_dilewati"] == 2
    skipped = report["skipped_rows"]
    assert set(skipped["Alasan"].tolist()) == {"sudah ada di file dasar", "duplikat di file baru"}
