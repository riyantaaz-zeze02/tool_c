"""Regresi: harga berprefix 'Rp' dan tanggal format campuran tidak boleh rusak diam-diam."""
import random
from datetime import datetime, timedelta

import pandas as pd

from cleaner.rules.date_cleaner import bersihkan_tanggal
from cleaner.rules.number_cleaner import bersihkan_angka, _parse_angka


def _price_styles(v):
    return [
        str(v),
        f"Rp {v:,}".replace(",", "."),
        f"{v:,}",
        f"{v:,}".replace(",", "."),
        f"{v:,}[a]".replace(",", "."),
        f" {v} ",
    ]


def test_rp_prefix_in_mixed_column_is_not_lost():
    random.seed(1)
    truth, raw = [], []
    for _ in range(600):
        v = random.randint(10_000, 320_000)
        truth.append(float(v))
        raw.append(random.choice(_price_styles(v)))
    df = pd.DataFrame({"Harga Satuan": raw})
    out, _ = bersihkan_angka(df, {"auto_detect": True})
    assert out["Harga Satuan"].tolist() == truth


def test_parse_angka_strips_currency_symbol():
    assert _parse_angka("Rp 75.579") == 75579.0
    assert _parse_angka("Rp. 1.500.000,50") == 1500000.5


def _date_formats():
    # dash+month-first, slash+day-first, year-first (dot & dash), month-name
    return ["%d/%m/%Y", "%m-%d-%Y", "%Y.%m.%d", "%d %b %Y", "%Y-%m-%d"]


def test_mixed_date_formats_are_not_swapped():
    random.seed(2)
    start = datetime(2025, 1, 1)
    truth, raw = [], []
    for _ in range(1500):
        dt = start + timedelta(days=random.randint(0, 300))
        truth.append(pd.Timestamp(dt))
        raw.append(dt.strftime(random.choice(_date_formats())))
    df = pd.DataFrame({"Tanggal Order": raw})
    out, log = bersihkan_tanggal(df, {"auto_detect": True})
    wrong = (out["Tanggal Order"].reset_index(drop=True) != pd.Series(truth)).sum()
    assert wrong == 0


def test_iso_dates_never_swapped_even_when_ambiguous():
    df = pd.DataFrame({"tanggal": ["2025-07-04", "2025.07.10", "2025/03/05"]})
    out, _ = bersihkan_tanggal(df, {"auto_detect": True})
    assert out["tanggal"].tolist() == [
        pd.Timestamp("2025-07-04"), pd.Timestamp("2025-07-10"), pd.Timestamp("2025-03-05"),
    ]


def test_truly_ambiguous_dates_are_flagged_not_silent():
    df = pd.DataFrame({"tanggal": ["03/04/2025", "05/06/2025", "01/02/2025"]})
    out, log = bersihkan_tanggal(df, {"auto_detect": True})
    assert log.get("peringatan"), "tanggal ambigu harus diberi peringatan"


def test_evidence_resolves_convention():
    # ada 25/03 -> day-first terbukti, jadi 03/04 = 3 April
    df = pd.DataFrame({"tanggal": ["25/03/2025", "03/04/2025"]})
    out, log = bersihkan_tanggal(df, {"auto_detect": True})
    assert out["tanggal"].iloc[1] == pd.Timestamp("2025-04-03")
    assert not log.get("peringatan")
