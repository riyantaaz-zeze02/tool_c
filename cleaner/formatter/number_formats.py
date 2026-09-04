"""
number_formats.py — Format Tampilan Angka, Mata Uang, & Tanggal di Excel
========================================================================
Menyediakan format mask native Excel (misal: "Rp #,##0", "yyyy-mm-dd", "#,##0.00").
"""

FORMAT_MASKS = {
    # Mata uang
    "currency_idr": '"Rp" #,##0',
    "currency_idr_dec": '"Rp" #,##0.00',
    "currency_usd": '"$"#,##0.00',
    "currency_eur": '€#,##0.00',
    # Angka
    "integer": '#,##0',
    "decimal": '#,##0.00',
    "decimal_3": '#,##0.000',
    # Persentase
    "percent": '0.0%',
    "percent_2": '0.00%',
    # Tanggal & Waktu
    "date_iso": 'yyyy-mm-dd',
    "date_id": 'dd/mm/yyyy',
    "date_medium": 'dd-mmm-yyyy',
    "datetime": 'yyyy-mm-dd hh:mm:ss',
    # Teks biasa
    "text": '@',
}


def get_column_format_mask(column_name, series_dtype, sample_val=None, default_currency="IDR"):
    """
    Menentukan format mask Excel terbaik secara otomatis berdasarkan nama kolom dan tipe datanya.
    """
    col_lower = column_name.lower()

    # Cek mata uang / finansial
    currency_keywords = ["gaji", "salary", "harga", "price", "biaya", "cost", "total", "subtotal", "nominal", "amount", "revenue", "untung", "laba", "rugi"]
    if any(kw in col_lower for kw in currency_keywords):
        if str(default_currency).upper() == "USD":
            return FORMAT_MASKS["currency_usd"]
        return FORMAT_MASKS["currency_idr"]

    # Cek persentase
    percent_keywords = ["persen", "percent", "pct", "diskon", "discount", "pajak_pct", "margin"]
    if any(kw in col_lower for kw in percent_keywords):
        return FORMAT_MASKS["percent"]

    # Cek tanggal
    date_keywords = ["tanggal", "date", "tgl", "lahir", "birth", "created", "expired"]
    if any(kw in col_lower for kw in date_keywords) or str(series_dtype).startswith("datetime"):
        return FORMAT_MASKS["date_iso"]

    # Cek numerik bulat (umur, kuantitas, stok)
    int_keywords = ["umur", "age", "qty", "quantity", "jumlah", "stok", "stock", "count", "tahun", "year"]
    if any(kw in col_lower for kw in int_keywords):
        return FORMAT_MASKS["integer"]

    # Tipe numerik umum
    if "int" in str(series_dtype).lower():
        return FORMAT_MASKS["integer"]
    elif "float" in str(series_dtype).lower():
        return FORMAT_MASKS["decimal"]

    return None
