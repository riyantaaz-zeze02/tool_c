"""
themes.py — Palet Warna dan Tema Desain Excel
==============================================
Menyediakan tema warna siap pakai untuk header, zebra striping, border, dan total row.
"""

THEMES = {
    "corporate_blue": {
        "name": "Corporate Blue",
        "header_fill": "1F4E79",       # Deep Navy
        "header_font": "FFFFFF",       # Putih
        "zebra_fill": "F2F5F9",        # Biru sangat muda
        "border_color": "BDD7EE",      # Biru pastel
        "total_fill": "D9E1F2",        # Biru sedang
        "total_font": "002060",        # Biru gelap
        "font_name": "Calibri",
    },
    "modern_slate": {
        "name": "Modern Slate",
        "header_fill": "2F3640",       # Dark Slate Charcoal
        "header_font": "FFFFFF",
        "zebra_fill": "F8F9FA",        # Light Gray
        "border_color": "DCDDE1",
        "total_fill": "E1E2E6",
        "total_font": "1E272E",
        "font_name": "Segoe UI",
    },
    "emerald_green": {
        "name": "Emerald Teal",
        "header_fill": "1B4D3E",       # Deep Emerald Green
        "header_font": "FFFFFF",
        "zebra_fill": "F0F9F5",        # Hijau mint lembut
        "border_color": "A3E4D7",
        "total_fill": "D4EFDF",
        "total_font": "0E6251",
        "font_name": "Calibri",
    },
    "sunset_coral": {
        "name": "Sunset Coral",
        "header_fill": "A93226",       # Deep Coral / Burgundy
        "header_font": "FFFFFF",
        "zebra_fill": "FDF2E9",        # Peach lembut
        "border_color": "F5CBA7",
        "total_fill": "EDBB99",
        "total_font": "78281F",
        "font_name": "Segoe UI",
    },
    "classic_excel": {
        "name": "Classic Office",
        "header_fill": "217346",       # Excel Office Green
        "header_font": "FFFFFF",
        "zebra_fill": "F3F8F4",
        "border_color": "C6EFCE",
        "total_fill": "E2EFDA",
        "total_font": "006100",
        "font_name": "Calibri",
    },
}


def get_theme(theme_name="corporate_blue"):
    """
    Mengambil konfigurasi tema berdasarkan nama.
    Fallback ke 'corporate_blue' jika nama tidak ditemukan.
    """
    key = str(theme_name).lower().replace("-", "_").replace(" ", "_")
    return THEMES.get(key, THEMES["corporate_blue"])
