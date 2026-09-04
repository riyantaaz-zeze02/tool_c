"""
text_cleaner.py — Standarisasi & Normalisasi Teks
===================================================
Trim spasi, normalisasi Unicode, format kapitalisasi cerdas (Smart Title Case),
pengecualian kata dalam kurung siku & akronim all-caps, serta pembersihan karakter ilegal.
"""

import re
import unicodedata
import pandas as pd

# Daftar akronim umum (Internasional, Indonesia, Musik, Bisnis, Angka Romawi)
DEFAULT_ACRONYMS = {
    # Geografi & Organisasi
    "USA", "UK", "UAE", "EU", "UN", "WHO", "UNESCO", "NYC", "LA", "DC",
    # Wilayah & Identitas Indonesia
    "JKT", "BDG", "SBY", "YGY", "JOGJA", "IDR", "USD", "EUR", "GBP", "SGD", "AUD",
    "WIB", "WITA", "WIT", "PT", "CV", "TBK", "BUMN", "BUMD", "DPR", "MPR", "KPK",
    "KTP", "SIM", "KK", "NPWP", "PPN", "PPH", "BPJS", "BMKG", "RSUD", "RS",
    # Musik, Tur, & Entertainment
    "VIP", "VVIP", "DJ", "MC", "OST", "MV", "EP", "LP", "CD", "DVD", "HQ", "HD",
    "BTS", "NCT", "EXO", "TXT", "SEVENTEEN", "BLACKPINK", "AKB48", "JKT48",
    "EDM", "R&B", "TV", "FM", "AM",
    # Teknologi, Bisnis, & Umum
    "CEO", "CFO", "CTO", "COO", "CMO", "HR", "IT", "PR", "AI", "ML", "API",
    "URL", "URI", "SQL", "HTML", "CSS", "JS", "FAQ", "SKU", "ID", "PO", "SPK",
    # Angka Romawi
    "II", "III", "IV", "VI", "VII", "VIII", "IX", "XI", "XII", "XIII", "XIV", "XV",
    "XVI", "XVII", "XVIII", "XIX", "XX",
}


def is_name_or_title_column(col_name):
    """
    Mendeteksi apakah nama kolom merupakan kolom nama, judul, artis, tur, atau entitas
    yang pantas diformat Title Case (bukan deskripsi, catatan, atau ID).
    """
    if not isinstance(col_name, str):
        return False

    col_lower = col_name.strip().lower()

    # Kata kunci yang menandakan kolom BUKAN nama murni
    EXCLUDE_KEYWORDS = [
        "deskripsi", "description", "desc", "keterangan", "notes", "catatan",
        "detail", "komentar", "comment", "review", "ulasan", "text", "teks",
        "lirik", "lyrics", "alamat", "address", "url", "uri", "link", "email",
        "e-mail", "kode", "code", "id", "uuid", "sku", "slug", "status",
        "query", "search", "log", "message", "pesan", "tanggal", "date",
        "waktu", "time", "harga", "price", "gaji", "salary", "nominal", "total"
    ]
    for exc in EXCLUDE_KEYWORDS:
        if exc in col_lower:
            return False

    # Kata kunci kolom yang merupakan nama / entitas / judul / lokasi
    NAME_KEYWORDS = [
        "nama", "name", "artist", "artis", "musician", "singer", "penyanyi", "band",
        "tour", "tour_title", "tour_name", "konser", "concert", "title", "judul",
        "song", "lagu", "album", "track", "kota", "city", "negara", "country",
        "provinsi", "province", "venue", "lokasi", "location", "tempat",
        "customer", "pelanggan", "client", "klien", "author", "penulis", "penerbit",
        "pembeli", "buyer", "seller", "penjual", "user", "pengguna", "member", "anggota"
    ]
    for kw in NAME_KEYWORDS:
        if kw in col_lower:
            return True

    return False


def smart_title_case(text, config=None):
    """
    Mengubah teks menjadi Title Case dengan aturan cerdas:
    1. Melindungi teks dalam kurung siku `[...]` (misal: [LIVE], [VIP], [OFFICIAL]).
    2. Mempertahankan akronim all-caps (misal: VIP, USA, JKT, IDR, BTS).
    3. Merapikan kapitalisasi dengan benar saat ada tanda apostrof (Don't, O'Connor).
    """
    if not isinstance(text, str) or not text:
        return text

    if config is None:
        config = {}

    preserve_brackets = config.get("preserve_brackets", True)
    preserve_acronyms = config.get("preserve_acronyms", True)
    custom_acronyms = set(config.get("custom_acronyms", []))
    all_acronyms = DEFAULT_ACRONYMS | {a.upper() for a in custom_acronyms}

    if preserve_brackets:
        # Pisahkan teks berdasarkan kurung siku: [ ... ]
        chunks = re.split(r"(\[[^\]]*\])", text)
    else:
        chunks = [text]

    # Evaluasi apakah bagian non-kurung-siku seluruhnya huruf besar (all-caps)
    # (contoh: "BUDI SANTOSO" atau "COLDPLAY LIVE IN JAKARTA")
    non_bracket_text = "".join(chunks[0::2])
    alpha_chars = [c for c in non_bracket_text if c.isalpha()]
    is_all_uppercase = len(alpha_chars) > 0 and all(c.isupper() for c in alpha_chars)

    def process_word(word):
        if not word:
            return word

        # Ekstrak awalan non-alfanumerik, inti kata, dan akhiran (misal: "('VIP')" -> "('", "VIP", "')")
        match = re.match(r"^([^a-zA-Z0-9]*)(.*?)([^a-zA-Z0-9]*)$", word)
        if not match:
            return word
        prefix, core, suffix = match.groups()
        if not core:
            return word

        upper_core = core.upper()

        if preserve_acronyms:
            # 1) Apakah termasuk akronim yang dikenal?
            if upper_core in all_acronyms:
                return f"{prefix}{upper_core}{suffix}"

            # 2) Jika teks aslinya mixed-case (bukan all-caps) dan kata ini all-caps (>=2 karakter):
            # Berarti kata ini sengaja ditulis all-caps sebagai akronim/singkatan
            if not is_all_uppercase and core.isupper() and len(core) >= 2 and any(c.isalpha() for c in core):
                return f"{prefix}{core}{suffix}"

        # Format Title Case untuk inti kata
        if "'" in core:
            parts = core.split("'")
            titled_parts = [parts[0].capitalize()]
            for p in parts[1:]:
                # Kontraksi bahasa Inggris umum tetap lowercase
                if p.lower() in {"s", "t", "d", "m", "ll", "re", "ve"}:
                    titled_parts.append(p.lower())
                else:
                    titled_parts.append(p.capitalize())
            titled_core = "'".join(titled_parts)
        elif "-" in core:
            titled_core = "-".join(p.capitalize() for p in core.split("-"))
        else:
            titled_core = core.capitalize()

        return f"{prefix}{titled_core}{suffix}"

    def process_chunk(chunk):
        tokens = re.split(r"(\s+)", chunk)
        return "".join(process_word(t) if not t.isspace() else t for t in tokens)

    result_chunks = []
    for i, chunk in enumerate(chunks):
        if preserve_brackets and i % 2 == 1:
            # Bagian dalam kurung siku [...] -> pertahankan apa adanya
            result_chunks.append(chunk)
        else:
            result_chunks.append(process_chunk(chunk))

    return "".join(result_chunks)


def bersihkan_teks(df, config=None):
    """
    Merapikan format data teks dalam DataFrame.

    Parameter:
        df (pd.DataFrame): DataFrame yang mau dibersihkan.
        config (dict|None): Konfigurasi pembersihan teks.
            - case_format: 'title' | 'upper' | 'lower' | 'sentence' | None (tanpa ubah)
            - only_name_columns: bool (default: True) — jika True, Title Case hanya diterapkan ke kolom nama/judul
            - kolom_case: list kolom spesifik untuk kapitalisasi (override deteksi otomatis)
            - preserve_brackets: bool (default: True) — pertahankan teks dalam kurung siku [...]
            - preserve_acronyms: bool (default: True) — pertahankan akronim all-caps (VIP, USA, dll.)
            - custom_acronyms: list akronim tambahan
            - strip_whitespace: bool (default: True)
            - normalize_unicode: bool (default: True)
            - remove_non_ascii: bool (default: False)
            - kolom_exclude: list kolom yang tidak diproses sama sekali (misal: email)

    Return:
        tuple: (pd.DataFrame bersih, dict log perubahan)
    """
    if config is None:
        config = {}

    df = df.copy()
    case_format = config.get("case_format", "title")
    only_name_columns = config.get("only_name_columns", True)
    kolom_case = config.get("kolom_case", [])
    strip_whitespace = config.get("strip_whitespace", True)
    normalize_unicode = config.get("normalize_unicode", True)
    remove_non_ascii = config.get("remove_non_ascii", False)
    kolom_exclude = config.get("kolom_exclude", [])

    log = {"kolom_diproses": [], "aksi": []}

    print("✨ Standarisasi format teks...")

    for kolom in df.columns:
        if kolom in kolom_exclude:
            continue

        if not (df[kolom].dtype == "object" or pd.api.types.is_string_dtype(df[kolom])):
            continue

        mask = df[kolom].notna()
        if not mask.any():
            continue

        aksi_kolom = []
        series = df.loc[mask, kolom].astype(str)

        # Normalisasi Unicode (hapus \u00a0 dan karakter aneh)
        if normalize_unicode:
            series = series.apply(
                lambda x: unicodedata.normalize("NFKC", x) if isinstance(x, str) else x
            )
            aksi_kolom.append("normalize unicode")

        # Strip spasi berlebih
        if strip_whitespace:
            series = series.str.strip().str.replace(r"\s+", " ", regex=True)
            aksi_kolom.append("strip spasi")

        # Hapus karakter non-ASCII
        if remove_non_ascii:
            series = series.apply(
                lambda x: re.sub(r"[^\x00-\x7F]+", "", x) if isinstance(x, str) else x
            )
            aksi_kolom.append("hapus non-ASCII")

        # Cek apakah kolom ini harus diformat kapitalisasinya
        should_apply_case = True
        if kolom_case:
            should_apply_case = (kolom in kolom_case)
        elif case_format == "title" and only_name_columns:
            should_apply_case = is_name_or_title_column(kolom)

        # Format kapitalisasi
        if should_apply_case:
            if case_format == "title":
                series = series.apply(lambda x: smart_title_case(x, config=config))
                aksi_kolom.append("Smart Title Case")
            elif case_format == "upper":
                series = series.str.upper()
                aksi_kolom.append("UPPERCASE")
            elif case_format == "lower":
                series = series.str.lower()
                aksi_kolom.append("lowercase")
            elif case_format == "sentence":
                series = series.str.capitalize()
                aksi_kolom.append("Sentence case")

        df.loc[mask, kolom] = series

        if aksi_kolom:
            log["kolom_diproses"].append(kolom)
            log["aksi"].append(f"'{kolom}': {', '.join(aksi_kolom)}")
            print(f"   🔤 Kolom '{kolom}' → {', '.join(aksi_kolom)}")

    if not log["kolom_diproses"]:
        print("   ✅ Tidak ada kolom teks yang perlu distandarisasi.")
    else:
        print(
            f"   ✅ Selesai! {len(log['kolom_diproses'])} kolom teks distandarisasi."
        )
    print()

    return df, log
