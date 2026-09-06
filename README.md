# 🧹 ExcelCleaner Pro — Data Cleaning & Excel Formatting Tool

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/pandas-3.0%2B-150458.svg)](https://pandas.pydata.org/)
[![openpyxl](https://img.shields.io/badge/openpyxl-3.1%2B-green.svg)](https://openpyxl.readthedocs.io/)
[![Streamlit](https://img.shields.io/badge/streamlit-1.40%2B-FF4B4B.svg)](https://streamlit.io/)
[![Tests](https://img.shields.io/badge/tests-9%20passed-brightgreen.svg)]()

**ExcelCleaner Pro** adalah tool otomatisasi pembersihan data (*data cleaning*) dan penataan format spreadsheet (*Excel formatting*) tingkat lanjut berbasis Python. Tool ini dirancang untuk mengubah data kotor berantakan (CSV / Excel) menjadi file Excel profesional berstandar korporat yang siap dipresentasikan.

---

## 🌟 Fitur Utama

### 1. 🧹 Advanced Data Cleaning Engine
* **Missing Value Handling Fleksibel**:
  - Kolom Angka: Median (tahan outlier), Mean, Mode, atau Nilai Konstan.
  - Kolom Teks: Nilai konstan (`"Unknown"`, `"N/A"`), Forward Fill, atau Backward Fill.
  - Kolom Tanggal: Mempertahankan `NaT` yang valid tanpa merusak tipe data.
* **Deduplikasi Tingkat Lanjut**:
  - Hapus duplikat berdasarkan seluruh kolom atau subset kolom kunci (misal: ID / Email).
  - Pilihan mempertahankan data pertama (`first`) atau terakhir (`last`).
* **Standarisasi & Normalisasi Teks**:
  - Pembersihan spasi ganda dan spasi tersembunyi (Unicode NFKC / non-breaking space).
  - Pilihan kapitalisasi: **Title Case**, **UPPERCASE**, **lowercase**, atau **Sentence case**.
* **Pembersihan Tipe Data Spesifik**:
  - **Mata Uang & Finansial**: Konversi otomatis string `Rp 1.500.000,00` atau `$1,500.50` ke angka float bersih.
  - **Persentase**: Konversi `"15%"` ke desimal `0.15`.
  - **Tanggal Multi-Format**: Parsing otomatis aneka format tanggal (`YYYY-MM-DD`, `DD/MM/YYYY`, dll) ke datetime native.

### 2. 🎨 Professional Excel Formatting (openpyxl)
* **Tema Warna Korporat**: Pilihan palet warna elegan (`Corporate Blue`, `Modern Slate`, `Emerald Teal`, `Sunset Coral`, `Classic Office`).
* **Zebra Striping**: Selang-seling warna baris genap/ganjil untuk kenyamanan membaca.
* **Auto-Fit Column Width**: Lebar kolom otomatis menyesuaikan panjang teks + margin aman.
* **Freeze Top Row**: Baris header tetap terlihat saat tabel di-scroll ke bawah.
* **Auto-Filter**: Tombol dropdown filter otomatis aktif di baris header.
* **Number Format Masks Native**: Angka diformat langsung sebagai uang (`Rp #,##0`), tanggal (`yyyy-mm-dd`), atau angka desimal (`#,##0.00`).
* **Baris Total Otomatis**: Menambahkan baris TOTAL di bagian bawah tabel dengan formula native `=SUM()`.

### 3. 📋 Lembar Laporan Audit Trail (Sheet Ringkasan)
Menghasilkan sheet kedua bernama **`Cleaning Summary`** pada file Excel output yang mencatat metrik transparansi:
* Total baris awal vs baris akhir.
* Jumlah duplikat yang dihapus.
* Rincian setiap aksi pembersihan yang dilakukan per kolom.

---

## 📁 Struktur Proyek

```
tool_cleaning/
├── cleaner/                     # Library Inti (Core Package)
│   ├── __init__.py              # Package entry point & Windows UTF-8 safety
│   ├── reader.py                # Pembaca multi-format (CSV, XLSX, XLS)
│   ├── engine.py                # Pipeline orkestrator cleaning & ekspor Excel
│   ├── rules/                   # Modul aturan pembersihan
│   │   ├── __init__.py
│   │   ├── missing_handler.py   # Penanganan missing values
│   │   ├── duplicate_handler.py # Deduplikasi
│   │   ├── text_cleaner.py      # Normalisasi teks & case formatting
│   │   ├── number_cleaner.py    # Konversi currency & angka
│   │   └── date_cleaner.py      # Parsing tanggal mixed format
│   ├── formatter/               # Modul styling Excel
│   │   ├── __init__.py
│   │   ├── excel_styler.py      # openpyxl engine (Header, Border, Zebra, Auto-Width)
│   │   ├── number_formats.py    # Format mask currency, date, numbers
│   │   └── themes.py            # Palet tema warna
│   └── reporter.py              # Generator sheet audit trail
├── config/
│   └── default_config.yaml      # File konfigurasi aturan pembersihan
├── data/
│   ├── input/                   # Folder file kotor (contoh_data.csv)
│   └── output/                  # Folder hasil output (.xlsx / .csv)
├── tests/                       # Test suite otomatis (pytest)
│   ├── test_cleaner.py
│   ├── test_formatter.py
│   └── test_pipeline.py
├── app.py                       # Web UI Studio interaktif (Streamlit)
├── cli.py                       # Antarmuka baris perintah (CLI runner)
├── data_cleaner.py              # Script kompatibilitas sederhana
├── requirements.txt             # Dependensi pustaka
└── README.md                    # Dokumentasi lengkap
```

---

## 🚀 Instalasi & Persiapan

1. Pastikan Python 3.10+ sudah terpasang di komputer Anda.
2. Install dependensi proyek:
   ```bash
   pip install -r requirements.txt
   ```

---

## 💻 Panduan Penggunaan

### 1. Menggunakan Web UI Studio (Paling Direkomendasikan)
Jalankan dashboard interaktif berbasis web dengan Streamlit:
```bash
streamlit run app.py
```
**Fitur Web Studio:**
- Drag-and-drop file CSV atau Excel.
- Tombol 1-klik *"Gunakan Data Contoh"*.
- Live preview sebelum vs sesudah pembersihan.
- Pilih tema warna, mata uang, dan strategi missing value langsung dari panel samping.
- Tombol *"Unduh Excel Rapi"* untuk menyimpan file `.xlsx` yang sudah diformat lengkap.

---

### 2. Menggunakan Command Line Interface (CLI)
Gunakan perintah terminal yang fleksibel:

* **Pembersihan Standar:**
  ```bash
  python cli.py data/input/contoh_data.csv
  ```
  *Output tersimpan otomatis di `data/output/contoh_data_bersih.xlsx`.*

* **Mengubah Tema Warna & Mata Uang:**
  ```bash
  python cli.py data/input/contoh_data.csv -t emerald_green -c IDR
  ```
  *Pilihan tema: `corporate_blue`, `modern_slate`, `emerald_green`, `sunset_coral`, `classic_excel`.*

* **Menentukan File Output Sendiri:**
  ```bash
  python cli.py data/input/contoh_data.csv -o laporan_keuangan_bersih.xlsx
  ```

* **Mode Batch (Membersihkan Semua File di Suatu Folder):**
  ```bash
  python cli.py data/input/ --batch -t modern_slate
  ```
  Setiap file diproses dan diekspor sebagai hasil terpisah.

* **Mode Merge (Menggabungkan Beberapa File Menjadi Satu):**
  ```bash
  python cli.py --merge "data/input/jan.csv,data/input/feb.csv,data/input/mar.csv"
  python cli.py --merge-folder data/input/ --format csv
  ```
  Mode merge hanya menerima file dengan urutan nama kolom yang sama persis. Setiap baris diberi kolom `Sumber File`, lalu seluruh data dibersihkan sebagai satu pipeline sehingga duplikat antar-file ikut terdeteksi. Excel output menyertakan sheet `Cleaning Summary` dengan jumlah baris awal per file dan rincian duplikat internal maupun antar-file. Mode `--merge` dan `--merge-folder` terpisah dari `--batch`.

* **Ekspor sebagai File CSV Saja:**
  ```bash
  python cli.py data/input/contoh_data.csv --format csv
  ```

---

### 3. Menggunakan Script Sederhana (`data_cleaner.py`)
Bagi yang terbiasa dengan script awal:
```bash
python data_cleaner.py data/input/contoh_data.csv
```
*Script ini otomatis menghasilkan file `.csv` sekaligus file `.xlsx` berformat rapi.*

---

### 4. Integrasi via Python Code (API)
Anda dapat mengimpor modul `cleaner` langsung ke dalam script atau notebook Anda:

```python
from cleaner import CleaningEngine

# Inisialisasi engine
engine = CleaningEngine()

# 1. Jalankan pembersihan
df_clean, report = engine.clean("data/input/contoh_data.csv")

# 2. Ekspor ke Excel dengan format profesional
engine.export_excel(
    df_clean,
    output_path="data/output/laporan.xlsx",
    report=report,
    theme="corporate_blue",
    currency="IDR",
    add_total_row=True,
    include_summary_sheet=True
)
```

---

## ⚙️ Kustomisasi Konfigurasi (`default_config.yaml`)

Anda dapat mengubah aturan pembersihan di [config/default_config.yaml](file:///c:/Users/haniw/OneDrive/Desktop/repos/tool_cleaning/config/default_config.yaml):

```yaml
missing:
  numeric_strategy: median    # median | mean | mode | constant | interpolate
  text_strategy: constant     # constant | ffill | bfill
  numeric_constant: 0
  text_constant: Unknown
  datetime_strategy: keep_nat

duplicate:
  subset: null                # null = semua kolom, atau ['email']
  keep: first                 # first | last

text:
  case_format: title          # title | upper | lower | sentence
  strip_whitespace: true
  normalize_unicode: true
  kolom_exclude:
    - email

formatting:
  theme: corporate_blue       # corporate_blue | modern_slate | emerald_green | sunset_coral | classic_excel
  currency: IDR
  add_total_row: true
  include_summary_sheet: true
```

---

## 🧪 Menjalankan Pengujian Otomatis (Testing)

Proyek ini dilengkapi dengan unit test dan integration test menyeluruh menggunakan `pytest`:

```bash
python -m pytest tests/
```

Hasil test memverifikasi:
- Penanganan missing values dan deteksi tipe numerik/datetime.
- Preservasi nilai `NaN` pada operasi teks.
- Pembersihan mata uang (Rp dan $) serta format persentase.
- Parsing tanggal dengan format campuran (*mixed date formats*).
- Pembuatan styling Excel, baris TOTAL `=SUM()`, freeze panes, dan auto-filter.
- Pipeline pembersihan end-to-end dari file CSV hingga file Excel akhir.
