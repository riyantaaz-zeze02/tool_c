"""
cli.py — Command Line Interface untuk ExcelCleaner Pro
======================================================
Cara penggunaan:
    python cli.py data/input/contoh_data.csv
    python cli.py data/input/contoh_data.csv -t modern_slate -c IDR
    python cli.py data/input/ -b
"""

import argparse
import os
import sys
import yaml

from cleaner.engine import CleaningEngine
from cleaner.formatter.themes import THEMES


def load_yaml_config(config_path):
    """Membaca file konfigurasi YAML jika tersedia."""
    if not config_path:
        default_path = os.path.join("config", "default_config.yaml")
        if os.path.exists(default_path):
            config_path = default_path
        else:
            return {}

    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            print(f"⚠️  Gagal memuat config '{config_path}': {e}. Menggunakan default.")
            return {}
    return {}


def process_single_file(file_path, args, engine, config):
    """Memproses satu file dan mengekspor hasilnya."""
    print("=" * 60)
    print(f"🚀 Memproses: {file_path}")
    print("=" * 60)

    try:
        df, report = engine.clean(file_path)
    except Exception as e:
        print(f"❌ Terjadi kesalahan saat membersihkan {file_path}: {e}")
        return False

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    out_format = args.format.lower()

    if args.output:
        if os.path.isdir(args.output):
            ext = ".xlsx" if out_format == "excel" else ".csv"
            out_file = os.path.join(args.output, f"{base_name}_bersih{ext}")
        else:
            out_file = args.output
    else:
        ext = ".xlsx" if out_format == "excel" else ".csv"
        out_file = os.path.join("data", "output", f"{base_name}_bersih{ext}")

    formatting_cfg = config.get("formatting", {})
    theme = args.theme or formatting_cfg.get("theme", "corporate_blue")
    currency = args.currency or formatting_cfg.get("currency", "IDR")
    sheet_name = formatting_cfg.get("sheet_name", "Data Bersih")
    add_total = not args.no_total if args.no_total is not None else formatting_cfg.get("add_total_row", True)
    include_summary = not args.no_summary if args.no_summary is not None else formatting_cfg.get("include_summary_sheet", True)

    if out_format == "excel":
        engine.export_excel(
            df,
            out_file,
            report=report,
            theme=theme,
            currency=currency,
            sheet_name=sheet_name,
            add_total_row=add_total,
            include_summary_sheet=include_summary,
        )
    else:
        os.makedirs(os.path.dirname(out_file), exist_ok=True)
        df.to_csv(out_file, index=False)
        print(f"💾 File CSV berhasil disimpan ke: {out_file}")

    print()
    print("📊 Ringkasan:")
    print(f"   - Baris awal: {report['baris_awal']} | Baris akhir: {report['baris_akhir']}")
    print(f"   - Duplikat dihapus: {report['total_duplikat_dihapus']}")
    print("=" * 60)
    print()
    return True


def main():
    parser = argparse.ArgumentParser(
        description="ExcelCleaner Pro — Tool Pembersih Data & Formatter Excel Profesional"
    )
    parser.add_argument(
        "input",
        help="Path ke file input (CSV/Excel) atau direktori jika menggunakan mode -b/--batch",
    )
    parser.add_argument(
        "-o", "--output",
        help="Path file output atau folder tujuan",
        default=None,
    )
    parser.add_argument(
        "-t", "--theme",
        help=f"Tema warna Excel. Pilihan: {', '.join(THEMES.keys())}",
        default=None,
    )
    parser.add_argument(
        "-c", "--currency",
        help="Format mata uang default (IDR atau USD)",
        choices=["IDR", "USD"],
        default=None,
    )
    parser.add_argument(
        "--format",
        help="Format file output (excel atau csv)",
        choices=["excel", "csv"],
        default="excel",
    )
    parser.add_argument(
        "--config",
        help="Path ke file konfigurasi custom YAML",
        default=None,
    )
    parser.add_argument(
        "--no-total",
        action="store_true",
        help="Jangan tambahkan baris TOTAL di bawah tabel Excel",
    )
    parser.add_argument(
        "--no-summary",
        action="store_true",
        help="Jangan buat sheet kedua 'Cleaning Summary'",
    )
    parser.add_argument(
        "-b", "--batch",
        action="store_true",
        help="Mode batch: bersihkan semua file .csv dan .xlsx di direktori input",
    )

    args = parser.parse_args()

    print("╔══════════════════════════════════════════════════════════╗")
    print("║          🧹 ExcelCleaner Pro — CLI Runner               ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()

    config = load_yaml_config(args.config)
    engine = CleaningEngine(config=config)

    if args.batch:
        if not os.path.isdir(args.input):
            print(f"❌ Error: '{args.input}' bukan folder yang valid untuk mode batch.")
            sys.exit(1)

        valid_exts = [".csv", ".xlsx", ".xls"]
        files = [
            os.path.join(args.input, f)
            for f in os.listdir(args.input)
            if os.path.splitext(f)[1].lower() in valid_exts
        ]

        if not files:
            print(f"⚠️  Tidak ditemukan file CSV/Excel di folder '{args.input}'.")
            sys.exit(0)

        print(f"📂 Ditemukan {len(files)} file untuk dibersihkan:")
        for f in files:
            print(f"   • {f}")
        print()

        sukses = 0
        for f in files:
            if process_single_file(f, args, engine, config):
                sukses += 1

        print(f"🎉 Selesai memproses {sukses}/{len(files)} file!")

    else:
        if not os.path.isfile(args.input):
            print(f"❌ Error: File '{args.input}' tidak ditemukan.")
            sys.exit(1)
        process_single_file(args.input, args, engine, config)


if __name__ == "__main__":
    main()
