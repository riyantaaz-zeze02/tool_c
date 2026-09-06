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
from cleaner.reader import baca_header, get_sheet_names



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


def analyze_auto_files(file_paths):
    """Menganalisis header file dan merekomendasikan merge, join, atau manual."""
    headers = {path: baca_header(path) for path in file_paths}

    if len({tuple(columns) for columns in headers.values()}) == 1:
        return {
            "mode": "merge",
            "files": list(file_paths),
            "headers": headers,
            "candidate_keys": {},
            "chain": list(file_paths),
            "keys": [],
            "reason": f"Terdeteksi {len(file_paths)} file dengan struktur kolom identik persis.",
        }

    key_files = {}
    for path, columns in headers.items():
        for column in columns:
            key_files.setdefault(column, []).append(path)
    candidate_keys = {key: paths for key, paths in key_files.items() if len(paths) > 1}

    shared_by_all = [key for key, paths in candidate_keys.items() if len(paths) == len(file_paths)]
    if shared_by_all:
        return {
            "mode": "unknown",
            "files": list(file_paths),
            "headers": headers,
            "candidate_keys": candidate_keys,
            "chain": [],
            "keys": [],
            "reason": f"Struktur berbeda dan kandidat key {', '.join(shared_by_all)} muncul di semua file; tool tidak yakin menentukan pasangan join.",
        }

    import itertools

    candidates = []
    for order in itertools.permutations(file_paths):
        edge_keys = []
        score = 0
        valid = True
        for left_path, right_path in zip(order, order[1:]):
            shared = [key for key in headers[left_path] if key in headers[right_path]]
            if not shared:
                valid = False
                break
            edge_keys.append(shared)
            score += len(shared)
        if valid:
            candidates.append((score, order, edge_keys))

    if not candidates:
        return {
            "mode": "unknown",
            "files": list(file_paths),
            "headers": headers,
            "candidate_keys": candidate_keys,
            "chain": [],
            "keys": [],
            "reason": "Struktur berbeda, tetapi tidak ditemukan rantai kolom kunci yang menghubungkan semua file.",
        }

    best_score = max(item[0] for item in candidates)
    best = [item for item in candidates if item[0] == best_score]
    supplied_order = [item for item in best if tuple(item[1]) == tuple(file_paths)]
    if supplied_order:
        best = supplied_order
    unique_shapes = {
        (tuple(item[1]), tuple(tuple(keys) for keys in item[2]))
        for item in best
    }
    if len(unique_shapes) != 1 or any(len(keys) != 1 for keys in best[0][2]):
        return {
            "mode": "unknown",
            "files": list(file_paths),
            "headers": headers,
            "candidate_keys": candidate_keys,
            "chain": [],
            "keys": [],
            "reason": "Struktur berbeda dan kandidat kolom kunci ambigu; tool tidak yakin menentukan urutan join.",
        }

    _, chain, edge_keys = best[0]
    return {
        "mode": "join",
        "files": list(file_paths),
        "headers": headers,
        "candidate_keys": candidate_keys,
        "chain": list(chain),
        "keys": [edge[0] for edge in edge_keys],
        "reason": f"Terdeteksi {len(file_paths)} file dengan struktur berbeda dan rantai key yang jelas.",
    }


def print_auto_analysis(analysis):
    """Menampilkan hasil analisis auto secara ringkas dan dapat diaudit."""
    files = analysis["files"]
    headers = analysis["headers"]
    print(f"🔎 Analisis header-only untuk {len(files)} file:")
    for path in files:
        print(f"   • {os.path.basename(path)}: {', '.join(headers[path])}")
    print(f"📌 {analysis['reason']}")
    for key, paths in analysis["candidate_keys"].items():
        names = " ↔ ".join(os.path.basename(path) for path in paths)
        print(f"🔑 Kandidat key '{key}': {names}")
    if analysis["mode"] == "join":
        chain_names = [os.path.basename(path) for path in analysis["chain"]]
        print(f"✅ Rekomendasi: mode JOIN ({' → '.join(chain_names)}; keys: {', '.join(analysis['keys'])})")
    elif analysis["mode"] == "merge":
        print("✅ Rekomendasi: mode TUMPUK / MERGE")
    else:
        print("⚠️  Rekomendasi: TIDAK YAKIN. Tentukan --merge atau --join secara manual.")


def process_single_file(file_path, args, engine, config):
    """Memproses satu file dan mengekspor hasilnya."""
    print("=" * 60)
    print(f"🚀 Memproses: {file_path}")
    print("=" * 60)

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    out_format = args.format.lower()
    ekstensi = os.path.splitext(file_path)[1].lower()

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

    if ekstensi in [".xlsx", ".xls"]:
        available_sheets = get_sheet_names(file_path)
        print(f"📋 File ini punya {len(available_sheets)} sheet: {', '.join(available_sheets)}")

        if getattr(args, "all_sheets", False):
            target_sheets = available_sheets
        elif getattr(args, "sheets", None):
            target_sheets = [s.strip() for s in args.sheets.split(",") if s.strip()]
            for s in target_sheets:
                if s not in available_sheets:
                    print(f"❌ Error: Sheet '{s}' tidak ditemukan di file. Sheet yang tersedia: {', '.join(available_sheets)}")
                    return False
        else:
            # Perilaku default: sheet pertama
            target_sheets = [available_sheets[0]] if available_sheets else [0]

        print(f"🎯 Sheet yang dipilih untuk dibersihkan: {', '.join(str(s) for s in target_sheets)}")
        unselected = [s for s in available_sheets if s not in target_sheets]
        if unselected:
            print(f"📌 Sheet yang tidak dipilih akan disalin apa adanya: {', '.join(unselected)}")
        print()

        try:
            cleaned_results = engine.clean_sheets(file_path, sheets=target_sheets)
        except Exception as e:
            print(f"❌ Terjadi kesalahan saat membersihkan {file_path}: {e}")
            return False

        if out_format == "excel":
            engine.export_excel(
                cleaned_results,
                out_file,
                theme=theme,
                currency=currency,
                add_total_row=add_total,
                include_summary_sheet=include_summary,
                original_file_path=file_path,
            )
        else:
            os.makedirs(os.path.dirname(out_file), exist_ok=True)
            for s_name, data in cleaned_results.items():
                s_out = out_file if len(cleaned_results) == 1 else out_file.replace(".csv", f"_{s_name}.csv")
                data["df"].to_csv(s_out, index=False)
                print(f"💾 File CSV berhasil disimpan ke: {s_out}")

        print()
        print("📊 Ringkasan Pembersihan per Sheet:")
        for s_name, res in cleaned_results.items():
            rep = res["report"]
            print(f"   • [{s_name}]")
            print(f"     - Baris: {rep['baris_awal']} ➔ {rep['baris_akhir']}")
            print(f"     - Duplikat dihapus: {rep['total_duplikat_dihapus']}")
            print(f"     - Baris didrop: {rep['total_baris_didrop']}")
        print("=" * 60)
        print()
        return True

    else:
        # File CSV tunggal
        try:
            df, report = engine.clean(file_path)
        except Exception as e:
            print(f"❌ Terjadi kesalahan saat membersihkan {file_path}: {e}")
            return False

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


def process_merged_files(file_paths, args, engine, config):
    """Menggabungkan, membersihkan, dan mengekspor beberapa file sebagai satu hasil."""
    print("=" * 60)
    print(f"🚀 Menggabungkan {len(file_paths)} file")
    print("=" * 60)

    base_name = os.path.splitext(os.path.basename(file_paths[0]))[0]
    out_format = args.format.lower()
    if args.output:
        if os.path.isdir(args.output):
            extension = ".xlsx" if out_format == "excel" else ".csv"
            out_file = os.path.join(args.output, f"{base_name}_merged_bersih{extension}")
        else:
            out_file = args.output
    else:
        extension = ".xlsx" if out_format == "excel" else ".csv"
        out_file = os.path.join("data", "output", f"{base_name}_merged_bersih{extension}")

    formatting_cfg = config.get("formatting", {})
    theme = args.theme or formatting_cfg.get("theme", "corporate_blue")
    currency = args.currency or formatting_cfg.get("currency", "IDR")
    sheet_name = formatting_cfg.get("sheet_name", "Data Bersih")
    add_total = not args.no_total if args.no_total is not None else formatting_cfg.get("add_total_row", True)
    include_summary = not args.no_summary if args.no_summary is not None else formatting_cfg.get("include_summary_sheet", True)

    try:
        df, report = engine.clean_merged_files(file_paths)
    except Exception as e:
        print(f"❌ Gagal menggabungkan file: {e}")
        return False

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
        output_dir = os.path.dirname(out_file)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        df.to_csv(out_file, index=False)
        print(f"💾 File CSV hasil merge berhasil disimpan ke: {out_file}")

    print()
    print("📊 Ringkasan Merge:")
    print(f"   - File digabung: {report['total_files']}")
    print(f"   - Baris awal total: {report['baris_awal']} | Baris akhir: {report['baris_akhir']}")
    print(f"   - Duplikat antar-file: {report['duplikat_antar_file']}")
    print("=" * 60)
    print()
    return True


def process_joined_files(file_paths, keys, args, engine, config):
    """Menjalankan chained join, cleaning, dan ekspor satu hasil akhir."""
    print("=" * 60)
    print(f"🚀 Menjalankan chained join untuk {len(file_paths)} file")
    print("=" * 60)

    base_name = os.path.splitext(os.path.basename(file_paths[0]))[0]
    out_format = args.format.lower()
    extension = ".xlsx" if out_format == "excel" else ".csv"
    if args.output:
        out_file = os.path.join(args.output, f"{base_name}_joined_bersih{extension}") if os.path.isdir(args.output) else args.output
    else:
        out_file = os.path.join("data", "output", f"{base_name}_joined_bersih{extension}")

    formatting_cfg = config.get("formatting", {})
    theme = args.theme or formatting_cfg.get("theme", "corporate_blue")
    currency = args.currency or formatting_cfg.get("currency", "IDR")
    sheet_name = formatting_cfg.get("sheet_name", "Data Bersih")
    add_total = not args.no_total if args.no_total is not None else formatting_cfg.get("add_total_row", True)
    include_summary = not args.no_summary if args.no_summary is not None else formatting_cfg.get("include_summary_sheet", True)

    try:
        df, report = engine.clean_joined_files(file_paths, keys)
    except Exception as e:
        print(f"❌ Gagal menjalankan join: {e}")
        return False

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
        output_dir = os.path.dirname(out_file)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        df.to_csv(out_file, index=False)
        print(f"💾 File CSV hasil join berhasil disimpan ke: {out_file}")

    print("\n📊 Match rate per tahap:")
    for stage in report["join_stages"]:
        print(f"   - {stage['left']} + {stage['right']} (key: {stage['key']}) → {stage['match_rate']}% match")
    print("=" * 60)
    print()
    return True


def main():
    parser = argparse.ArgumentParser(
        description="ExcelCleaner Pro — Tool Pembersih Data & Formatter Excel Profesional"
    )
    parser.add_argument(
        "input",
        nargs="?",
        help="Path ke file input (CSV/Excel) atau direktori jika menggunakan mode -b/--batch",
    )
    parser.add_argument(
        "-s", "--sheets",
        help="Daftar nama sheet yang mau dibersihkan, dipisah koma (contoh: --sheets 'Penjualan,Pelanggan')",
        default=None,
    )
    parser.add_argument(
        "--all-sheets",
        action="store_true",
        help="Bersihkan SEMUA sheet yang ada di file Excel sekaligus",
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
    parser.add_argument(
        "--merge",
        help="Gabungkan daftar file CSV/Excel yang dipisahkan koma menjadi satu input",
        default=None,
    )
    parser.add_argument(
        "--merge-folder",
        help="Gabungkan semua file CSV/Excel dalam folder menjadi satu input",
        default=None,
    )
    parser.add_argument(
        "--join",
        help="Join berantai daftar file CSV/Excel yang dipisahkan koma",
        default=None,
    )
    parser.add_argument(
        "--keys",
        help="Daftar key join berurutan, dipisahkan koma (N file membutuhkan N-1 key)",
        default=None,
    )
    parser.add_argument(
        "--auto",
        help="Analisis header beberapa file lalu rekomendasikan merge atau join",
        default=None,
    )


    args = parser.parse_args()

    print("╔══════════════════════════════════════════════════════════╗")
    print("║          🧹 ExcelCleaner Pro — CLI Runner               ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()

    if args.merge and args.merge_folder:
        parser.error("Gunakan salah satu --merge atau --merge-folder, bukan keduanya.")
    active_modes = [bool(args.batch), bool(args.merge or args.merge_folder), bool(args.join), bool(args.auto)]
    if sum(active_modes) > 1:
        parser.error("Gunakan hanya satu mode: input tunggal, --batch, --merge, --join, atau --auto.")
    if args.join and not args.keys:
        parser.error("Mode --join membutuhkan --keys.")
    if not args.input and not (args.merge or args.merge_folder or args.join or args.auto):
        parser.error("Input file wajib diisi, kecuali menggunakan --merge, --merge-folder, --join, atau --auto.")

    config = load_yaml_config(args.config)
    engine = CleaningEngine(config=config)

    if args.auto:
        files = [path.strip() for path in args.auto.split(",") if path.strip()]
        valid_exts = [".csv", ".xlsx", ".xls"]
        if len(files) < 2:
            parser.error("Mode --auto membutuhkan minimal 2 file.")
        missing_files = [path for path in files if not os.path.isfile(path)]
        unsupported_files = [path for path in files if os.path.splitext(path)[1].lower() not in valid_exts]
        if missing_files:
            parser.error(f"File tidak ditemukan: {', '.join(missing_files)}")
        if unsupported_files:
            parser.error(f"Format file tidak didukung: {', '.join(unsupported_files)}")
        try:
            analysis = analyze_auto_files(files)
        except Exception as error:
            parser.error(f"Gagal membaca header file: {error}")
        print_auto_analysis(analysis)
        if analysis["mode"] == "unknown":
            sys.exit(2)
        confirmation = input("Lanjutkan dengan mode ini? [y/n]: ").strip().lower()
        if confirmation not in {"y", "yes"}:
            print("⏹️  Dibatalkan. Tidak ada data yang diproses.")
            sys.exit(0)
        if analysis["mode"] == "merge":
            if not process_merged_files(files, args, engine, config):
                sys.exit(1)
        elif not process_joined_files(analysis["chain"], analysis["keys"], args, engine, config):
            sys.exit(1)

    elif args.join:
        files = [path.strip() for path in args.join.split(",") if path.strip()]
        keys = [key.strip() for key in args.keys.split(",") if key.strip()]
        valid_exts = [".csv", ".xlsx", ".xls"]
        if len(files) < 2:
            parser.error("Mode --join membutuhkan minimal 2 file.")
        if len(keys) != len(files) - 1:
            parser.error(
                f"Jumlah key tidak sesuai: {len(files)} file membutuhkan {len(files) - 1} key, tetapi menerima {len(keys)}."
            )
        invalid_files = [path for path in files if not os.path.isfile(path)]
        unsupported_files = [path for path in files if os.path.splitext(path)[1].lower() not in valid_exts]
        if invalid_files:
            parser.error(f"File tidak ditemukan: {', '.join(invalid_files)}")
        if unsupported_files:
            parser.error(f"Format file tidak didukung: {', '.join(unsupported_files)}")
        print("📂 File yang akan di-join:")
        for file_path in files:
            print(f"   • {file_path}")
        print(f"🔑 Key berurutan: {', '.join(keys)}\n")
        if not process_joined_files(files, keys, args, engine, config):
            sys.exit(1)

    elif args.merge or args.merge_folder:
        valid_exts = [".csv", ".xlsx", ".xls"]
        if args.merge:
            files = [path.strip() for path in args.merge.split(",") if path.strip()]
        else:
            if not os.path.isdir(args.merge_folder):
                print(f"❌ Error: '{args.merge_folder}' bukan folder yang valid untuk mode merge.")
                sys.exit(1)
            files = [
                os.path.join(args.merge_folder, name)
                for name in sorted(os.listdir(args.merge_folder))
                if os.path.splitext(name)[1].lower() in valid_exts
            ]

        if len(files) < 2:
            print("❌ Error: Mode merge membutuhkan minimal 2 file CSV/Excel.")
            sys.exit(1)
        missing_files = [path for path in files if not os.path.isfile(path)]
        if missing_files:
            print(f"❌ Error: File tidak ditemukan: {', '.join(missing_files)}")
            sys.exit(1)

        print("📂 File yang akan digabungkan:")
        for file_path in files:
            print(f"   • {file_path}")
        print()
        if not process_merged_files(files, args, engine, config):
            sys.exit(1)

    elif args.batch:
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
