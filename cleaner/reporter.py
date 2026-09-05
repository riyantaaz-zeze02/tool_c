"""
reporter.py — Generator Sheet Laporan Audit & Ringkasan Pembersihan (Cleaning Summary)
====================================================================================
Menghasilkan lembar kerja Excel kedua ("Cleaning Summary")
yang mencatat metrik transparansi data sebelum vs sesudah.
"""

from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from cleaner.formatter.themes import get_theme


def _translate_action(action_str: str) -> str:
    """Menerjemahkan deskripsi aksi pembersihan ke bahasa Inggris."""
    replacements = [
        ("strip spasi", "strip whitespace"),
        ("hapus non-ASCII", "remove non-ASCII"),
        ("persen → float (desimal)", "percentage → float (decimal)"),
        ("string angka → float", "numeric string → float"),
        ("(numerik)", "(numeric)"),
        ("(datetime)", "(datetime)"),
        ("(teks)", "(text)"),
        ("diisi rata-rata (mean)", "imputed with mean"),
        ("diisi median", "imputed with median"),
        ("diisi modus", "imputed with mode"),
        ("diisi konstanta", "filled with constant"),
        ("diisi 'Tidak Diketahui'", "filled with 'Unknown'"),
        ("diisi", "filled with"),
        ("dikonversi ke datetime", "converted to datetime"),
        ("(dibiarkan)", "(retained)"),
        ("Data sudah bersih", "Data already clean"),
    ]
    res = action_str
    for id_term, en_term in replacements:
        res = res.replace(id_term, en_term)
    return res


class CleaningReporter:
    """
    Membuat sheet audit trail dan ringkasan metrik pembersihan pada workbook openpyxl.
    """

    def __init__(self, theme="corporate_blue"):
        self.theme = get_theme(theme)

    def add_summary_sheet(self, wb, report_dict, df_cleaned, title="Cleaning Summary"):
        """
        Menambahkan worksheet 'Cleaning Summary' ke workbook Excel.

        Parameter:
            wb (openpyxl.Workbook): Workbook yang sedang dibuat.
            report_dict (dict): Laporan dari CleaningEngine.clean().
            df_cleaned (pd.DataFrame): Dataframe hasil pembersihan.
            title (str): Judul sheet summary (default: 'Cleaning Summary').
        """
        ws = wb.create_sheet(title=title)
        ws.views.sheetView[0].showGridLines = True

        font_family = self.theme["font_name"]

        # Styles
        title_font = Font(name=font_family, size=14, bold=True, color=self.theme["header_fill"])
        subtitle_font = Font(name=font_family, size=9, italic=True, color="595959")
        section_font = Font(name=font_family, size=11, bold=True, color="262626")
        label_font = Font(name=font_family, size=10, bold=True, color="404040")
        value_font = Font(name=font_family, size=10)

        header_fill = PatternFill(
            start_color=self.theme["header_fill"],
            end_color=self.theme["header_fill"],
            fill_type="solid"
        )
        header_font = Font(name=font_family, size=10, bold=True, color=self.theme["header_font"])

        border_side = Side(border_style="thin", color=self.theme["border_color"])
        thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
        zebra_fill = PatternFill(
            start_color=self.theme["zebra_fill"],
            end_color=self.theme["zebra_fill"],
            fill_type="solid"
        )

        # 1. Judul & Waktu (Header & Subtitle)
        ws.cell(row=2, column=2, value="DATA CLEANING AUDIT REPORT").font = title_font
        waktu_str = datetime.now().strftime("%d %B %Y, %H:%M:%S")
        ws.cell(row=3, column=2, value=f"YANTTT — {waktu_str}").font = subtitle_font

        # 2. Ringkasan Metrik Kunci (Key Metrics)
        ws.cell(row=5, column=2, value="KEY METRICS").font = section_font

        metrics = [
            ("Source File", report_dict.get("sumber", "Input File")),
            ("Initial Row Count", report_dict.get("baris_awal", 0)),
            ("Final Row Count", report_dict.get("baris_akhir", 0)),
            ("Duplicate Rows Removed", report_dict.get("total_duplikat_dihapus", 0)),
            ("Dropped Rows (Missing)", report_dict.get("total_baris_didrop", 0)),
            ("Total Columns", report_dict.get("kolom_akhir", 0)),
        ]

        for idx, (label, val) in enumerate(metrics, start=6):
            c_label = ws.cell(row=idx, column=2, value=label)
            c_label.font = label_font
            c_label.border = thin_border
            c_label.fill = zebra_fill

            c_val = ws.cell(row=idx, column=3, value=val)
            c_val.font = value_font
            c_val.border = thin_border
            if isinstance(val, (int, float)):
                c_val.alignment = Alignment(horizontal="right")
                c_val.number_format = "#,##0"

        # 3. Tabel Detail Perubahan per Kolom (Column Status Details)
        row_detail_start = 14
        ws.cell(row=row_detail_start - 1, column=2, value="COLUMN STATUS DETAILS").font = section_font

        headers = ["No", "Column Name", "Final Data Type", "Missing Found", "Cleaning Action"]
        for col_i, h_text in enumerate(headers, start=2):
            cell = ws.cell(row=row_detail_start, column=col_i, value=h_text)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        missing_log = report_dict.get("langkah", {}).get("missing", {}).get("missing_sebelum", {})
        actions_list = []
        for step in ["number", "date", "missing", "text"]:
            step_actions = report_dict.get("langkah", {}).get(step, {}).get("aksi", [])
            actions_list.extend(step_actions)

        for col_idx, col_name in enumerate(df_cleaned.columns, start=1):
            r = row_detail_start + col_idx
            is_z = (col_idx % 2 == 1)
            row_fill = zebra_fill if is_z else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

            # Cari aksi yang relevan untuk kolom ini
            col_actions = [a for a in actions_list if f"'{col_name}'" in a]
            if col_actions:
                translated_actions = [_translate_action(a) for a in col_actions]
                action_desc = "; ".join(translated_actions)
            else:
                action_desc = "Data already clean"

            row_data = [
                col_idx,
                col_name,
                str(df_cleaned[col_name].dtype),
                missing_log.get(col_name, 0),
                action_desc,
            ]

            for c_i, val in enumerate(row_data, start=2):
                cell = ws.cell(row=r, column=c_i, value=val)
                cell.font = value_font
                cell.fill = row_fill
                cell.border = thin_border
                if c_i == 2:
                    cell.alignment = Alignment(horizontal="center")
                elif c_i == 5:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "#,##0"

        # Auto-width kolom di sheet Ringkasan
        for c in range(2, 7):
            c_letter = get_column_letter(c)
            max_len = 0
            for r in range(2, row_detail_start + len(df_cleaned.columns) + 2):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[c_letter].width = max(max_len + 4, 15)
