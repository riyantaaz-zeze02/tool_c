"""
excel_styler.py — Engine Styling Tabel Excel Profesional (openpyxl)
===================================================================
Menghasilkan tampilan spreadsheet standar korporat:
- Header styling (warna tema, bold, kontras)
- Zebra striping (baris selang-seling)
- Auto-fit lebar kolom dinamis + margin
- Freeze top row (header tetap terlihat saat scroll)
- Auto-filter pada header
- Border tipis dan rapi
- Format sel native (Currency Rp/$, Tanggal, Desimal, Persen)
- Baris total otomatis dengan formula Excel native (=SUM)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from cleaner.formatter.themes import get_theme
from cleaner.formatter.number_formats import get_column_format_mask


class ExcelStyler:
    """
    Penata gaya lembar kerja Excel berbasis openpyxl.
    """

    def __init__(self, theme="corporate_blue", currency="IDR"):
        self.theme = get_theme(theme)
        self.currency = currency

    def style_worksheet(self, ws, df, add_total_row=True, auto_filter=True, freeze_header=True):
        """
        Menerapkan styling penuh ke worksheet openpyxl berdasarkan DataFrame referensi.

        Parameter:
            ws (openpyxl.worksheet.worksheet.Worksheet): Worksheet target.
            df (pd.DataFrame): DataFrame yang berisi data dalam worksheet tersebut.
            add_total_row (bool): Apakah menambahkan baris Total di bawah tabel.
            auto_filter (bool): Apakah mengaktifkan auto-filter pada baris header.
            freeze_header (bool): Apakah membekukan (freeze) baris header.
        """
        font_family = self.theme["font_name"]

        # Definisi Styles
        header_font = Font(name=font_family, size=11, bold=True, color=self.theme["header_font"])
        header_fill = PatternFill(
            start_color=self.theme["header_fill"],
            end_color=self.theme["header_fill"],
            fill_type="solid"
        )
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=False)

        data_font = Font(name=font_family, size=10)
        zebra_fill = PatternFill(
            start_color=self.theme["zebra_fill"],
            end_color=self.theme["zebra_fill"],
            fill_type="solid"
        )
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        border_side = Side(border_style="thin", color=self.theme["border_color"])
        thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

        num_rows = len(df)
        num_cols = len(df.columns)

        if num_rows == 0 or num_cols == 0:
            return

        # 1. Styling Baris Header (Row 1)
        ws.row_dimensions[1].height = 28
        for col_idx in range(1, num_cols + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # 2. Styling Sel Data & Format Mask (Row 2 sampai num_rows + 1)
        col_formats = {}
        for col_idx, col_name in enumerate(df.columns, start=1):
            series = df[col_name]
            mask = get_column_format_mask(col_name, series.dtype, default_currency=self.currency)
            col_formats[col_idx] = mask

        for row_idx in range(2, num_rows + 2):
            ws.row_dimensions[row_idx].height = 20
            # Zebra striping: baris genap berwarna, ganjil putih
            is_zebra = (row_idx % 2 == 1)
            row_fill = zebra_fill if is_zebra else white_fill

            for col_idx in range(1, num_cols + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = data_font
                cell.fill = row_fill
                cell.border = thin_border

                # Tentukan alignment & format number
                mask = col_formats.get(col_idx)
                if mask:
                    cell.number_format = mask
                    if "currency" in mask.lower() or "#" in mask or "%" in mask:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif "yy" in mask.lower():
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        current_last_row = num_rows + 1

        # 3. Baris Total / Ringkasan (Opsional)
        if add_total_row and num_rows > 0:
            total_row_idx = current_last_row + 1
            ws.row_dimensions[total_row_idx].height = 24

            total_fill = PatternFill(
                start_color=self.theme["total_fill"],
                end_color=self.theme["total_fill"],
                fill_type="solid"
            )
            total_font = Font(name=font_family, size=10, bold=True, color=self.theme["total_font"])
            top_thin = Side(border_style="thin", color=self.theme["border_color"])
            bottom_double = Side(border_style="double", color=self.theme["header_fill"])
            total_border = Border(left=border_side, right=border_side, top=top_thin, bottom=bottom_double)

            # Label Total di kolom 1
            first_cell = ws.cell(row=total_row_idx, column=1)
            first_cell.value = "TOTAL"
            first_cell.font = total_font
            first_cell.fill = total_fill
            first_cell.alignment = Alignment(horizontal="center", vertical="center")
            first_cell.border = total_border

            # Cari kolom numerik untuk di-SUM otomatis
            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = ws.cell(row=total_row_idx, column=col_idx)
                cell.font = total_font
                cell.fill = total_fill
                cell.border = total_border

                if col_idx == 1:
                    continue

                series = df[col_name]
                col_letter = get_column_letter(col_idx)

                # Jika kolom numerik dan bukan kolom umur/tahun/ID, beri rumus SUM
                col_lower = col_name.lower()
                skip_sum = ["umur", "age", "id", "nik", "nip", "tahun", "year", "kode"]
                if ("int" in str(series.dtype).lower() or "float" in str(series.dtype).lower()) and not any(k in col_lower for k in skip_sum):
                    cell.value = f"=SUM({col_letter}2:{col_letter}{current_last_row})"
                    mask = col_formats.get(col_idx)
                    if mask:
                        cell.number_format = mask
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            current_last_row = total_row_idx

        # 4. Freeze Panes pada Baris Header
        if freeze_header:
            ws.freeze_panes = "A2"

        # 5. Aktifkan Auto-Filter pada Baris Header
        if auto_filter:
            last_col_letter = get_column_letter(num_cols)
            ws.auto_filter.ref = f"A1:{last_col_letter}{num_rows + 1}"

        # 6. Auto-Fit Column Width (Lebar kolom dinamis dengan padding)
        for col_idx in range(1, num_cols + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0

            # Cek panjang header
            header_val = ws.cell(row=1, column=col_idx).value
            if header_val is not None:
                max_len = max(max_len, len(str(header_val)))

            # Cek sampel baris data (maks 100 baris untuk efisiensi)
            for r_idx in range(2, min(current_last_row + 1, 102)):
                val = ws.cell(row=r_idx, column=col_idx).value
                if val is not None:
                    # Perhitungkan formatting currency yang membuat tampilan lebih panjang
                    val_str = str(val)
                    if col_formats.get(col_idx) and "Rp" in col_formats.get(col_idx):
                        val_str = "Rp " + val_str
                    max_len = max(max_len, len(val_str))

            # Beri padding minimum 13 dan margin +4
            adjusted_width = max(max_len + 5, 13)
            # Batasi maksimum agar tidak terlalu lebar
            adjusted_width = min(adjusted_width, 50)
            ws.column_dimensions[col_letter].width = adjusted_width
