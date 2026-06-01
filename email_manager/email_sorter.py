"""
Task 2: Sortir email di Excel berdasarkan penerima langsung
Baca sheet harian, kelompokkan per kategori, tulis ke sheet 'Sorted'.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from pathlib import Path
import config


# Urutan prioritas kategori
CATEGORY_ORDER = [
    "Direct To",
    "To (Multiple)",
    "CC",
    "BCC",
    "Broadcast",
    "Other",
]

CATEGORY_COLORS = {
    "Direct To":     ("1A5276", "D6E4F0"),   # header biru tua, baris biru muda
    "To (Multiple)": ("1F618D", "EBF5FB"),
    "CC":            ("9A7D0A", "FEF9E7"),
    "BCC":           ("922B21", "FDEDEC"),
    "Broadcast":     ("6C3483", "F4ECF7"),
    "Other":         ("424949", "F2F3F4"),
}


def read_today_emails(wb):
    """Baca sheet harian hari ini dan return list of dict."""
    sheet_name = datetime.now().strftime("%Y-%m-%d")
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan. "
            "Jalankan email_fetcher.py terlebih dahulu."
        )

    ws = wb[sheet_name]
    headers = [cell.value for cell in ws[1]]
    emails = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        emails.append(dict(zip(headers, row)))
    return emails


def group_by_category(emails):
    """Kelompokkan email per kategori sesuai CATEGORY_ORDER."""
    groups = {cat: [] for cat in CATEGORY_ORDER}
    for email in emails:
        cat = email.get("Kategori", "Other")
        if cat not in groups:
            cat = "Other"
        groups[cat].append(email)
    return groups


def _section_header(ws, row, label, count, header_color):
    """Tulis baris judul seksi (misal: DIRECT TO — 5 email)."""
    merge_to = 9
    ws.merge_cells(
        start_row=row, start_column=1,
        end_row=row, end_column=merge_to
    )
    cell = ws.cell(row=row, column=1)
    cell.value = f"  {label.upper()}  —  {count} email"
    cell.font = Font(bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill("solid", fgColor=header_color)
    cell.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 20
    return row + 1


def _col_header(ws, row):
    """Tulis header kolom standar."""
    headers = [
        "No", "Waktu", "Subject", "Dari", "Email Pengirim",
        "To", "CC", "Kategori", "Preview Isi",
    ]
    fill = PatternFill("solid", fgColor="2C3E50")
    font = Font(bold=True, color="FFFFFF", size=10)
    border = Border(bottom=Side(style="thin", color="7F8C8D"))
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[row].height = 18
    return row + 1


def _data_row(ws, row, no, email, row_color):
    """Tulis satu baris data email."""
    thin = Border(
        left=Side(style="thin", color="D5D8DC"),
        right=Side(style="thin", color="D5D8DC"),
        top=Side(style="thin", color="D5D8DC"),
        bottom=Side(style="thin", color="D5D8DC"),
    )
    fill = PatternFill("solid", fgColor=row_color)
    values = [
        no,
        email.get("Waktu Diterima", ""),
        email.get("Subject", ""),
        email.get("Dari", ""),
        email.get("Email Pengirim", ""),
        email.get("To", ""),
        email.get("CC", ""),
        email.get("Kategori", ""),
        email.get("Preview Isi", ""),
    ]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.fill = fill
        cell.border = thin
        cell.alignment = Alignment(
            vertical="center",
            wrap_text=(col == 9),
        )
    ws.row_dimensions[row].height = 18


def _setup_col_widths(ws):
    widths = [5, 18, 45, 25, 30, 30, 25, 15, 60]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


def build_sorted_sheet(wb, groups):
    """
    Buat atau timpa sheet 'Sorted_<tanggal>' dengan email yang sudah dikelompokkan.
    """
    sheet_name = "Sorted_" + datetime.now().strftime("%Y-%m-%d")
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    _setup_col_widths(ws)

    # Judul utama
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = (
        f"EMAIL SORTED — {datetime.now().strftime('%A, %d %B %Y')}"
        f"   |   Target: {config.TARGET_NAME}"
    )
    title_cell.font = Font(bold=True, color="FFFFFF", size=13)
    title_cell.fill = PatternFill("solid", fgColor="0B3954")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    current_row = 3  # baris kosong pemisah setelah judul

    stats = {}
    for category in CATEGORY_ORDER:
        emails = groups[category]
        stats[category] = len(emails)
        if not emails:
            continue

        header_color, row_color = CATEGORY_COLORS[category]

        current_row = _section_header(
            ws, current_row, category, len(emails), header_color
        )
        current_row = _col_header(ws, current_row)

        for no, email in enumerate(emails, 1):
            _data_row(ws, current_row, no, email, row_color)
            current_row += 1

        current_row += 1  # baris kosong antar seksi

    return ws, sheet_name, stats


def build_stats_sidebar(wb, stats, sheet_name):
    """Tambah mini-stats di sheet Sorted: total per kategori di kolom K-L."""
    ws = wb[sheet_name]
    start = 3
    ws.cell(start, 11, "Kategori").font = Font(bold=True)
    ws.cell(start, 12, "Jumlah").font = Font(bold=True)
    for i, cat in enumerate(CATEGORY_ORDER, 1):
        ws.cell(start + i, 11, cat)
        ws.cell(start + i, 12, stats[cat])
    ws.cell(start + len(CATEGORY_ORDER) + 1, 11, "TOTAL").font = Font(bold=True)
    ws.cell(
        start + len(CATEGORY_ORDER) + 1, 12, sum(stats.values())
    ).font = Font(bold=True)

    ws.column_dimensions["K"].width = 18
    ws.column_dimensions["L"].width = 10


def run():
    """Entry point Task 2."""
    print("=" * 55)
    print("TASK 2 — Sortir Email by Recipient")
    print("=" * 55)

    filepath = Path(config.EXCEL_PATH)
    if not filepath.exists():
        print(f"  File tidak ditemukan: {filepath}")
        print("  Jalankan Task 1 terlebih dahulu.")
        return

    wb = openpyxl.load_workbook(filepath)
    emails = read_today_emails(wb)
    print(f"  Membaca {len(emails)} email dari sheet harian...")

    groups = group_by_category(emails)

    for cat, lst in groups.items():
        if lst:
            print(f"    {cat:20s}: {len(lst)} email")

    ws, sheet_name, stats = build_sorted_sheet(wb, groups)
    build_stats_sidebar(wb, stats, sheet_name)
    wb.save(filepath)

    print(f"  Sheet '{sheet_name}' berhasil dibuat.")
    print(f"  File disimpan: {filepath}")


if __name__ == "__main__":
    run()
