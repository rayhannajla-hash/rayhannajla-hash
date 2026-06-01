"""
Task 1: Copy email baru dari Outlook ke Excel
Baca inbox Outlook via win32com, simpan ke sheet harian di Excel.
"""

import win32com.client
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
from pathlib import Path
import re
import config


def get_outlook_inbox():
    """Koneksi ke Outlook dan return folder Inbox."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    return inbox


def parse_recipients(recipients_obj):
    """Ekstrak semua alamat email dari Recipients object."""
    addresses = []
    for recipient in recipients_obj:
        addr = recipient.Address.lower().strip()
        name = recipient.Name.strip()
        rtype = recipient.Type  # 1=To, 2=CC, 3=BCC
        addresses.append({"address": addr, "name": name, "type": rtype})
    return addresses


def classify_email(recipients, target_name=None, target_email=None):
    """
    Tentukan apakah email ini 'direct' ke target atau bukan.
    Returns: 'Direct To', 'CC', 'BCC', 'Broadcast'
    """
    target_name = (target_name or config.TARGET_NAME).lower()
    target_email = (target_email or config.TARGET_EMAIL).lower()

    to_recipients = [r for r in recipients if r["type"] == 1]
    cc_recipients = [r for r in recipients if r["type"] == 2]

    for r in to_recipients:
        if target_email in r["address"] or target_name in r["name"].lower():
            if len(to_recipients) == 1:
                return "Direct To"
            else:
                return "To (Multiple)"

    for r in cc_recipients:
        if target_email in r["address"] or target_name in r["name"].lower():
            return "CC"

    if len(to_recipients) > config.BROADCAST_THRESHOLD:
        return "Broadcast"

    return "Other"


def fetch_emails_since(hours_back=None):
    """
    Ambil semua email dari Inbox dalam rentang waktu tertentu.
    Default: email sejak pukul 00:00 hari ini (atau config).
    Returns list of dict.
    """
    hours_back = hours_back or config.FETCH_HOURS_BACK
    since_dt = datetime.now() - timedelta(hours=hours_back)

    inbox = get_outlook_inbox()
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Terbaru dulu

    emails = []
    for msg in messages:
        try:
            received = msg.ReceivedTime
            # win32com mengembalikan datetime dengan timezone info
            received_naive = received.replace(tzinfo=None)
            if received_naive < since_dt:
                break  # sudah melewati rentang waktu, stop

            recipients = parse_recipients(msg.Recipients)
            to_names = "; ".join(
                r["name"] for r in recipients if r["type"] == 1
            )
            cc_names = "; ".join(
                r["name"] for r in recipients if r["type"] == 2
            )
            category = classify_email(recipients)

            body_preview = re.sub(r"\s+", " ", msg.Body or "").strip()
            body_preview = body_preview[:config.BODY_PREVIEW_LENGTH]

            emails.append({
                "received": received_naive,
                "subject": (msg.Subject or "").strip(),
                "sender": msg.SenderName or "",
                "sender_email": msg.SenderEmailAddress or "",
                "to": to_names,
                "cc": cc_names,
                "category": category,
                "body_preview": body_preview,
                "entry_id": msg.EntryID,
            })
        except Exception as e:
            # Skip email yang tidak bisa dibaca (e.g. meeting requests)
            print(f"  [skip] {e}")
            continue

    return emails


def get_or_create_workbook(filepath):
    """Buka Excel yang sudah ada atau buat baru."""
    if Path(filepath).exists():
        wb = openpyxl.load_workbook(filepath)
    else:
        wb = openpyxl.Workbook()
        # Hapus sheet default
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
    return wb


def get_today_sheet(wb):
    """Buat atau ambil sheet dengan nama tanggal hari ini."""
    sheet_name = datetime.now().strftime("%Y-%m-%d")
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)
        _setup_header(ws)
    return ws, sheet_name


def _setup_header(ws):
    """Tulis header row dengan styling."""
    headers = [
        "No", "Waktu Diterima", "Subject", "Dari", "Email Pengirim",
        "To", "CC", "Kategori", "Preview Isi",
    ]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        bottom=Side(style="medium", color="FFFFFF")
    )
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Lebar kolom
    col_widths = [5, 18, 45, 25, 30, 30, 25, 15, 60]
    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(col)
        ].width = w
    ws.row_dimensions[1].height = 22


def _row_fill(category):
    """Warna baris berdasarkan kategori."""
    palette = {
        "Direct To": "D6E4F0",
        "To (Multiple)": "EBF5FB",
        "CC": "FEF9E7",
        "BCC": "FDEDEC",
        "Broadcast": "F4ECF7",
        "Other": "F2F3F4",
    }
    return PatternFill("solid", fgColor=palette.get(category, "FFFFFF"))


def write_emails_to_sheet(ws, emails, start_row=None):
    """Tulis list email ke worksheet mulai dari baris kosong berikutnya."""
    if start_row is None:
        start_row = ws.max_row + 1
        if start_row == 2 and ws.cell(2, 1).value is None:
            start_row = 2  # sheet baru, mulai dari baris 2

    thin_border = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )

    written = 0
    for i, email in enumerate(emails, start=1):
        row = start_row + i - 1
        fill = _row_fill(email["category"])
        values = [
            i,
            email["received"].strftime("%H:%M  %d/%m/%Y"),
            email["subject"],
            email["sender"],
            email["sender_email"],
            email["to"],
            email["cc"],
            email["category"],
            email["body_preview"],
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = fill
            cell.border = thin_border
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=(col == 9),
            )
        ws.row_dimensions[row].height = 18
        written += 1

    return written


def run():
    """Entry point Task 1."""
    print("=" * 55)
    print("TASK 1 — Copy Email dari Outlook ke Excel")
    print("=" * 55)

    filepath = Path(config.EXCEL_PATH)
    filepath.parent.mkdir(parents=True, exist_ok=True)

    print(f"  Mengambil email {config.FETCH_HOURS_BACK} jam terakhir...")
    emails = fetch_emails_since()
    print(f"  Ditemukan {len(emails)} email baru.")

    if not emails:
        print("  Tidak ada email baru. Selesai.")
        return 0

    wb = get_or_create_workbook(filepath)
    ws, sheet_name = get_today_sheet(wb)
    written = write_emails_to_sheet(ws, emails)
    wb.save(filepath)

    print(f"  {written} email ditulis ke sheet '{sheet_name}'")
    print(f"  File disimpan: {filepath}")
    return written


if __name__ == "__main__":
    run()
