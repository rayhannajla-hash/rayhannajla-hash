"""
Task 3: Buat summary email setiap pagi
Generate sheet Summary dengan statistik, topik utama, dan action items.
Opsional: kirim ringkasan ke Telegram via bot AINUN.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from datetime import datetime
from pathlib import Path
from collections import Counter
import re
import config

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False


# Kata-kata yang sering menandai action item di subject/body
ACTION_KEYWORDS = [
    "mohon", "tolong", "please", "request", "urgent", "segera",
    "deadline", "due", "approval", "approve", "follow up", "reminder",
    "konfirmasi", "confirm", "review", "check", "update",
]


def read_sorted_data(wb):
    """Baca sheet Sorted hari ini, return emails grouped by category."""
    sheet_name = "Sorted_" + datetime.now().strftime("%Y-%m-%d")
    today_sheet = datetime.now().strftime("%Y-%m-%d")

    # Fallback ke sheet raw jika sorted belum ada
    source = sheet_name if sheet_name in wb.sheetnames else today_sheet
    if source not in wb.sheetnames:
        raise ValueError(
            f"Tidak ada data untuk hari ini. "
            "Pastikan Task 1 & 2 sudah dijalankan."
        )

    ws = wb[source]
    headers = None
    emails = []
    for row in ws.iter_rows(values_only=True):
        if row[0] in (None, "No", "") and headers is None:
            if "Subject" in (row[2] if row[2] else ""):
                headers = list(row)
            continue
        if headers and row[0] not in (None, ""):
            try:
                int(row[0])  # pastikan kolom No berisi angka
                emails.append(dict(zip(headers, row)))
            except (TypeError, ValueError):
                continue
    return emails


def extract_action_items(emails):
    """Cari email yang kemungkinan butuh tindak lanjut."""
    items = []
    for email in emails:
        subject = (email.get("Subject") or "").lower()
        preview = (email.get("Preview Isi") or "").lower()
        text = subject + " " + preview
        matched = [kw for kw in ACTION_KEYWORDS if kw in text]
        if matched:
            items.append({
                "subject": email.get("Subject", ""),
                "dari": email.get("Dari", ""),
                "kategori": email.get("Kategori", ""),
                "keyword": ", ".join(matched[:3]),
            })
    return items


def top_senders(emails, n=5):
    """Return N pengirim terbanyak."""
    counter = Counter(
        email.get("Dari", "Unknown") for email in emails
    )
    return counter.most_common(n)


def keyword_topics(emails, n=8):
    """Ekstrak kata kunci terbanyak dari subject email."""
    stopwords = {
        "re", "fwd", "fw", "dari", "untuk", "dan", "atau", "yang", "dengan",
        "di", "ke", "the", "and", "or", "for", "is", "in", "on", "to",
        "a", "of", "by", "an", "be", "at", "as", "it", "we", "all",
    }
    word_counter = Counter()
    for email in emails:
        subject = email.get("Subject") or ""
        words = re.findall(r"[a-zA-Z]{3,}", subject.lower())
        word_counter.update(w for w in words if w not in stopwords)
    return word_counter.most_common(n)


def _write_section_title(ws, row, col, text, bg="1F4E79"):
    ws.merge_cells(
        start_row=row, start_column=col,
        end_row=row, end_column=col + 4
    )
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22
    return row + 1


def _kv_row(ws, row, col, key, value, alt=False):
    """Tulis pasangan key-value dengan warna alternating."""
    fill_color = "EAF2FF" if alt else "FFFFFF"
    fill = PatternFill("solid", fgColor=fill_color)
    thin = Border(
        left=Side(style="thin", color="BDC3C7"),
        right=Side(style="thin", color="BDC3C7"),
        top=Side(style="thin", color="BDC3C7"),
        bottom=Side(style="thin", color="BDC3C7"),
    )
    key_cell = ws.cell(row=row, column=col, value=key)
    key_cell.font = Font(bold=True, size=10)
    key_cell.fill = fill
    key_cell.border = thin
    key_cell.alignment = Alignment(vertical="center", indent=1)

    val_cell = ws.cell(row=row, column=col + 1, value=value)
    val_cell.fill = fill
    val_cell.border = thin
    val_cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 17


def build_summary_sheet(wb, emails):
    """Buat sheet Summary_<tanggal> dengan statistik lengkap."""
    sheet_name = "Summary_" + datetime.now().strftime("%Y-%m-%d")
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 0)  # taruh di depan

    # Lebar kolom
    for col, w in zip("ABCDEFGHIJ", [22, 28, 5, 22, 28, 5, 22, 28, 5, 22]):
        ws.column_dimensions[col].width = w

    now = datetime.now()
    today_str = now.strftime("%A, %d %B %Y")

    # ── JUDUL UTAMA ──────────────────────────────────────────────
    ws.merge_cells("A1:J1")
    title = ws["A1"]
    title.value = f"DAILY EMAIL SUMMARY  |  {today_str}"
    title.font = Font(bold=True, color="FFFFFF", size=14)
    title.fill = PatternFill("solid", fgColor="0B3954")
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:J2")
    sub = ws["A2"]
    sub.value = (
        f"Target: {config.TARGET_NAME}   |   "
        f"Generated: {now.strftime('%H:%M:%S')}   |   "
        f"Total Email: {len(emails)}"
    )
    sub.font = Font(italic=True, color="5D6D7E", size=10)
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # ── SECTION 1: STATISTIK PER KATEGORI (kiri) ─────────────────
    from email_sorter import CATEGORY_ORDER, CATEGORY_COLORS
    categories = {cat: 0 for cat in CATEGORY_ORDER}
    for e in emails:
        cat = e.get("Kategori", "Other")
        if cat in categories:
            categories[cat] += 1

    row = 4
    row = _write_section_title(ws, row, 1, "STATISTIK PER KATEGORI", "154360")
    for i, (cat, count) in enumerate(categories.items()):
        _kv_row(ws, row, 1, cat, count, alt=(i % 2 == 0))
        row += 1
    _kv_row(ws, row, 1, "TOTAL", len(emails), alt=False)
    ws.cell(row, 1).font = Font(bold=True, size=10, color="FFFFFF")
    ws.cell(row, 1).fill = PatternFill("solid", fgColor="1A5276")
    ws.cell(row, 2).font = Font(bold=True, size=10, color="FFFFFF")
    ws.cell(row, 2).fill = PatternFill("solid", fgColor="1A5276")
    row += 2

    # ── SECTION 2: TOP SENDERS ────────────────────────────────────
    row = _write_section_title(ws, row, 1, "TOP PENGIRIM", "6E2FA0")
    for i, (sender, count) in enumerate(top_senders(emails)):
        _kv_row(ws, row, 1, sender, f"{count} email", alt=(i % 2 == 0))
        row += 1
    row += 1

    # ── SECTION 3: TOPIK POPULER (tengah) ────────────────────────
    topic_start_row = 4
    row2 = topic_start_row
    row2 = _write_section_title(ws, row2, 4, "KATA KUNCI TOPIK", "0E6655")
    for i, (word, count) in enumerate(keyword_topics(emails)):
        _kv_row(ws, row2, 4, word.title(), f"{count}x", alt=(i % 2 == 0))
        row2 += 1
    row2 += 2

    # ── SECTION 4: ACTION ITEMS ───────────────────────────────────
    action_items = extract_action_items(emails)
    row2 = _write_section_title(
        ws, row2, 4, f"ACTION ITEMS  ({len(action_items)} email)", "922B21"
    )
    if action_items:
        for i, item in enumerate(action_items[:10]):  # max 10
            label = f"{item['dari']} | {item['keyword']}"
            _kv_row(ws, row2, 4, item["subject"][:40], label, alt=(i % 2 == 0))
            row2 += 1
    else:
        ws.cell(row2, 4, "Tidak ada action items terdeteksi")
        ws.cell(row2, 4).font = Font(italic=True, color="7F8C8D")
        row2 += 1

    # ── SECTION 5: DIRECT EMAIL DETAIL (kanan) ───────────────────
    direct = [e for e in emails if e.get("Kategori") == "Direct To"]
    row3 = 4
    row3 = _write_section_title(
        ws, row3, 7, f"DIRECT TO {config.TARGET_NAME.upper()} ({len(direct)})", "1A5276"
    )
    if direct:
        for i, e in enumerate(direct[:15]):  # max 15
            _kv_row(
                ws, row3, 7,
                e.get("Dari", ""),
                (e.get("Subject") or "")[:35],
                alt=(i % 2 == 0)
            )
            row3 += 1
    else:
        ws.cell(row3, 7, "Tidak ada email direct hari ini")
        ws.cell(row3, 7).font = Font(italic=True, color="7F8C8D")

    return ws, sheet_name


def format_telegram_message(emails, categories):
    """Buat pesan ringkas untuk dikirim ke Telegram."""
    now = datetime.now()
    direct = categories.get("Direct To", 0)
    action_count = len(extract_action_items(emails))

    lines = [
        f"☀️ *DAILY EMAIL SUMMARY*",
        f"📅 {now.strftime('%A, %d %B %Y')}",
        f"🕐 Generated: {now.strftime('%H:%M')}",
        "",
        f"📧 *Total Email:* {len(emails)}",
        f"🎯 *Direct to {config.TARGET_NAME}:* {direct}",
        f"⚡ *Action Items:* {action_count}",
        "",
        "*Per Kategori:*",
    ]
    for cat, count in categories.items():
        if count > 0:
            icon = {
                "Direct To": "🔵", "To (Multiple)": "🟦",
                "CC": "🟡", "BCC": "🔴",
                "Broadcast": "🟣", "Other": "⚪",
            }.get(cat, "▪️")
            lines.append(f"{icon} {cat}: {count}")

    action_items = extract_action_items(emails)
    if action_items:
        lines.append("")
        lines.append("*⚠️ Perlu Perhatian:*")
        for item in action_items[:5]:
            lines.append(f"• {item['subject'][:50]}")

    return "\n".join(lines)


def send_to_telegram(message):
    """Kirim pesan ke Telegram bot (opsional)."""
    if not REQUESTS_AVAILABLE:
        print("  [Telegram] requests tidak terinstall, skip.")
        return False
    if not config.TELEGRAM_BOT_TOKEN or not config.TELEGRAM_CHAT_ID:
        print("  [Telegram] Token/Chat ID belum dikonfigurasi di config.py, skip.")
        return False
    try:
        url = f"https://api.telegram.org/bot{config.TELEGRAM_BOT_TOKEN}/sendMessage"
        resp = requests.post(url, json={
            "chat_id": config.TELEGRAM_CHAT_ID,
            "text": message,
            "parse_mode": "Markdown",
        }, timeout=10)
        if resp.status_code == 200:
            print("  [Telegram] Pesan terkirim.")
            return True
        else:
            print(f"  [Telegram] Gagal: {resp.status_code} {resp.text[:100]}")
            return False
    except Exception as e:
        print(f"  [Telegram] Error: {e}")
        return False


def run():
    """Entry point Task 3."""
    print("=" * 55)
    print("TASK 3 — Generate Summary Email Harian")
    print("=" * 55)

    filepath = Path(config.EXCEL_PATH)
    if not filepath.exists():
        print(f"  File tidak ditemukan: {filepath}")
        return

    wb = openpyxl.load_workbook(filepath)
    emails = read_sorted_data(wb)
    print(f"  Memproses {len(emails)} email...")

    ws, sheet_name = build_summary_sheet(wb, emails)
    wb.save(filepath)
    print(f"  Sheet '{sheet_name}' berhasil dibuat.")
    print(f"  File disimpan: {filepath}")

    # Hitung statistik untuk Telegram
    from email_sorter import CATEGORY_ORDER
    categories = {cat: 0 for cat in CATEGORY_ORDER}
    for e in emails:
        cat = e.get("Kategori", "Other")
        if cat in categories:
            categories[cat] += 1

    if config.SEND_TELEGRAM:
        msg = format_telegram_message(emails, categories)
        print("\n--- PREVIEW TELEGRAM MESSAGE ---")
        print(msg)
        print("--------------------------------")
        send_to_telegram(msg)


if __name__ == "__main__":
    run()
