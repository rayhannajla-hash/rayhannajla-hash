"""
Konfigurasi Email Manager — sesuaikan sebelum digunakan.
"""

from pathlib import Path

# ── TARGET PENERIMA ──────────────────────────────────────────────
# Nama dan email yang dianggap "Direct To" (misal: Budhiarso)
TARGET_NAME = "Budhiarso"
TARGET_EMAIL = "budhiarso@"  # prefix email, atau isi lengkap misal: budhiarso@company.com

# ── PATH FILE EXCEL OUTPUT ───────────────────────────────────────
EXCEL_PATH = Path.home() / "Documents" / "EmailManager" / "email_log.xlsx"

# ── PENGATURAN FETCH ─────────────────────────────────────────────
# Ambil email dari berapa jam terakhir (default 24 jam)
FETCH_HOURS_BACK = 24

# Jika penerima To lebih dari angka ini, diklasifikasikan sebagai Broadcast
BROADCAST_THRESHOLD = 5

# Panjang preview isi email (karakter)
BODY_PREVIEW_LENGTH = 200

# ── TELEGRAM (OPSIONAL) ──────────────────────────────────────────
# Isi token dan chat_id untuk aktifkan pengiriman summary ke Telegram
# Kosongkan ("") untuk menonaktifkan
SEND_TELEGRAM = False
TELEGRAM_BOT_TOKEN = ""     # Contoh: "123456789:AABBccDDeeffGGhhIIjj..."
TELEGRAM_CHAT_ID = ""       # Contoh: "-1001234567890" atau "123456789"
