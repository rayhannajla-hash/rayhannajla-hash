# Setup YouTube Analyzer

## Langkah 1 – Buat API Key (wajib)

1. Buka https://console.cloud.google.com/
2. Buat project baru (misal: "youtube-analyzer")
3. Aktifkan **YouTube Data API v3**:
   - API & Services → Library → cari "YouTube Data API v3" → Enable
4. Buat credentials:
   - API & Services → Credentials → Create Credentials → **API Key**
   - Copy API key tersebut

## Langkah 2 – Cari Channel ID kamu

1. Buka channel YouTube kamu
2. Klik kanan → View Page Source
3. Cari `"channelId"` — nilainya format `UCxxxxxxxxxx`

**Atau** buka: `https://www.youtube.com/account_advanced` saat login

## Langkah 3 – Setup file .env

```bash
cp .env.example .env
```

Edit `.env`:
```
YOUTUBE_API_KEY=AIzaxxxxxxxxxxxxxxxxxxxx
YOUTUBE_CHANNEL_ID=UCxxxxxxxxxxxxxxxxxxxx
```

## Langkah 4 – Install & Jalankan

```bash
pip install -r requirements.txt
python analyzer.py
```

---

## (Opsional) Aktifkan Data Jam Tayang – YouTube Analytics API

Ini diperlukan untuk cek progress **4.000 jam tayang** secara akurat.

1. Di Google Cloud Console, aktifkan **YouTube Analytics API**
2. Buat OAuth credentials:
   - Credentials → Create Credentials → **OAuth client ID**
   - Application type: **Desktop app**
   - Download JSON → simpan sebagai `client_secrets.json` di folder ini
3. Jalankan ulang `python analyzer.py` → browser terbuka untuk login Google
4. Token disimpan otomatis di `token.json` (tidak perlu login lagi)

---

## Jalankan Rutin (Otomatis)

### Linux/Mac – crontab setiap hari jam 08:00
```bash
crontab -e
# Tambahkan:
0 8 * * * cd /path/to/youtube_analyzer && python analyzer.py >> logs/cron.log 2>&1
```

### Windows – Task Scheduler
Buat task yang menjalankan:
```
python C:\path\to\youtube_analyzer\analyzer.py
```

Laporan JSON tersimpan otomatis di folder `reports/`.
