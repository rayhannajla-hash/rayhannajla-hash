"""
YouTube Channel Monetization Analyzer
Menganalisa channel YouTube dan memberikan rekomendasi untuk mempercepat monetisasi.
"""

import os
import sys
import json
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.progress import track
from rich import box

load_dotenv()

console = Console()

# Syarat monetisasi YouTube Partner Program (YPP)
YPP_SUBSCRIBERS   = 1_000
YPP_WATCH_HOURS   = 4_000   # jam, dalam 12 bulan terakhir
YPP_SHORTS_VIEWS  = 10_000_000  # views Shorts dalam 90 hari (alternatif)

SCOPES = ["https://www.googleapis.com/auth/yt-analytics.readonly"]


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def build_data_client():
    api_key = os.getenv("YOUTUBE_API_KEY")
    if not api_key:
        console.print("[red]ERROR:[/] Set YOUTUBE_API_KEY di file .env")
        sys.exit(1)
    return build("youtube", "v3", developerKey=api_key)


def build_analytics_client(client_secrets_file="client_secrets.json"):
    """OAuth2 untuk YouTube Analytics API (watch time, dll)."""
    creds = None
    token_file = "token.json"

    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(client_secrets_file):
                return None  # mode tanpa OAuth
            flow = InstalledAppFlow.from_client_secrets_file(client_secrets_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w") as f:
            f.write(creds.to_json())

    return build("youtubeAnalytics", "v2", credentials=creds)


# ---------------------------------------------------------------------------
# Data fetchers
# ---------------------------------------------------------------------------

def get_channel_stats(yt, channel_id):
    resp = yt.channels().list(
        part="snippet,statistics,contentDetails,brandingSettings",
        id=channel_id
    ).execute()

    if not resp.get("items"):
        console.print(f"[red]Channel ID tidak ditemukan:[/] {channel_id}")
        sys.exit(1)

    item = resp["items"][0]
    stats = item["statistics"]
    snippet = item["snippet"]

    return {
        "title":        snippet.get("title", ""),
        "description":  snippet.get("description", ""),
        "created_at":   snippet.get("publishedAt", ""),
        "country":      snippet.get("country", "N/A"),
        "subscribers":  int(stats.get("subscriberCount", 0)),
        "views":        int(stats.get("viewCount", 0)),
        "video_count":  int(stats.get("videoCount", 0)),
        "hidden_subs":  stats.get("hiddenSubscriberCount", False),
    }


def get_recent_videos(yt, channel_id, max_results=20):
    # Ambil uploads playlist ID
    ch = yt.channels().list(part="contentDetails", id=channel_id).execute()
    uploads_id = ch["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]

    items = []
    next_page = None
    while len(items) < max_results:
        resp = yt.playlistItems().list(
            part="snippet,contentDetails",
            playlistId=uploads_id,
            maxResults=min(50, max_results - len(items)),
            pageToken=next_page
        ).execute()
        items.extend(resp.get("items", []))
        next_page = resp.get("nextPageToken")
        if not next_page:
            break

    if not items:
        return []

    video_ids = [i["contentDetails"]["videoId"] for i in items]

    # Ambil stats per video
    stats_resp = yt.videos().list(
        part="statistics,contentDetails,snippet",
        id=",".join(video_ids)
    ).execute()

    videos = []
    for v in stats_resp.get("items", []):
        s = v.get("statistics", {})
        duration_iso = v["contentDetails"]["duration"]  # misal PT5M30S
        videos.append({
            "id":           v["id"],
            "title":        v["snippet"]["title"],
            "published_at": v["snippet"]["publishedAt"],
            "duration":     parse_duration(duration_iso),
            "views":        int(s.get("viewCount", 0)),
            "likes":        int(s.get("likeCount", 0)),
            "comments":     int(s.get("commentCount", 0)),
            "is_short":     parse_duration(duration_iso) <= 60,
        })

    return sorted(videos, key=lambda x: x["published_at"], reverse=True)


def get_watch_hours_analytics(analytics_client, channel_id):
    """Ambil total watch hours 12 bulan terakhir via Analytics API."""
    if not analytics_client:
        return None

    end_date   = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    start_date = (datetime.now(timezone.utc) - timedelta(days=365)).strftime("%Y-%m-%d")

    try:
        resp = analytics_client.reports().query(
            ids=f"channel=={channel_id}",
            startDate=start_date,
            endDate=end_date,
            metrics="estimatedMinutesWatched,views,averageViewDuration",
            dimensions="",
        ).execute()

        rows = resp.get("rows", [[0, 0, 0]])
        minutes_watched  = float(rows[0][0]) if rows else 0
        total_views      = int(rows[0][1]) if rows else 0
        avg_view_duration = float(rows[0][2]) if rows else 0

        return {
            "watch_hours":       minutes_watched / 60,
            "views_12m":         total_views,
            "avg_view_duration": avg_view_duration,
        }
    except HttpError:
        return None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_duration(iso_duration):
    """ISO 8601 duration -> detik. Contoh: PT1H5M30S -> 3930"""
    import re
    pattern = r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?"
    m = re.match(pattern, iso_duration)
    if not m:
        return 0
    h = int(m.group(1) or 0)
    mn = int(m.group(2) or 0)
    s = int(m.group(3) or 0)
    return h * 3600 + mn * 60 + s


def channel_age_days(created_at_str):
    created = datetime.fromisoformat(created_at_str.replace("Z", "+00:00"))
    return (datetime.now(timezone.utc) - created).days


def videos_last_n_days(videos, n):
    cutoff = datetime.now(timezone.utc) - timedelta(days=n)
    return [
        v for v in videos
        if datetime.fromisoformat(v["published_at"].replace("Z", "+00:00")) >= cutoff
    ]


# ---------------------------------------------------------------------------
# Display
# ---------------------------------------------------------------------------

def print_header(channel):
    console.print(Panel(
        f"[bold cyan]{channel['title']}[/]\n"
        f"Negara: {channel['country']}  |  "
        f"Dibuat: {channel['created_at'][:10]}  |  "
        f"Umur channel: {channel_age_days(channel['created_at'])} hari",
        title="[bold]YouTube Channel Analyzer[/]",
        border_style="cyan"
    ))


def print_ypp_status(channel, watch_hours):
    subs = channel["subscribers"]
    hours = watch_hours["watch_hours"] if watch_hours else None

    # Progress subscriber
    sub_pct   = min(subs / YPP_SUBSCRIBERS * 100, 100)
    hours_pct = min(hours / YPP_WATCH_HOURS * 100, 100) if hours else None

    sub_bar   = progress_bar(sub_pct)
    hours_bar = progress_bar(hours_pct) if hours_pct is not None else "[dim]perlu OAuth[/]"

    sub_status   = "[green]TERCAPAI[/]" if subs >= YPP_SUBSCRIBERS else f"[yellow]Kurang {YPP_SUBSCRIBERS - subs:,}[/]"
    hours_status = ""
    if hours is not None:
        hours_status = "[green]TERCAPAI[/]" if hours >= YPP_WATCH_HOURS else f"[yellow]Kurang {YPP_WATCH_HOURS - hours:,.0f} jam[/]"

    console.print("\n[bold underline]Syarat Monetisasi (YPP)[/]")

    t = Table(box=box.ROUNDED, show_header=True, header_style="bold magenta")
    t.add_column("Syarat",          style="bold", width=28)
    t.add_column("Target",          justify="right")
    t.add_column("Kamu",            justify="right")
    t.add_column("Progress",        width=22)
    t.add_column("Status")

    t.add_row(
        "Subscriber",
        f"{YPP_SUBSCRIBERS:,}",
        f"{subs:,}",
        sub_bar,
        sub_status,
    )
    t.add_row(
        "Jam Tayang (12 bln)",
        f"{YPP_WATCH_HOURS:,} jam",
        f"{hours:,.0f} jam" if hours is not None else "?",
        hours_bar,
        hours_status or "[dim]–[/]",
    )

    console.print(t)

    if hours is None:
        console.print(
            "[dim]  Tip: Tambahkan client_secrets.json untuk akses data jam tayang via OAuth.[/]"
        )


def progress_bar(pct, width=18):
    if pct is None:
        return "[dim]" + "─" * width + "[/]"
    filled = int(width * pct / 100)
    color  = "green" if pct >= 100 else "yellow" if pct >= 50 else "red"
    bar    = f"[{color}]{'█' * filled}[/]" + "░" * (width - filled)
    return f"{bar} {pct:.0f}%"


def print_video_table(videos):
    if not videos:
        console.print("[dim]Tidak ada video ditemukan.[/]")
        return

    console.print("\n[bold underline]Performa Video Terbaru[/]")
    t = Table(box=box.SIMPLE_HEAVY, show_header=True, header_style="bold blue")
    t.add_column("Judul",        max_width=38)
    t.add_column("Tipe",         justify="center", width=7)
    t.add_column("Tanggal",      width=12)
    t.add_column("Views",        justify="right")
    t.add_column("Likes",        justify="right")
    t.add_column("Komentar",     justify="right")

    for v in videos[:15]:
        tipe  = "[cyan]Short[/]" if v["is_short"] else "Video"
        date  = v["published_at"][:10]
        t.add_row(
            v["title"][:38],
            tipe,
            date,
            f"{v['views']:,}",
            f"{v['likes']:,}",
            f"{v['comments']:,}",
        )

    console.print(t)


def print_recommendations(channel, videos, watch_hours):
    subs          = channel["subscribers"]
    video_count   = channel["video_count"]
    hours         = watch_hours["watch_hours"] if watch_hours else None
    avg_dur       = watch_hours["avg_view_duration"] if watch_hours else None
    recent_7d     = videos_last_n_days(videos, 7)
    recent_30d    = videos_last_n_days(videos, 30)
    shorts_count  = sum(1 for v in videos if v["is_short"])
    long_count    = sum(1 for v in videos if not v["is_short"])

    console.print("\n[bold underline]Rekomendasi Monetisasi[/]\n")

    recs = []

    # --- Subscriber ---
    if subs < YPP_SUBSCRIBERS:
        kekurangan = YPP_SUBSCRIBERS - subs
        recs.append((
            "Subscriber",
            f"Masih kurang [bold]{kekurangan:,}[/] subscriber. "
            "Fokus pada call-to-action di akhir video dan community post."
        ))

    # --- Frekuensi upload ---
    if len(recent_7d) == 0:
        recs.append((
            "Konsistensi",
            "Tidak ada upload 7 hari terakhir. Upload minimal [bold]2–3x/minggu[/] "
            "agar algoritma merekomendasikan channel kamu."
        ))
    elif len(recent_30d) < 8:
        recs.append((
            "Konsistensi",
            f"Hanya [bold]{len(recent_30d)}[/] video dalam 30 hari. "
            "Target minimal [bold]8–12 video/bulan[/] untuk traksi optimal."
        ))

    # --- Watch time ---
    if hours is not None and hours < YPP_WATCH_HOURS:
        kekurangan_jam = YPP_WATCH_HOURS - hours
        # Estimasi hari lagi berdasarkan rata-rata 30 hari terakhir
        views_30d = sum(v["views"] for v in recent_30d)
        if avg_dur and views_30d > 0:
            hours_per_day = (views_30d * avg_dur / 60) / 30
            if hours_per_day > 0:
                hari_lagi = kekurangan_jam * 60 / (hours_per_day * 60)
                recs.append((
                    "Jam Tayang",
                    f"Butuh [bold]{kekurangan_jam:,.0f} jam[/] lagi. "
                    f"Dengan tren saat ini, estimasi [bold]{hari_lagi:.0f} hari[/] lagi. "
                    "Buat video berdurasi 8–15 menit untuk memaksimalkan jam tayang."
                ))
            else:
                recs.append((
                    "Jam Tayang",
                    f"Butuh [bold]{kekurangan_jam:,.0f} jam[/] lagi. "
                    "Tingkatkan views dan durasi rata-rata tonton."
                ))

    # --- Mix konten ---
    if shorts_count > long_count * 2:
        recs.append((
            "Jenis Konten",
            f"Terlalu banyak Shorts ([bold]{shorts_count}[/]) dibanding video panjang ([bold]{long_count}[/]). "
            "Shorts tidak menghitung jam tayang YPP. Tambah video panjang (>1 menit)."
        ))

    # --- Rata-rata views ---
    if videos:
        avg_views = sum(v["views"] for v in videos[:10]) / min(10, len(videos))
        if avg_views < 100:
            recs.append((
                "CTR & Thumbnail",
                f"Rata-rata views terlalu rendah ([bold]{avg_views:.0f}[/]/video). "
                "Perbaiki thumbnail (wajah + teks besar) dan judul (gunakan angka/emosi)."
            ))

    # --- Engagement ---
    if videos:
        total_views = sum(v["views"] for v in videos[:10])
        total_likes = sum(v["likes"] for v in videos[:10])
        if total_views > 0:
            like_rate = total_likes / total_views * 100
            if like_rate < 2:
                recs.append((
                    "Engagement",
                    f"Like rate rendah ([bold]{like_rate:.1f}%[/]). "
                    "Minta penonton like di momen spesifik video (jangan di awal)."
                ))

    # --- Deskripsi channel ---
    if len(channel.get("description", "")) < 100:
        recs.append((
            "SEO Channel",
            "Deskripsi channel terlalu pendek. Tambahkan kata kunci niche kamu "
            "agar lebih mudah ditemukan di pencarian YouTube."
        ))

    if not recs:
        console.print("[green]Channel kamu sudah di jalur yang benar! Pertahankan konsistensi.[/]")
        return

    for i, (kategori, pesan) in enumerate(recs, 1):
        console.print(
            f"  [bold cyan]{i}. {kategori}[/]\n"
            f"     {pesan}\n"
        )


def print_summary(channel, videos, watch_hours):
    hours = watch_hours["watch_hours"] if watch_hours else None
    subs  = channel["subscribers"]

    sub_done   = subs >= YPP_SUBSCRIBERS
    hours_done = hours >= YPP_WATCH_HOURS if hours is not None else None

    if sub_done and hours_done:
        status = "[bold green]SIAP DAFTAR MONETISASI![/]"
    elif sub_done or hours_done:
        status = "[bold yellow]HAMPIR SIAP – tinggal 1 syarat lagi[/]"
    else:
        status = "[bold red]BELUM MEMENUHI SYARAT – fokus subscriber & jam tayang[/]"

    console.print(Panel(status, title="Kesimpulan", border_style="blue"))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    channel_id = os.getenv("YOUTUBE_CHANNEL_ID")
    if not channel_id:
        console.print("[red]ERROR:[/] Set YOUTUBE_CHANNEL_ID di file .env")
        sys.exit(1)

    console.print("[dim]Menghubungkan ke YouTube API...[/]")
    yt        = build_data_client()
    analytics = build_analytics_client()

    console.print("[dim]Mengambil data channel...[/]")
    channel = get_channel_stats(yt, channel_id)

    console.print("[dim]Mengambil data video...[/]")
    videos = get_recent_videos(yt, channel_id, max_results=50)

    watch_hours = None
    if analytics:
        console.print("[dim]Mengambil data jam tayang (Analytics API)...[/]")
        watch_hours = get_watch_hours_analytics(analytics, channel_id)

    # Tampilkan hasil
    console.print()
    print_header(channel)
    print_ypp_status(channel, watch_hours)
    print_video_table(videos)
    print_recommendations(channel, videos, watch_hours)
    print_summary(channel, videos, watch_hours)

    # Simpan JSON untuk keperluan historis
    output = {
        "timestamp":   datetime.now(timezone.utc).isoformat(),
        "channel":     channel,
        "watch_hours": watch_hours,
        "videos":      videos[:20],
    }
    os.makedirs("reports", exist_ok=True)
    filename = f"reports/report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(filename, "w") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    console.print(f"\n[dim]Laporan disimpan: {filename}[/]")


if __name__ == "__main__":
    main()
