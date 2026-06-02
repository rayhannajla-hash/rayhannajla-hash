"""
Setup Windows Task Scheduler untuk menjalankan run_daily.py otomatis setiap pagi.
Jalankan SEKALI dengan: python setup_scheduler.py

Tidak butuh admin — menggunakan /rl LIMITED (user-level).
"""

import subprocess
import sys
from pathlib import Path


def create_task(task_name, script_path, python_path, hour=7, minute=0):
    """Daftarkan task di Windows Task Scheduler via schtasks."""
    trigger_time = f"{hour:02d}:{minute:02d}"
    cmd = [
        "schtasks", "/create",
        "/tn", task_name,
        "/tr", f'"{python_path}" "{script_path}"',
        "/sc", "DAILY",
        "/st", trigger_time,
        "/f",           # overwrite jika sudah ada
        # Tanpa /rl HIGHEST — cukup user-level, tidak butuh admin
    ]
    print(f"  Mendaftarkan task : {task_name}")
    print(f"  Waktu             : setiap hari jam {trigger_time}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        print("  ✅ Task berhasil dibuat.")
    else:
        err = result.stderr.strip() or result.stdout.strip()
        print(f"  ❌ Gagal: {err}")
        print()
        print("  Alternatif manual:")
        print(f"    1. Buka Task Scheduler (cari di Start Menu)")
        print(f"    2. Create Basic Task → Daily → jam 07:00")
        print(f"    3. Action: Start a Program")
        print(f"       Program : {python_path}")
        print(f"       Argument: \"{script_path}\"")
    return result.returncode == 0


def main():
    script_dir = Path(__file__).parent.resolve()
    script_path = script_dir / "run_daily.py"
    python_path = sys.executable

    if not script_path.exists():
        print(f"ERROR: {script_path} tidak ditemukan.")
        sys.exit(1)

    print("=" * 55)
    print("SETUP WINDOWS TASK SCHEDULER")
    print("=" * 55)
    print(f"  Script : {script_path}")
    print(f"  Python : {python_path}")
    print()

    ok = create_task(
        task_name="OutlookEmailManager_DailySummary",
        script_path=script_path,
        python_path=python_path,
        hour=7,
        minute=0,
    )

    if ok:
        print()
        print("Verifikasi task:")
        subprocess.run(
            ["schtasks", "/query", "/tn", "OutlookEmailManager_DailySummary", "/fo", "LIST"],
            text=True,
        )


if __name__ == "__main__":
    main()
