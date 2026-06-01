"""
Orchestrator — jalankan semua 3 task secara berurutan.
Gunakan ini sebagai entry point scheduler (Task Scheduler / cron).

  Cara pakai manual : python run_daily.py
  Cara pakai Task 1 saja: python run_daily.py --task 1
  Cara pakai Task 2+3 : python run_daily.py --task 2 3
"""

import sys
import argparse
from datetime import datetime


def main():
    parser = argparse.ArgumentParser(
        description="Outlook Email Manager — Daily Runner"
    )
    parser.add_argument(
        "--task",
        nargs="+",
        type=int,
        choices=[1, 2, 3],
        help="Task yang dijalankan (default: semua)"
    )
    args = parser.parse_args()
    tasks = args.task or [1, 2, 3]

    print("╔" + "═" * 53 + "╗")
    print("║  OUTLOOK EMAIL MANAGER — DAILY RUN".center(55) + "║")
    print(f"║  {datetime.now().strftime('%A, %d %B %Y  %H:%M:%S')}".center(55) + "║")
    print("╚" + "═" * 53 + "╝")
    print()

    results = {}

    if 1 in tasks:
        try:
            import email_fetcher
            count = email_fetcher.run()
            results[1] = ("OK", f"{count} email diimport")
        except Exception as e:
            results[1] = ("ERROR", str(e))
        print()

    if 2 in tasks:
        try:
            import email_sorter
            email_sorter.run()
            results[2] = ("OK", "Sorting selesai")
        except Exception as e:
            results[2] = ("ERROR", str(e))
        print()

    if 3 in tasks:
        try:
            import summary_generator
            summary_generator.run()
            results[3] = ("OK", "Summary selesai")
        except Exception as e:
            results[3] = ("ERROR", str(e))
        print()

    # ── Laporan akhir ──────────────────────────────────────────────
    print()
    print("─" * 55)
    print("HASIL:")
    labels = {1: "Copy Email (Task 1)", 2: "Sort Email (Task 2)", 3: "Summary (Task 3)"}
    for task_no, (status, msg) in results.items():
        icon = "✅" if status == "OK" else "❌"
        print(f"  {icon} {labels[task_no]}: {msg}")
    print("─" * 55)

    has_error = any(s == "ERROR" for s, _ in results.values())
    sys.exit(1 if has_error else 0)


if __name__ == "__main__":
    main()
