#!/usr/bin/env python3
"""Backfill structured daily JSON reports from the authoritative jobs.csv history."""

from __future__ import annotations

import argparse
import json
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import update_jobs_core as core
from export_daily_json import export_daily_json


def parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"无效日期 {value!r}，请使用 YYYY-MM-DD") from exc


def iter_dates(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def main() -> int:
    config = json.loads(core.DEFAULT_CONFIG_PATH.read_text(encoding="utf-8"))
    timezone = ZoneInfo(str(config["timezone"]))
    today = datetime.now(timezone).date()

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start", type=parse_date, default=date(2026, 9, 3))
    parser.add_argument("--end", type=parse_date, default=today)
    args = parser.parse_args()
    if args.start > args.end:
        parser.error("--start 不能晚于 --end")

    generated_at = datetime.now(timezone).strftime("%Y-%m-%d %H:%M:%S")
    created: list[Path] = []
    for target_date in iter_dates(args.start, args.end):
        path = export_daily_json(
            core.CSV_PATH,
            core.REPORTS_DIR,
            target_date,
            generated_at,
        )
        created.append(path)
        print(f"已生成：{path.relative_to(core.ROOT)}")

    print(f"共生成/刷新 {len(created)} 个 JSON 日报。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
