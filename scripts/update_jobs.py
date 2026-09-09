#!/usr/bin/env python3
"""Run the UESTC job monitor and export a structured JSON daily report."""

from __future__ import annotations

import argparse
import json
import logging
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

if __package__:
    from . import update_jobs_core as _core
    from .export_daily_json import export_daily_json
    from .update_jobs_core import *  # noqa: F401,F403 - preserve the historical module API
else:
    import update_jobs_core as _core
    from export_daily_json import export_daily_json
    from update_jobs_core import *  # noqa: F401,F403 - preserve the historical module API


def _config_path_from_argv() -> Path:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--config", type=Path, default=_core.DEFAULT_CONFIG_PATH)
    args, _unknown = parser.parse_known_args()
    return args.config.resolve()


def main() -> int:
    config_path = _config_path_from_argv()
    exit_code = _core.main()
    if exit_code != 0:
        return exit_code

    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
        timezone = ZoneInfo(str(config["timezone"]))
        now = datetime.now(timezone)
        output_path = export_daily_json(
            _core.CSV_PATH,
            _core.REPORTS_DIR,
            now.date(),
            now.strftime("%Y-%m-%d %H:%M:%S"),
        )
    except (OSError, ValueError, KeyError, json.JSONDecodeError) as exc:
        logging.error("JSON 日报生成失败：%s", exc)
        return 2

    print(f"JSON日报：{output_path.relative_to(_core.ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
