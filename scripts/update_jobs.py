#!/usr/bin/env python3
"""Run the UESTC job monitor with fail-closed data-integrity guards."""

from __future__ import annotations

import argparse
import csv
import json
import logging
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

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


def _read_csv_ids(path: Path) -> list[str]:
    if not path.exists() or path.stat().st_size == 0:
        raise ValueError(f"招聘历史 CSV 不存在或为空：{path}")
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fields = list(reader.fieldnames or [])
        if "唯一ID" not in fields:
            raise ValueError("招聘历史 CSV 缺少“唯一ID”列")
        ids = [str(row.get("唯一ID") or "").strip() for row in reader]
    if not ids or not all(ids):
        raise ValueError("招聘历史 CSV 存在空唯一ID或没有记录")
    if len(ids) != len(set(ids)):
        raise ValueError("招聘历史 CSV 存在重复唯一ID")
    return ids


def _read_excel_ids(path: Path) -> list[str]:
    if not path.exists() or path.stat().st_size == 0:
        raise ValueError(f"招聘 Excel 不存在或为空：{path}")
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in workbook.worksheets:
            header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
            headers = [str(value).strip() if value is not None else "" for value in header_row]
            if "唯一ID" not in headers:
                continue
            id_index = headers.index("唯一ID")
            ids: list[str] = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                value = row[id_index] if id_index < len(row) else None
                text = str(value).strip() if value is not None else ""
                if text:
                    ids.append(text)
            if not ids:
                raise ValueError(f"Excel 工作表“{sheet.title}”没有招聘记录")
            if len(ids) != len(set(ids)):
                raise ValueError(f"Excel 工作表“{sheet.title}”存在重复唯一ID")
            return ids
    finally:
        workbook.close()
    raise ValueError("招聘 Excel 中找不到包含“唯一ID”的工作表")


def _validate_dataset(csv_path: Path, excel_path: Path) -> tuple[list[str], list[str]]:
    csv_ids = _read_csv_ids(csv_path)
    excel_ids = _read_excel_ids(excel_path)
    if len(csv_ids) != len(excel_ids):
        raise ValueError(f"CSV/Excel 记录数不一致：{len(csv_ids)} != {len(excel_ids)}")
    if set(csv_ids) != set(excel_ids):
        raise ValueError("CSV/Excel 唯一ID集合不一致")
    return csv_ids, excel_ids


def _ensure_history_preserved(before_ids: list[str], after_ids: list[str]) -> None:
    before = set(before_ids)
    after = set(after_ids)
    missing = before - after
    if missing:
        sample = ", ".join(sorted(missing)[:5])
        raise ValueError(
            f"检测到历史记录丢失：运行前 {len(before_ids)} 条，运行后 {len(after_ids)} 条，"
            f"缺失 {len(missing)} 个唯一ID；示例：{sample}"
        )
    if len(after_ids) < len(before_ids):
        raise ValueError(f"运行后总记录数减少：{len(before_ids)} -> {len(after_ids)}")


def main() -> int:
    config_path = _config_path_from_argv()

    try:
        baseline_ids, _ = _validate_dataset(_core.CSV_PATH, _core.EXCEL_PATH)
        print(f"运行前数据完整性校验通过：CSV/Excel 均为 {len(baseline_ids)} 条")
    except Exception as exc:
        logging.error("运行前数据完整性校验失败，已停止抓取以保护历史数据：%s", exc)
        return 3

    exit_code = _core.main()
    if exit_code != 0:
        return exit_code

    try:
        final_ids, _ = _validate_dataset(_core.CSV_PATH, _core.EXCEL_PATH)
        _ensure_history_preserved(baseline_ids, final_ids)

        config = json.loads(config_path.read_text(encoding="utf-8"))
        timezone = ZoneInfo(str(config["timezone"]))
        now = datetime.now(timezone)
        markdown_path = _core.REPORTS_DIR / f"{now.date().isoformat()}.md"
        if not markdown_path.exists() or markdown_path.stat().st_size == 0:
            raise ValueError(f"Markdown 日报不存在或为空：{markdown_path}")

        output_path = export_daily_json(
            _core.CSV_PATH,
            _core.REPORTS_DIR,
            now.date(),
            now.strftime("%Y-%m-%d %H:%M:%S"),
        )
        payload = json.loads(output_path.read_text(encoding="utf-8"))
        if payload.get("date") != now.date().isoformat():
            raise ValueError("JSON 日报日期与本次运行日期不一致")
        if not isinstance(payload.get("records"), list):
            raise ValueError("JSON 日报缺少 records 数组")
    except (OSError, ValueError, KeyError, json.JSONDecodeError) as exc:
        logging.error("运行后完整性校验失败；禁止将本次结果持久化到 GitHub：%s", exc)
        return 3

    print(f"运行后数据完整性校验通过：历史 {len(baseline_ids)} 条全部保留，当前 {len(final_ids)} 条")
    print(f"JSON日报：{output_path.relative_to(_core.ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
