#!/usr/bin/env python3
"""Export daily recruitment changes, with an optional first-day rolling-window snapshot."""

from __future__ import annotations

import csv
import json
import os
import tempfile
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Mapping, Sequence


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _date_part(value: Any) -> str:
    text = _text(value)
    return text[:10] if len(text) >= 10 else text


def _as_date(value: Any) -> date | None:
    text = _date_part(value)
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def read_jobs_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    if not path.exists() or path.stat().st_size == 0:
        raise FileNotFoundError(f"招聘历史 CSV 不存在或为空：{path}")
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                fields = list(reader.fieldnames or [])
                if "唯一ID" not in fields:
                    raise ValueError("招聘历史 CSV 缺少“唯一ID”列")
                rows = [
                    {field: _text(row.get(field)) for field in fields}
                    for row in reader
                    if _text(row.get("唯一ID"))
                ]
                return fields, rows
        except UnicodeError as exc:
            last_error = exc
    raise UnicodeError(f"无法读取招聘历史 CSV：{last_error}")


def _in_forward_window(row: Mapping[str, Any], target_date: date, forward_days: int) -> bool:
    """Return True when an onsite/recruitment date falls in [today, today+N]."""
    end_date = target_date + timedelta(days=forward_days)
    for field in ("宣讲日期", "招聘日期"):
        value = _as_date(row.get(field))
        if value is not None and target_date <= value <= end_date:
            return True
    return False


def build_daily_payload(
    rows: Sequence[Mapping[str, Any]],
    target_date: date,
    generated_at: str,
    *,
    bootstrap_date: date | None = None,
    forward_days: int = 30,
) -> dict[str, Any]:
    """Build a daily report.

    On bootstrap_date, include every historical/current record whose recruitment or
    onsite date is inside the forward rolling window. This creates the initial
    snapshot. On later days, include only records first discovered or materially
    updated that day, so previously reported records are not repeated.
    """
    date_text = target_date.isoformat()
    is_bootstrap = bootstrap_date == target_date
    records: list[dict[str, str]] = []

    for raw in rows:
        first_date = _date_part(raw.get("首次获取时间"))
        updated_date = _date_part(raw.get("最后更新时间"))
        in_window = _in_forward_window(raw, target_date, forward_days)

        if is_bootstrap and in_window:
            change_type = "新增"
        elif first_date == date_text:
            change_type = "新增"
        elif updated_date == date_text:
            change_type = "更新"
        else:
            continue

        record = {str(key): _text(value) for key, value in raw.items()}
        record = {"变更类型": change_type, **record}
        records.append(record)

    records.sort(
        key=lambda row: (
            row.get("信息类型", ""),
            row.get("宣讲日期", "") or row.get("招聘日期", ""),
            row.get("公司名称", ""),
            row.get("岗位", ""),
            row.get("唯一ID", ""),
        )
    )
    change_counts = Counter(row["变更类型"] for row in records)
    source_counts = Counter(row.get("信息类型", "") or "未分类" for row in records)
    return {
        "schema_version": 2,
        "date": date_text,
        "generated_at": generated_at,
        "source": "data/jobs.csv",
        "report_mode": "rolling_window_bootstrap" if is_bootstrap else "daily_increment",
        "window": {
            "start": date_text,
            "end": (target_date + timedelta(days=forward_days)).isoformat(),
            "forward_days": forward_days,
        },
        "summary": {
            "total": len(records),
            "new": change_counts.get("新增", 0),
            "updated": change_counts.get("更新", 0),
            "by_source": dict(sorted(source_counts.items())),
        },
        "records": records,
    }


def atomic_write_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    os.close(descriptor)
    temporary = Path(temp_name)
    try:
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        with temporary.open("r", encoding="utf-8") as handle:
            validated = json.load(handle)
        if not isinstance(validated, dict) or validated.get("schema_version") not in {1, 2}:
            raise ValueError("JSON 日报临时文件校验失败")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def export_daily_json(
    csv_path: Path,
    reports_dir: Path,
    target_date: date,
    generated_at: str,
    *,
    bootstrap_date: date | None = None,
    forward_days: int = 30,
) -> Path:
    _fields, rows = read_jobs_csv(csv_path)
    payload = build_daily_payload(
        rows,
        target_date,
        generated_at,
        bootstrap_date=bootstrap_date,
        forward_days=forward_days,
    )
    output_path = reports_dir / f"{target_date.isoformat()}.json"
    atomic_write_json(output_path, payload)
    return output_path


def export_daily_markdown(payload: Mapping[str, Any], reports_dir: Path, target_date: date) -> Path:
    """Write Markdown from the same payload so MD and JSON always contain the same records."""
    summary = payload.get("summary", {})
    window = payload.get("window", {})
    mode = payload.get("report_mode", "daily_increment")
    lines = [
        f"# 电子科技大学招聘监控日报 {target_date.isoformat()}",
        "",
        f"- 生成时间：{_text(payload.get('generated_at'))}",
        f"- 扫描窗口：{_text(window.get('start'))} ～ {_text(window.get('end'))}",
        f"- 报告模式：{'首日30天全量快照' if mode == 'rolling_window_bootstrap' else '滚动30天新增/更新'}",
        f"- 新增：{summary.get('new', 0)}",
        f"- 更新：{summary.get('updated', 0)}",
        f"- 合计：{summary.get('total', 0)}",
        "",
    ]
    records = payload.get("records", [])
    if not records:
        lines.append("本日未发现新增或更新记录。")
    else:
        for index, row in enumerate(records, 1):
            title = _text(row.get("公司名称")) or "未明确公司"
            job = _text(row.get("岗位")) or "未明确岗位"
            lines.extend([
                f"## {index}. [{_text(row.get('变更类型'))}] {title}｜{job}",
                "",
                f"- 信息类型：{_text(row.get('信息类型'))}",
                f"- 发布时间：{_text(row.get('发布时间'))}",
                f"- 招聘日期：{_text(row.get('招聘日期'))}",
                f"- 宣讲日期：{_text(row.get('宣讲日期'))}",
                f"- 工作地点：{_text(row.get('工作地点'))}",
                f"- 学历要求：{_text(row.get('学历要求'))}",
                f"- 专业要求：{_text(row.get('专业要求'))}",
                f"- 网申截止时间：{_text(row.get('网申截止时间'))}",
                f"- 详情链接：{_text(row.get('详情链接'))}",
                "",
            ])
    output_path = reports_dir / f"{target_date.isoformat()}.md"
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return output_path
