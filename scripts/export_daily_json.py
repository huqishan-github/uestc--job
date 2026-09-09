#!/usr/bin/env python3
"""Export one day's changed recruitment records from jobs.csv as structured JSON."""

from __future__ import annotations

import csv
import json
import os
import tempfile
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any, Mapping, Sequence


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _date_part(value: Any) -> str:
    text = _text(value)
    return text[:10] if len(text) >= 10 else text


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


def build_daily_payload(
    rows: Sequence[Mapping[str, Any]],
    target_date: date,
    generated_at: str,
) -> dict[str, Any]:
    date_text = target_date.isoformat()
    records: list[dict[str, str]] = []
    for raw in rows:
        first_date = _date_part(raw.get("首次获取时间"))
        updated_date = _date_part(raw.get("最后更新时间"))
        if first_date == date_text:
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
            row.get("公司名称", ""),
            row.get("岗位", ""),
            row.get("唯一ID", ""),
        )
    )
    change_counts = Counter(row["变更类型"] for row in records)
    source_counts = Counter(row.get("信息类型", "") or "未分类" for row in records)
    return {
        "schema_version": 1,
        "date": date_text,
        "generated_at": generated_at,
        "source": "data/jobs.csv",
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
    descriptor, temp_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
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
        if not isinstance(validated, dict) or validated.get("schema_version") != 1:
            raise ValueError("JSON 日报临时文件校验失败")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def export_daily_json(
    csv_path: Path,
    reports_dir: Path,
    target_date: date,
    generated_at: str,
) -> Path:
    _fields, rows = read_jobs_csv(csv_path)
    payload = build_daily_payload(rows, target_date, generated_at)
    output_path = reports_dir / f"{target_date.isoformat()}.json"
    atomic_write_json(output_path, payload)
    return output_path
