import csv
import json
from datetime import date
from pathlib import Path

from scripts.export_daily_json import export_daily_json


def _write_csv(path: Path) -> None:
    fields = [
        "唯一ID",
        "首次获取时间",
        "最后更新时间",
        "信息类型",
        "公司名称",
        "岗位",
        "详情链接",
    ]
    rows = [
        {
            "唯一ID": "new-1",
            "首次获取时间": "2026-09-09 08:00:00",
            "最后更新时间": "2026-09-09 08:00:00",
            "信息类型": "现场招聘",
            "公司名称": "甲公司",
            "岗位": "机械工程师",
            "详情链接": "https://example.com/a",
        },
        {
            "唯一ID": "updated-1",
            "首次获取时间": "2026-09-08 08:00:00",
            "最后更新时间": "2026-09-09 09:00:00",
            "信息类型": "需求信息",
            "公司名称": "乙公司",
            "岗位": "工艺工程师",
            "详情链接": "https://example.com/b",
        },
        {
            "唯一ID": "old-1",
            "首次获取时间": "2026-09-08 08:00:00",
            "最后更新时间": "2026-09-08 09:00:00",
            "信息类型": "需求信息",
            "公司名称": "丙公司",
            "岗位": "测试工程师",
            "详情链接": "https://example.com/c",
        },
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def test_export_daily_json_classifies_new_and_updated(tmp_path: Path) -> None:
    csv_path = tmp_path / "jobs.csv"
    reports_dir = tmp_path / "reports"
    _write_csv(csv_path)

    output = export_daily_json(
        csv_path,
        reports_dir,
        date(2026, 9, 9),
        "2026-09-09 10:00:00",
    )

    payload = json.loads(output.read_text(encoding="utf-8"))
    assert payload["date"] == "2026-09-09"
    assert payload["summary"] == {
        "total": 2,
        "new": 1,
        "updated": 1,
        "by_source": {"现场招聘": 1, "需求信息": 1},
    }
    assert {row["唯一ID"] for row in payload["records"]} == {"new-1", "updated-1"}
    assert {row["唯一ID"]: row["变更类型"] for row in payload["records"]} == {
        "new-1": "新增",
        "updated-1": "更新",
    }
    assert "甲公司" in output.read_text(encoding="utf-8")
