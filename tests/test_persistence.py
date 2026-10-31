from __future__ import annotations

from pathlib import Path

import pytest

from scripts import update_jobs_core as core


def _record(unique_id: str, company: str) -> dict[str, str]:
    row = {field: "" for field in core.CSV_FIELDS}
    row.update(
        {
            "唯一ID": unique_id,
            "首次获取时间": "2026-10-31 09:00:00",
            "最后更新时间": "2026-10-31 09:00:00",
            "信息类型": "需求信息",
            "公司名称": company,
            "详情链接": f"https://jiuye.uestc.edu.cn/career/recruitment/online/{unique_id}",
            "链接状态": "verified",
        }
    )
    return row


def test_atomic_write_dataset_updates_both_stores(tmp_path: Path) -> None:
    csv_path = tmp_path / "jobs.csv"
    excel_path = tmp_path / "jobs.xlsx"
    rows = [_record("one", "甲公司"), _record("two", "乙公司")]

    core.atomic_write_dataset(csv_path, excel_path, rows, [], [])

    core.validate_data_consistency(csv_path, excel_path, 2)
    assert not list(tmp_path.glob(".*.backup"))


def test_atomic_write_dataset_rolls_back_both_stores_on_final_validation_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    csv_path = tmp_path / "jobs.csv"
    excel_path = tmp_path / "jobs.xlsx"
    original_rows = [_record("one", "甲公司")]
    core.atomic_write_csv(csv_path, original_rows)
    core.atomic_write_excel(excel_path, original_rows, [], [])
    original_csv = csv_path.read_bytes()
    original_excel = excel_path.read_bytes()

    real_validate = core.validate_data_consistency

    def fail_final_validation(
        candidate_csv: Path, candidate_excel: Path, expected_rows: int
    ) -> None:
        real_validate(candidate_csv, candidate_excel, expected_rows)
        if candidate_csv == csv_path and candidate_excel == excel_path:
            raise core.DataFileError("模拟目标文件校验失败")

    monkeypatch.setattr(core, "validate_data_consistency", fail_final_validation)

    with pytest.raises(core.DataFileError, match="已恢复原文件"):
        core.atomic_write_dataset(
            csv_path,
            excel_path,
            original_rows + [_record("two", "乙公司")],
            [],
            [],
        )

    assert csv_path.read_bytes() == original_csv
    assert excel_path.read_bytes() == original_excel
    real_validate(csv_path, excel_path, 1)
