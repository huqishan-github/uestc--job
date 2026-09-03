from __future__ import annotations

from copy import deepcopy

from scripts.update_jobs import (
    ONSITE_ROUTE_SLUGS,
    Candidate,
    candidate_detail_url,
    extract_positions,
    extract_submission,
    make_unique_id,
    merge_scraped_records,
    normalize_text,
    normalize_url,
)


def sample_record(**overrides: str) -> dict[str, str]:
    record = {
        "唯一ID": "",
        "首次获取时间": "",
        "最后更新时间": "",
        "信息类型": "需求信息",
        "公司名称": "示例科技有限公司",
        "招聘标题": "2027 届校园招聘",
        "岗位": "机械设计工程师",
        "发布时间": "2026-09-03 10:00:00",
        "招聘日期": "2026-09-03",
        "工作地点": "成都",
        "宣讲日期": "",
        "宣讲开始时间": "",
        "宣讲结束时间": "",
        "宣讲地点": "",
        "学历要求": "本科",
        "专业要求": "机械工程",
        "网申截止时间": "",
        "投递方式": "",
        "投递链接": "",
        "详情链接": "https://jiuye.uestc.edu.cn/career/recruitment/online/abc",
        "链接状态": "verified",
        "来源页面": "https://jiuye.uestc.edu.cn/career/recruitment/online",
        "备注": "",
    }
    record.update(overrides)
    record["唯一ID"] = make_unique_id(record)
    return record


def test_text_normalization() -> None:
    assert normalize_text("  机械\n\t设计　工程师  ") == "机械 设计 工程师"
    assert normalize_text("ＡＢＣ") == "ABC"


def test_url_normalization_removes_tracking_and_trailing_slash() -> None:
    url = "HTTPS://JIUYE.UESTC.EDU.CN/career/item/?utm_source=test&id=7#top"
    assert normalize_url(url) == "https://jiuye.uestc.edu.cn/career/item?id=7"


def test_unique_id_is_stable_under_normalization() -> None:
    left = sample_record()
    right = sample_record(
        公司名称="  示例科技有限公司\n",
        岗位="机械设计工程师  ",
        详情链接="https://jiuye.uestc.edu.cn/career/recruitment/online/abc/",
    )
    assert make_unique_id(left) == make_unique_id(right)


def test_utm_url_and_plain_url_have_same_id() -> None:
    plain = sample_record()
    tracked = sample_record(
        详情链接="https://jiuye.uestc.edu.cn/career/recruitment/online/abc?utm_source=automation"
    )
    assert make_unique_id(plain) == make_unique_id(tracked)


def test_identical_record_is_not_added_twice() -> None:
    record = sample_record()
    first = merge_scraped_records([], [record], "2026-09-03 10:00:00")
    second = merge_scraped_records(first.rows, [record], "2026-09-04 10:00:00")
    assert len(second.rows) == 1
    assert not second.new_records
    assert not second.updated_records
    assert second.unchanged == 1


def test_same_company_different_jobs_do_not_deduplicate() -> None:
    mechanical = sample_record(岗位="机械设计工程师")
    structural = sample_record(岗位="结构设计工程师")
    result = merge_scraped_records([], [mechanical, structural], "2026-09-03 10:00:00")
    assert len(result.rows) == 2
    assert mechanical["唯一ID"] != structural["唯一ID"]


def test_same_detail_url_different_jobs_get_different_ids() -> None:
    first = sample_record(岗位="CAE工程师")
    second = sample_record(岗位="光机结构工程师")
    assert make_unique_id(first) != make_unique_id(second)


def test_existing_record_update_does_not_add_row() -> None:
    original = sample_record(网申截止时间="", 链接状态="unverified")
    initialized = merge_scraped_records([], [original], "2026-09-03 10:00:00")
    incoming = deepcopy(original)
    incoming["网申截止时间"] = "2026-10-15 23:59:59"
    incoming["链接状态"] = "verified"
    updated = merge_scraped_records(initialized.rows, [incoming], "2026-09-04 10:00:00")
    assert len(updated.rows) == 1
    assert not updated.new_records
    assert len(updated.updated_records) == 1
    assert updated.rows[0]["网申截止时间"] == "2026-10-15 23:59:59"
    assert updated.rows[0]["链接状态"] == "verified"
    assert updated.rows[0]["首次获取时间"] == "2026-09-03 10:00:00"


def test_transient_blank_or_failed_status_does_not_overwrite_good_data() -> None:
    original = sample_record(专业要求="机械工程", 链接状态="verified")
    initialized = merge_scraped_records([], [original], "2026-09-03 10:00:00")
    incoming = deepcopy(original)
    incoming["专业要求"] = ""
    incoming["链接状态"] = "failed"
    result = merge_scraped_records(initialized.rows, [incoming], "2026-09-04 10:00:00")
    assert not result.updated_records
    assert result.rows[0]["专业要求"] == "机械工程"
    assert result.rows[0]["链接状态"] == "verified"


def test_explicit_table_jobs_are_split_with_row_specific_requirements() -> None:
    detail = {
        "title": "",
        "recruitmentContent": """
        <table>
          <tr><th>职位类别</th><th>学历要求</th><th>专业要求</th><th>工作地点</th></tr>
          <tr><td>机械设计工程师</td><td>硕士</td><td>机械工程</td><td>成都</td></tr>
          <tr><td>CAE工程师</td><td>本科及以上</td><td>力学</td><td>绵阳</td></tr>
        </table>
        """,
    }
    positions = extract_positions(detail)
    assert [position.name for position in positions] == ["机械设计工程师", "CAE工程师"]
    assert positions[0].education == "硕士"
    assert positions[1].major == "力学"


def test_plain_text_application_url_stops_before_chinese_description() -> None:
    content = "<p>简历投递</p><p>登录招聘网站：</p><p>http://career.example.com，线上投递简历</p>"
    method, url = extract_submission({}, content)
    assert method == ""
    assert url == "http://career.example.com/"


def test_rowspan_keeps_following_rows_in_the_correct_job_column() -> None:
    detail = {
        "title": "",
        "recruitmentContent": """
        <table>
          <tr><th>岗位分类</th><th>岗位名称</th><th>需求专业</th></tr>
          <tr><td rowspan="2">技术类</td><td>机械工程师</td><td>机械工程</td></tr>
          <tr><td>电气工程师</td><td>电气工程</td></tr>
        </table>
        """,
    }
    positions = extract_positions(detail)
    assert [(position.name, position.major) for position in positions] == [
        ("机械工程师", "机械工程"),
        ("电气工程师", "电气工程"),
    ]


def test_rowspan_job_is_repeated_in_description_only_followup_row() -> None:
    detail = {
        "title": "",
        "recruitmentContent": """
        <table>
          <tr><th>岗位</th><th>需求专业及职责</th></tr>
          <tr><td rowspan="2">运行岗</td><td>机械、电气、自动化</td></tr>
          <tr><td>负责设备运行监控。</td></tr>
        </table>
        """,
    }
    positions = extract_positions(detail)
    assert [position.name for position in positions] == ["运行岗"]


def test_live_site_recruitment_types_use_their_own_route_slugs() -> None:
    assert ONSITE_ROUTE_SLUGS == {
        "ONSITE": "onsite",
        "GROUP": "group",
        "RECRUITMENT": "aerial",
    }


def test_candidate_detail_url_uses_real_type_and_base36_id() -> None:
    candidate = Candidate(
        source_type="现场招聘",
        source_code="GROUP",
        source_slug="group",
        source_page="https://jiuye.uestc.edu.cn/career/recruitment/group",
        raw={"id": "36"},
    )
    assert candidate_detail_url("https://jiuye.uestc.edu.cn/career/", candidate) == (
        "https://jiuye.uestc.edu.cn/career/recruitment/group/10"
    )


def test_explicit_line_breaks_and_long_lists_split_without_splitting_compound_job() -> None:
    detail = {
        "title": "",
        "recruitmentContent": """
        <table>
          <tr><th>岗位名称</th></tr>
          <tr><td>软件工程师<br/>硬件工程师</td></tr>
          <tr><td>AI算法、软件开发、嵌入式软件、项目管理、产品规划、技术支持</td></tr>
          <tr><td>安全、质量管理岗</td></tr>
        </table>
        """,
    }
    names = [position.name for position in extract_positions(detail)]
    assert names == [
        "软件工程师",
        "硬件工程师",
        "AI算法",
        "软件开发",
        "嵌入式软件",
        "项目管理",
        "产品规划",
        "技术支持",
        "安全、质量管理岗",
    ]


def test_bracketed_locations_stay_with_role_and_space_concat_roles_split() -> None:
    detail = {
        "title": "",
        "recruitmentContent": """
        <table>
          <tr><th>岗位名称</th></tr>
          <tr><td>海外技术支持工程师（韩国、日本、美国）、硬件工程师</td></tr>
          <tr><td>SLAM算法工程师3DGS算法工程师DevOps开发工程师</td></tr>
          <tr><td>BOM工程师设备工程师、质量工程师</td></tr>
          <tr><td>技术型销售技术支持工程师</td></tr>
          <tr><td>销售工程师储干</td></tr>
          <tr><td>计算处理器开发-ASIC设计工程师等40+岗位</td></tr>
          <tr><td>集成电路设计 硬件电路设计 嵌入式开发 伺服控制 信号处理</td></tr>
          <tr><td>Product Operation Manager</td></tr>
        </table>
        """,
    }
    names = [position.name for position in extract_positions(detail)]
    assert names == [
        "海外技术支持工程师(韩国、日本、美国)",
        "硬件工程师",
        "SLAM算法工程师",
        "3DGS算法工程师",
        "DevOps开发工程师",
        "BOM工程师",
        "设备工程师",
        "质量工程师",
        "技术型销售",
        "技术支持工程师",
        "销售工程师储干",
        "计算处理器开发-ASIC设计工程师等40+岗位",
        "集成电路设计",
        "硬件电路设计",
        "嵌入式开发",
        "伺服控制",
        "信号处理",
        "Product Operation Manager",
    ]


def test_slash_list_splits_outside_parentheses_but_not_direction_suffix() -> None:
    detail = {
        "title": "",
        "recruitmentContent": """
        <table>
          <tr><th>岗位</th></tr>
          <tr><td>嵌入式（MUC/BSP）/应用软件/总线/中间件/测试开发/功能安全/软件架构</td></tr>
          <tr><td>算法工程师：语音/计算机视觉/算法平台</td></tr>
          <tr><td>工艺/设备/采购/质量</td></tr>
          <tr><td>光学装调工艺 生产装调工艺</td></tr>
        </table>
        """,
    }
    names = [position.name for position in extract_positions(detail)]
    assert names == [
        "嵌入式(MUC/BSP)",
        "应用软件",
        "总线",
        "中间件",
        "测试开发",
        "功能安全",
        "软件架构",
        "算法工程师:语音/计算机视觉/算法平台",
        "工艺",
        "设备",
        "采购",
        "质量",
        "光学装调工艺",
        "生产装调工艺",
    ]
