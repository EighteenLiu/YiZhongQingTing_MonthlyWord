from core.statistics import problem_display, summarize_records


def test_summarize_records_counts_indicators_by_street():
    records = [
        {
            "street": "新街口",
            "location": "点位A",
            "problem": "无防火标识",
            "indicators": ["无防火标识"],
            "has_problem": True,
        },
        {
            "street": "新街口",
            "location": "点位B",
            "problem": "无问题",
            "indicators": [],
            "has_problem": False,
        },
        {
            "street": "月坛",
            "location": "点位C",
            "problem": "灭火器不合格",
            "indicators": ["灭火器不合格"],
            "has_problem": True,
        },
    ]

    stats = summarize_records(records, "2026年4月20日", "2026年5月19日")

    assert stats["street_count"] == 2
    assert stats["record_count"] == 3
    assert stats["problem_record_count"] == 2
    assert stats["problem_count"] == 2
    assert stats["table_rows"][0]["无防火标识"] == 1
    assert stats["table_rows"][0]["灭火器不合格"] == 1


def test_problem_display_prefers_indicators():
    record = {"problem": "现场描述较长", "indicators": ["无回收价格表", "占道经营"]}

    assert problem_display(record) == "无回收价格表、占道经营"


def test_transfer_station_stats_prefer_second_level_indicator():
    records = [
        {
            "street": "西长安街街道",
            "location": "广州大厦",
            "problem": "无七禁收八不准承诺书",
            "specific_problem": "无七禁收八不准承诺书",
            "secondary_indicator": "无回收价格表",
            "indicators": ["无七禁收八不准承诺书"],
            "has_problem": True,
        }
    ]

    stats = summarize_records(records, "2026年4月20日", "2026年5月19日", report_type="transfer_station")

    assert stats["problem_record_count"] == 1
    assert stats["problem_count"] == 1
    assert stats["table_rows"][0]["无回收价格表"] == 1
    assert stats["table_rows"][0]["无七禁收八不准承诺书"] == 0
    assert "无回收价格表" in stats["summary_text"]
    assert "无七禁收八不准承诺书" not in stats["summary_text"]
    assert problem_display(records[0]) == "无回收价格表"


def test_summarize_records_ignores_uncategorized_problem_text():
    records = [
        {
            "street": "新街口",
            "location": "点位A",
            "problem": "表格里没有的描述",
            "indicators": [],
            "has_problem": False,
        },
        {
            "street": "新街口",
            "location": "点位B",
            "problem": "占道经营",
            "indicators": ["占道经营"],
            "has_problem": True,
        },
    ]

    stats = summarize_records(records, "2026年4月20日", "2026年5月19日")

    assert stats["problem_record_count"] == 1
    assert stats["problem_count"] == 1
    assert stats["table_rows"][0]["占道经营"] == 1
    assert stats["table_rows"][0]["问题总数"] == 1
    assert stats["table_rows"][1]["问题总数"] == 1


def test_summarize_records_uses_sorting_station_table_columns():
    records = [
        {
            "street": "大栅栏街道",
            "location": "前门西河沿街",
            "problem": "四分类桶不成组",
            "indicators": ["四分类桶不成组"],
            "has_problem": True,
        }
    ]

    stats = summarize_records(records, "2026年4月20日", "2026年5月19日", report_type="sorting_station")

    assert "四分类桶不成组" in stats["table_rows"][0]
    assert "无七禁收八不准承诺书" not in stats["table_rows"][0]
    assert stats["table_rows"][0]["四分类桶不成组"] == 1
