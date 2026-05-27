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
