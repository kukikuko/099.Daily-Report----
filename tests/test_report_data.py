from models.report_data import DailyReportData


def test_daily_report_data_validation():
    data = DailyReportData(
        report_date="2026-07-30",
        department="TE",
        work_location="본사",
        author_name="홍길동",
        headcount=1,
    )
    assert data.validate() == []

    invalid_data = DailyReportData(
        report_date="2026-07-30",
        department="",
        work_location="",
        author_name="",
        headcount=0,
    )
    errors = invalid_data.validate()
    assert len(errors) == 4


def test_from_dict_safe_headcount():
    raw = {
        "report_date": "2026-07-30",
        "department": "TE",
        "work_location": "본사",
        "author_name": "홍길동",
        "headcount": "invalid_number",
    }
    obj = DailyReportData.from_dict(raw)
    assert obj.headcount == 1
    assert obj.author_name == "홍길동"
