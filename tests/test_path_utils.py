from utils.path_utils import parse_doc_number, sanitize_filename_part


def test_sanitize_filename_part():
    assert sanitize_filename_part("Daily/Report:Test*") == "Daily_Report_Test_"
    assert sanitize_filename_part("  ") == "default"
    assert sanitize_filename_part(None, fallback="fallback") == "fallback"


def test_parse_doc_number():
    assert parse_doc_number("005", "001") == "005"
    assert parse_doc_number("", "001") == "001"
    assert parse_doc_number("005/abc", "001") is None
