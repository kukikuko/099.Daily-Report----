from services.outlook_service import (
    build_safe_html_text,
    dedupe_email_list,
    find_unknown_outlook_tokens,
    normalize_recipient_addresses,
    render_outlook_template,
)


def test_normalize_recipient_addresses():
    raw = "user1@example.com; User2@example.com, user1@example.com; "
    res = normalize_recipient_addresses(raw)
    assert res == ["user1@example.com", "User2@example.com"]

    assert normalize_recipient_addresses([]) == []
    assert normalize_recipient_addresses(["a@b.com", "A@b.com", "c@b.com"]) == ["a@b.com", "c@b.com"]


def test_dedupe_email_list():
    emails = ["test@example.com", "TEST@example.com", "other@example.com"]
    assert dedupe_email_list(emails) == ["test@example.com", "other@example.com"]


def test_find_unknown_outlook_tokens():
    text = "Hello [[년]] [[월]] [[알수없는토큰]] {unknown_var}"
    unknown = find_unknown_outlook_tokens(text)
    assert "[[알수없는토큰]]" in unknown
    assert "{unknown_var}" in unknown
    assert "[[년]]" not in unknown


def test_render_outlook_template():
    template = "[[년]]년 [[월]]월 [[일]]일 보고서"
    values = {"[[년]]": "2026", "[[월]]": "07", "[[일]]": "30"}
    rendered = render_outlook_template(template, values)
    assert rendered == "2026년 07월 30일 보고서"


def test_build_safe_html_text():
    text = "<script>alert('xss')</script>\nLine2"
    escaped = build_safe_html_text(text)
    assert "&lt;script&gt;" in escaped
    assert "<br>" in escaped
