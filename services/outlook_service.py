import json
import os
from datetime import datetime
from html import escape

import win32com.client as win32

from config import (
    DEFAULT_BODY,
    DEFAULT_SUBJECT,
    OUTLOOK_TOKEN_ALIASES,
    OUTLOOK_TOKEN_PATTERN,
)
from utils.logger_utils import logger


def normalize_recipient_addresses(raw_addresses) -> list[str]:
    if not raw_addresses:
        return []

    if isinstance(raw_addresses, (list, tuple, set)):
        candidates = raw_addresses
    else:
        candidates = str(raw_addresses).replace(",", ";").split(";")

    recipients = []
    seen = set()

    for item in candidates:
        email = str(item).strip()
        if not email:
            continue

        key = email.lower()
        if key in seen:
            continue

        seen.add(key)
        recipients.append(email)

    return recipients


def dedupe_email_list(emails: list[str]) -> list[str]:
    return normalize_recipient_addresses(emails)


def build_outlook_template_values(report_data: dict) -> dict[str, str]:
    r_date = report_data.get("report_date", "")
    try:
        dt = datetime.strptime(r_date, "%Y-%m-%d")
        year_str = str(dt.year)
        month_str = str(dt.month)
        day_str = str(dt.day)
        weekday_str = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
    except (ValueError, IndexError):
        year_str, month_str, day_str, weekday_str = "", "", "", ""

    data_map = {
        "report_date": r_date,
        "year": year_str,
        "month": month_str,
        "day": day_str,
        "weekday": weekday_str,
        "department": report_data.get("department", ""),
        "author_name": report_data.get("author_name", ""),
    }
    return {token: str(data_map.get(key, "")) for token, key in OUTLOOK_TOKEN_ALIASES.items()}


def find_unknown_outlook_tokens(text: str) -> list[str]:
    unknown = []
    seen = set()

    for token in OUTLOOK_TOKEN_PATTERN.findall(text or ""):
        if token in OUTLOOK_TOKEN_ALIASES:
            continue

        if token in seen:
            continue

        seen.add(token)
        unknown.append(token)

    return unknown


def render_outlook_template(
    template_text: str,
    template_values: dict[str, str],
) -> str:
    result = str(template_text or "")

    for token, value in template_values.items():
        replacement = "" if value is None else str(value)
        result = result.replace(token, replacement)

        alias_name = OUTLOOK_TOKEN_ALIASES.get(token)
        if alias_name:
            result = result.replace(
                f"{{{alias_name}}}",
                replacement,
            )

    unknown_tokens = find_unknown_outlook_tokens(result)
    if unknown_tokens:
        logger.warning(
            "메일 템플릿에 미치환 토큰이 남아 있음: count=%d",
            len(unknown_tokens),
        )

    return result


def build_safe_html_text(text: str) -> str:
    return escape(text or "").replace("\n", "<br>")


def create_outlook_draft(report_data: dict, output_paths: tuple) -> str:
    """PDF/PNG 생성 후 Outlook 초안 메일을 작성(Display)한다. 절대 자동 발송(Send)하지 않는다."""
    if os.getenv("OUTLOOK_ENABLE", "False") != "True":
        logger.info("아웃룩 기능 비활성화로 메일 초안 생성을 건너뜁니다.")
        return "disabled"

    output_xlsx, output_pdf, output_png = output_paths

    norm_pdf = os.path.abspath(os.path.normpath(output_pdf)) if output_pdf else ""
    if not os.path.isfile(norm_pdf):
        logger.error("Outlook 초안 생성 중단: PDF 없음")
        return "failed"

    if os.path.getsize(norm_pdf) <= 0:
        logger.error("Outlook 초안 생성 중단: PDF가 0바이트")
        return "failed"

    logger.info("아웃룩 초안 생성 시작")

    # 1. 수신자 / 참조 주소 정제
    raw_to = os.getenv("OUTLOOK_TO", "")
    raw_cc = os.getenv("OUTLOOK_CC", "")
    to_list = normalize_recipient_addresses(raw_to)
    cc_list = normalize_recipient_addresses(raw_cc)

    target_email = "; ".join(to_list)
    cc_email = "; ".join(cc_list)
    sender_email = os.getenv("OUTLOOK_SENDER", "").strip().lower()

    subject_tmpl = os.getenv("OUTLOOK_SUBJECT", DEFAULT_SUBJECT)
    raw_body = os.getenv("OUTLOOK_BODY", "")
    if raw_body:
        try:
            body_tmpl = json.loads(raw_body)
        except (json.JSONDecodeError, ValueError):
            body_tmpl = raw_body
    else:
        body_tmpl = DEFAULT_BODY

    template_values = build_outlook_template_values(report_data)
    final_subject = render_outlook_template(subject_tmpl, template_values)
    final_body_text = render_outlook_template(body_tmpl, template_values)
    escaped_body = build_safe_html_text(final_body_text)

    # 2. 아웃룩 실행
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        if sender_email:
            selected_account = None
            try:
                accounts = outlook.Session.Accounts
                for i in range(1, int(accounts.Count) + 1):
                    acct = accounts.Item(i)
                    smtp = (getattr(acct, "SmtpAddress", "") or "").strip().lower()
                    if smtp == sender_email:
                        selected_account = acct
                        break
            except Exception:
                logger.exception("발신 계정 조회 중 예외 발생")

            if selected_account is not None:
                try:
                    mail.SendUsingAccount = selected_account
                    logger.info("지정 발신 계정 적용 완료")
                except Exception:
                    logger.exception("발신 계정 적용 실패 (기본 계정으로 진행)")
            else:
                logger.warning("지정한 발신 계정을 찾지 못함 (기본 계정 사용)")

        if target_email:
            mail.To = target_email
        if cc_email:
            mail.CC = cc_email

        mail.Subject = final_subject

        # PDF 필수 첨부
        try:
            mail.Attachments.Add(norm_pdf)
        except Exception:
            logger.exception("PDF 파일 메일 첨부 실패")
            return "failed"

        # 메일 초안 창 먼저 표시 (Outlook 기본 서명이 자동 생성됨)
        mail.Display()
        existing_signature_html = mail.HTMLBody or ""

        # PNG 선택 첨부
        image_html = ""
        norm_png = os.path.abspath(os.path.normpath(output_png)) if output_png else ""
        if os.path.isfile(norm_png) and os.path.getsize(norm_png) > 0:
            try:
                attachment = mail.Attachments.Add(norm_png)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    "daily_report_img",
                )
                image_html = "<br><br><img src='cid:daily_report_img'>"
            except Exception:
                logger.exception("PNG 본문 이미지 첨부 실패")

        # HTML 본문 + 이미지 + 기존 서명 병합
        mail.HTMLBody = (
            "<div style=\"font-family:'맑은 고딕',Arial,sans-serif;font-size:10pt;\">"
            f"{escaped_body}"
            f"{image_html}"
            "</div>"
            "<br>"
            f"{existing_signature_html}"
        )

        logger.info("아웃룩 메일 초안 창 표시 완료 (자동 전송 안 함)")
        return "created"

    except Exception:
        logger.exception("아웃룩 메일 초안 생성 중 예외 발생")
        return "failed"
