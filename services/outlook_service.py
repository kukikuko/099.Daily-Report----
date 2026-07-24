import os
import json
from datetime import datetime
import win32com.client as win32
from config import DEFAULT_SUBJECT, DEFAULT_BODY, OUTLOOK_TOKEN_ALIASES, OUTLOOK_TOKEN_PATTERN
from utils.logger_utils import logger

def normalize_recipient_addresses(raw_text: str) -> list:
    recipients = []
    seen = set()
    for token in (raw_text or "").split(";"):
        email = token.strip()
        if not email:
            continue
        email_key = email.lower()
        if email_key in seen:
            continue
        seen.add(email_key)
        recipients.append(email)
    return recipients

def dedupe_email_list(emails: list) -> list:
    recipients = []
    seen = set()
    for email in emails:
        value = str(email).strip()
        if not value:
            continue
        email_key = value.lower()
        if email_key in seen:
            continue
        seen.add(email_key)
        recipients.append(value)
    return recipients

def build_outlook_template_values(report_data: dict) -> dict:
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

def render_outlook_template(template_text: str, template_values: dict) -> str:
    result = str(template_text or "")
    for token, value in template_values.items():
        sub_val = "" if value is None else str(value)
        result = result.replace(token, sub_val)
        if token in OUTLOOK_TOKEN_ALIASES:
            alias = f"{{{OUTLOOK_TOKEN_ALIASES[token]}}}"
            result = result.replace(alias, sub_val)

    unrendered = OUTLOOK_TOKEN_PATTERN.findall(result)
    if unrendered:
        logger.warning(f"메일 템플릿에 미치환된 토큰이 남아있습니다: {unrendered}")

    return result

def find_unknown_outlook_tokens(text: str) -> list:
    unknown = []
    seen = set()
    for token in OUTLOOK_TOKEN_PATTERN.findall(text or ""):
        if token in OUTLOOK_TOKEN_ALIASES or token in seen:
            continue
        seen.add(token)
    for token in raw_tokens:
        if token not in OUTLOOK_FRIENDLY_TOKENS:
            unknown.append(token)
    return list(dict.fromkeys(unknown))

def render_outlook_template(template_text: str, template_values: dict) -> str:
    if not template_text:
        return ""

    result = template_text
    for token, val in template_values.items():
        result = result.replace(token, str(val))

    unknown_tokens = find_unknown_outlook_tokens(result)
    if unknown_tokens:
        logger.warning(f"미치환 템플릿 토큰 존재: {unknown_tokens}")

    return result

def normalize_recipient_addresses(raw_addresses) -> list:
    if not raw_addresses:
        return []

    if isinstance(raw_addresses, list):
        candidates = raw_addresses
    else:
        text = str(raw_addresses).replace(',', ';')
        candidates = text.split(';')

    result = []
    seen = set()
    for item in candidates:
        email = str(item).strip()
        if email:
            key = email.lower()
            if key not in seen:
                seen.add(key)
                result.append(email)

    return result

def create_outlook_draft(report_data: dict, output_paths: tuple) -> str:
    """PDF/PNG 생성 후 Outlook 초안 메일을 작성(Display)한다. 절대 자동 발송(Send)하지 않는다."""
    if os.getenv("OUTLOOK_ENABLE", "False") != "True":
        logger.info("아웃룩 기능 비활성화로 메일 초안 생성을 건너뜁니다.")
        return "disabled"

    output_xlsx, output_pdf, output_png = output_paths

    norm_pdf = os.path.abspath(os.path.normpath(output_pdf)) if output_pdf else ""
    if not os.path.exists(norm_pdf) or os.path.getsize(norm_pdf) == 0:
        logger.error("PDF 첨부 파일이 미존재하거나 0바이트여서 아웃룩 초안 생성을 중단합니다.")
        return "failed"

    logger.info("아웃룩 초안 생성 시작")
    
    # 1. 수신자 / 참조 주소 정제 (세미콜론 분할 및 중복 제거)
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
    escaped_body = html.escape(final_body_text).replace('\n', '<br>')

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

        # 메일 초안 창 먼저 표시 (Outlook 기본 서명이 자동 생성됨)
        mail.Display()
        existing_signature_html = mail.HTMLBody or ""

        # PNG 유효 시 CID 본문 이미지 구성
        img_tag = ""
        norm_png = os.path.abspath(os.path.normpath(output_png)) if output_png else ""
        if os.path.exists(norm_png) and os.path.getsize(norm_png) > 0:
            try:
                attachment = mail.Attachments.Add(norm_png)
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "daily_report_img")
                img_tag = "<br><br><img src='cid:daily_report_img'>"
            except Exception:
                logger.exception("PNG 본문 이미지 CID 첨부 중 예외 발생")

        # PDF 필수 첨부
        try:
            mail.Attachments.Add(norm_pdf)
        except Exception:
            logger.exception("PDF 파일 메일 첨부 중 예외 발생")

        # HTML 본문 + 이미지 + 기존 서명 병합
        mail.HTMLBody = f'<div style="font-family: \'맑은 고딕\', Arial; font-size: 10pt;">{escaped_body}{img_tag}</div><br>{existing_signature_html}'

        logger.info("아웃룩 메일 초안 창 표시 완료 (자동 전송 안 함)")
        return "created"

    except Exception:
        logger.exception("아웃룩 메일 초안 생성 중 예외 발생")
        return "failed"
