# -*- coding: utf-8 -*-
"""
발주서 이메일 발송.
엑셀 EmailTemplate 시트 형식({{STORE_NAME}}, {{SUPPLIER_NAME}}, {{ORDER_DATE}}, {{ITEM_LIST}}, {{INTERNAL_OWNER}}) 사용.
발신: jhleele2@gmail.com (설정 가능)
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from typing import List, Dict, Any, Optional

# 기본 발신 이메일 (엑셀/요청 기준)
DEFAULT_SENDER_EMAIL = "jhleele2@gmail.com"
DEFAULT_STORE_NAME = "도미노피자"
DEFAULT_INTERNAL_OWNER = "jhleele2@gmail.com"


def build_item_list(items: List[Dict]) -> str:
    lines = []
    for i in items:
        lines.append(
            f"- {i.get('name', '')} ({i.get('code', '')}): {i.get('order_quantity', 0)} {i.get('unit', '')}"
        )
    return "\n".join(lines) if lines else "(항목 없음)"


def fill_template(
    subject_tpl: str,
    body_tpl: str,
    supplier_name: str,
    items: List[Dict],
    store_name: str = DEFAULT_STORE_NAME,
    order_date: Optional[str] = None,
    internal_owner: str = DEFAULT_INTERNAL_OWNER,
) -> tuple:
    order_date = order_date or datetime.now().strftime("%Y-%m-%d")
    item_list = build_item_list(items)
    subject = (
        subject_tpl.replace("{{STORE_NAME}}", store_name)
        .replace("{{SUPPLIER_NAME}}", supplier_name)
        .replace("{{ORDER_DATE}}", order_date)
    )
    body = (
        body_tpl.replace("{{STORE_NAME}}", store_name)
        .replace("{{SUPPLIER_NAME}}", supplier_name)
        .replace("{{ORDER_DATE}}", order_date)
        .replace("{{ITEM_LIST}}", item_list)
        .replace("{{INTERNAL_OWNER}}", internal_owner)
    )
    return subject, body


def send_order_email(
    to_email: str,
    subject: str,
    body: str,
    sender_email: str = DEFAULT_SENDER_EMAIL,
    sender_password: str = "",
    smtp_host: str = "smtp.gmail.com",
    smtp_port: int = 587,
    bcc: Optional[str] = None,
) -> tuple:
    """
    발주서 이메일 전송.
    bcc: 사본 수신 주소(예: jhleele2@gmail.com) — 발송한 메일을 본인 메일함에서도 확인 가능.
    Returns: (성공 여부, 메시지)
    """
    to_email = (to_email or "").strip()
    if not to_email or "@" not in to_email:
        return False, "수신 이메일 없음. 엑셀 Inventory 시트 '공급업체이메일' 열 또는 Suppliers 시트에 이메일을 입력하세요."
    if not sender_password:
        return False, "발신 비밀번호 미설정. Vercel: 프로젝트 설정 → Environment Variables에 SMTP_PASSWORD( Gmail 앱 비밀번호) 추가 후 재배포."
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = to_email
        if bcc and "@" in bcc:
            msg["Bcc"] = bcc
        msg.attach(MIMEText(body, "plain", "utf-8"))
        recipients = [to_email]
        if bcc and "@" in bcc:
            recipients.append(bcc)
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipients, msg.as_string())
        return True, "발송 완료"
    except smtplib.SMTPAuthenticationError as e:
        return False, f"Gmail 인증 실패. 앱 비밀번호를 확인하세요: {e}"
    except smtplib.SMTPRecipientsRefused as e:
        return False, f"수신 주소 거부: {e}"
    except Exception as e:
        return False, str(e)
