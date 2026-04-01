"""
只读邮箱客户端（协作版）
- 统一使用 IMAP readonly 模式打开收件箱
- 仅提供查询、搜索、读取能力，不提供任何写入操作
"""

import imaplib
import email
from email.header import decode_header
from datetime import date, timedelta
from typing import Dict, List, Optional


DEFAULT_KEYWORDS = [
    "基金",
    "净值",
    "持仓",
    "日报",
    "周报",
    "月报",
    "NAV",
    "Fund",
    "估值",
]


class ReadonlyFareportMailClient:
    """只读 IMAP 客户端，面向协作同事使用。"""

    def __init__(
        self,
        email_address: str,
        password: str,
        imap_server: str = "imap.exmail.qq.com",
        imap_port: int = 993,
    ):
        self._email_address = email_address
        self._password = password
        self._imap_server = imap_server
        self._imap_port = imap_port
        self._mailbox: Optional[imaplib.IMAP4_SSL] = None

    def connect(self) -> bool:
        """连接邮箱并以只读方式选择 INBOX。"""
        try:
            self._mailbox = imaplib.IMAP4_SSL(self._imap_server, self._imap_port)
            self._mailbox.login(self._email_address, self._password)
            status, _ = self._mailbox.select("INBOX", readonly=True)
            return status == "OK"
        except Exception:
            return False

    def disconnect(self) -> None:
        if not self._mailbox:
            return
        try:
            self._mailbox.close()
        except Exception:
            pass
        try:
            self._mailbox.logout()
        except Exception:
            pass
        self._mailbox = None

    def get_mail_count(self) -> int:
        if not self._mailbox:
            return 0
        try:
            status, messages = self._mailbox.status("INBOX", "(MESSAGES)")
            if status != "OK":
                return 0
            return int(messages[0].decode().split()[2].strip(")"))
        except Exception:
            return 0

    def search_emails(
        self,
        days: int = 7,
        unread: bool = False,
        sender_keywords: Optional[List[str]] = None,
        subject_keywords: Optional[List[str]] = None,
    ) -> List[str]:
        """搜索邮件并做发件人/主题关键词过滤。"""
        if not self._mailbox:
            return []

        sender_keywords = [] if sender_keywords is None else sender_keywords
        subject_keywords = DEFAULT_KEYWORDS if subject_keywords is None else subject_keywords

        criteria: List[str] = []
        if unread:
            criteria.append("UNSEEN")

        since_date = date.today() - timedelta(days=days)
        criteria.append(f'SINCE "{since_date.strftime("%d-%b-%Y")}"')

        status, messages = self._mailbox.search(None, " ".join(criteria))
        if status != "OK":
            return []

        matched: List[str] = []
        for raw_id in messages[0].split():
            mail_id = raw_id.decode()
            status, msg_data = self._mailbox.fetch(mail_id, "(BODY[HEADER.FIELDS (SUBJECT FROM DATE)])")
            if status != "OK":
                continue

            header_text = msg_data[0][1].decode("utf-8", errors="ignore").lower()
            sender_ok = True
            subject_ok = True

            if sender_keywords:
                sender_ok = any(k.lower() in header_text for k in sender_keywords)
            if subject_keywords:
                subject_ok = any(k.lower() in header_text for k in subject_keywords)

            if sender_ok and subject_ok:
                matched.append(mail_id)

        return matched

    def fetch_email_summary(self, mail_id: str) -> Optional[Dict[str, str]]:
        """只返回邮件摘要，不返回正文和附件二进制。"""
        if not self._mailbox:
            return None
        try:
            status, msg_data = self._mailbox.fetch(mail_id, "(RFC822.HEADER)")
            if status != "OK":
                return None

            email_message = email.message_from_bytes(msg_data[0][1])
            return {
                "id": mail_id,
                "subject": _decode_mime_header(email_message.get("Subject", "")),
                "from": _decode_mime_header(email_message.get("From", "")),
                "date": email_message.get("Date", ""),
            }
        except Exception:
            return None

    def fetch_email_body_preview(self, mail_id: str, max_chars: int = 200) -> str:
        """返回正文预览文本（纯文本优先），仅用于只读查看。"""
        if not self._mailbox:
            return ""
        try:
            status, msg_data = self._mailbox.fetch(mail_id, "(RFC822)")
            if status != "OK":
                return ""
            email_message = email.message_from_bytes(msg_data[0][1])

            body_parts: List[str] = []
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))
                if "attachment" in content_disposition.lower():
                    continue
                if content_type == "text/plain":
                    payload = part.get_payload(decode=True) or b""
                    text = payload.decode("utf-8", errors="ignore").strip()
                    if text:
                        body_parts.append(text)
                        break

            if not body_parts:
                return ""

            preview = body_parts[0].replace("\r", " ").replace("\n", " ")
            preview = " ".join(preview.split())
            return preview[:max_chars]
        except Exception:
            return ""


def _decode_mime_header(value: str) -> str:
    if not value:
        return ""
    parts = decode_header(value)
    output = ""
    for text, encoding in parts:
        if isinstance(text, bytes):
            try:
                output += text.decode(encoding or "utf-8")
            except Exception:
                output += text.decode("gbk", errors="ignore")
        else:
            output += text
    return output
