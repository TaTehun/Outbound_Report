"""
email_pipeline_example.py

Reference implementation of the email pipeline.
Copy and split into email_setup.py / downloader.py / sender.py for production use.

Sections:
  1. Configuration  (email_setup.py)
  2. Downloader     (downloader.py)
  3. Sender         (sender.py)
"""

# ============================================================
# 1. CONFIGURATION  (email_setup.py)
# ============================================================
from pathlib import Path
from utils.date_utils import today_str

DESKTOP_PATH  = Path.home() / 'Desktop'
_BASE_DATA    = DESKTOP_PATH / 'YOUR_PROJECT' / 'Data'
_BASE_OUTPUT  = DESKTOP_PATH / 'YOUR_PROJECT' / 'Output'
DATA_DIR      = _BASE_DATA   / today_str('f')
OUTPUT_DIR    = _BASE_OUTPUT / today_str('f')
TEMPLATE_PATH = _BASE_DATA   / 'template.xlsx'

# .env file path, holds POP3/SMTP credentials
ENV_PATH = Path(r'C:\path\to\your\.env')

# POP3
POP3_SERVER   = 'pop3.your-mail-server.com'
POP3_PORT     = 995
POP3_USER_ENV = 'POP3_USERNAME'   # key name in .env
POP3_PASS_ENV = 'POP3_PASSWORD'   # key name in .env

EMAIL_SEARCH_COUNT = 1000
MAX_DAILY_FOLDERS  = 5

# SMTP
SMTP_SERVER   = 'smtp.your-mail-server.com'
SMTP_PORT     = 25
SMTP_USER_ENV = 'SMTP_USERNAME'   # key name in .env
SMTP_PASS_ENV = 'SMTP_PASSWORD'   # key name in .env

NOTIFICATION_FROM = 'your-sender@example.com'

# Email targets: subjects/senders to match for attachment download
EMAIL_TARGETS = [
    {
        "name": "Carrier A - Report Name",
        "sender": "sender@carrier-a.com",        # set None to match any sender
        "contains": ["keyword in subject"],
        "dates": [today_str('g')],               # omit key if no date filter needed
    },
    {
        "name": "Carrier B - Report Name",
        "sender": None,
        "contains": ["keyword in subject"],
        "dates": [today_str('i')],
    },
    # add more targets as needed
]

ATTACHMENT_FILENAME = "Consolidated Report - {}.xlsb"  # .format(date)
ATTACHMENT_MIME     = ("application", "vnd.ms-excel.sheet.binary.macroEnabled.12")

# Failure alert
FAILURE_SUBJECT  = 'Daily Report - Files missing'
NOTIFICATION_TO  = ['recipient@example.com']
NOTIFICATION_CC  = []
NOTIFICATION_BCC = []

# Report
REPORT_SUBJECT = 'Consolidated Report - {}'   # .format(date)
REPORT_TO  = ['recipient@example.com']
REPORT_CC  = []
REPORT_BCC = []

# Skip (No Update)
SKIP_SUBJECT = '[No Update] Consolidated Report - {}'  # .format(date)
SKIP_TO  = ['recipient@example.com']
SKIP_CC  = []
SKIP_BCC = []


# ============================================================
# 2. DOWNLOADER  (downloader.py)
# ============================================================
import poplib
import logging
import shutil

from email.parser import Parser
from email.utils import parsedate_to_datetime
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import List, Optional
import pytz

logger = logging.getLogger(__name__)
_tz = pytz.timezone('America/Chicago')  # adjust to your timezone


@dataclass
class DownloadResult:
    success: bool
    downloaded_files: List[str] = field(default_factory=list)
    received_times: dict = field(default_factory=dict)
    error_message: Optional[str] = None


class EmailDownloader:
    """
    Downloads attachments from a POP3 server.

    Usage:
        downloader = EmailDownloader(user, password)
        result = downloader.download()
    """

    def __init__(self, user: str, password: str):
        self.user = user
        self.password = password
        self._connection: Optional[poplib.POP3_SSL] = None
        self._downloaded: List[str] = []

    def download(self) -> DownloadResult:
        try:
            self._downloaded = []
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            self._prune_old_folders()
            self._connect()
            count, size = self._connection.stat()
            logger.info(f"Mailbox: {count} messages, {size} bytes")
            return self._scan()
        except Exception as e:
            logger.error(f"Download failed: {e}", exc_info=True)
            return DownloadResult(success=False, error_message=str(e))
        finally:
            self._close()

    def _connect(self) -> None:
        if self._connection:
            return
        logger.info(f"Connecting to {POP3_SERVER}:{POP3_PORT}")
        self._connection = poplib.POP3_SSL(POP3_SERVER, POP3_PORT)
        self._connection.user(self.user)
        self._connection.pass_(self.password)
        logger.info("POP3 login successful")

    def _close(self) -> None:
        if self._connection:
            try:
                self._connection.quit()
                logger.info("POP3 connection closed")
            except Exception as e:
                logger.warning(f"Error during quit: {e}")
            finally:
                self._connection = None

    def _prune_old_folders(self) -> None:
        for base in [DATA_DIR.parent, OUTPUT_DIR.parent]:
            folders = sorted([f for f in base.iterdir() if f.is_dir()], key=lambda f: f.name, reverse=True)
            for old in folders[MAX_DAILY_FOLDERS:]:
                shutil.rmtree(old)
                logger.info(f"Removed old folder: {old}")

    def _scan(self) -> DownloadResult:
        today = datetime.now(_tz).date()
        _, mails_list, _ = self._connection.list()
        recent = list(reversed(mails_list[-EMAIL_SEARCH_COUNT:]))

        matched_keys: set = set()
        received_times: dict = {}

        for mail in recent:
            idx = mail.split()[0].decode() if isinstance(mail, bytes) else str(mail)[2:].split(' ')[0].strip()

            headers = self._fetch_headers(idx)
            try:
                if parsedate_to_datetime(headers.get('Date', '')).astimezone(_tz).date() < today:
                    logger.info(f"Reached mail older than today at index {idx}, stopping")
                    break
            except Exception:
                pass

            matched = self._match(headers, matched_keys)
            if matched:
                msg = self._fetch_email(idx)
                email_dt = parsedate_to_datetime(msg.get('Date', '')).astimezone(_tz)
                logger.info(f"Match [{matched['name']}] | {email_dt.strftime('%Y-%m-%d %H:%M %Z')} | {msg.get('Subject', '')}")
                self._save_attachments(msg, email_dt)
                received_times[matched['name']] = email_dt.strftime('%m/%d/%Y %H:%M CST')
                matched_keys.add(tuple(matched['contains']))

                if len(matched_keys) >= len(EMAIL_TARGETS):
                    logger.info("All expected attachments downloaded")
                    return DownloadResult(success=True, downloaded_files=self._downloaded, received_times=received_times)

        return DownloadResult(
            success=False,
            downloaded_files=self._downloaded,
            received_times=received_times,
            error_message=f"Expected {len(EMAIL_TARGETS)} files, got {len(self._downloaded)}",
        )

    def _fetch_headers(self, idx: str):
        _, lines, _ = self._connection.top(idx, 0)
        return Parser().parsestr(b'\r\n'.join(lines).decode('utf-8', 'replace'))

    def _fetch_email(self, idx: str):
        _, lines, _ = self._connection.retr(idx)
        return Parser().parsestr(b'\r\n'.join(lines).decode('utf-8', 'replace'))

    @staticmethod
    def _match(msg, exclude: set) -> Optional[dict]:
        from_addr = msg.get('From', '')
        subject = msg.get('Subject', '').strip()
        try:
            email_date = parsedate_to_datetime(msg.get('Date', '')).astimezone(_tz).date()
            if email_date < (datetime.now(_tz) - timedelta(days=5)).date():
                return None
        except Exception:
            pass
        for t in EMAIL_TARGETS:
            if tuple(t['contains']) in exclude:
                continue
            if (
                (t.get('sender') is None or t['sender'] in from_addr)
                and all(c in subject for c in t['contains'])
                and (t.get('dates') is None or any(d in subject for d in t['dates']))
            ):
                return t
        return None

    def _save_attachments(self, msg, email_dt) -> None:
        content_type = msg.get_content_type().lower()
        if content_type.startswith('multipart'):
            for part in msg.get_payload():
                self._save_attachments(part, email_dt)
            return
        if not (content_type.startswith('application') or content_type == 'text/csv'):
            return
        filename = msg.get_filename()
        if not filename:
            return
        filename = filename.replace('\r\n', '').replace('\t', ' ')
        stem, _, suffix = filename.rpartition('.')
        timestamp = email_dt.strftime('%H%M')
        final_name = f"{stem}_{timestamp}.{suffix}" if suffix else f"{filename}_{timestamp}"
        path = DATA_DIR / final_name
        path.write_bytes(msg.get_payload(decode=True))
        self._downloaded.append(str(path))
        logger.info(f"Saved: {path}")


# ============================================================
# 3. SENDER  (sender.py)
# ============================================================
import os
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders
from dotenv import load_dotenv

from utils import pivot_utils


class EmailSender:
    """
    Low-level SMTP sender. Use send_report / send_failure_alert / send_skip
    for pre-built email types.
    """

    def __init__(self, user: str, password: str):
        self.user = user
        self.password = password

    def send(
        self,
        subject: str,
        body: str,
        to: List[str],
        cc: List[str] = None,
        bcc: List[str] = None,
        sender: str = None,
        attachment: Optional[Path] = None,
        attachment_mime: tuple = ("application", "octet-stream"),
    ) -> bool:
        cc  = cc  or []
        bcc = bcc or []
        from_addr = sender or NOTIFICATION_FROM

        root = MIMEMultipart('related')
        root['From']    = from_addr
        root['To']      = ','.join(to)
        root['Cc']      = ','.join(cc)
        root['Date']    = formatdate(localtime=True)
        root['Subject'] = subject

        alt = MIMEMultipart('alternative')
        root.attach(alt)
        alt.attach(MIMEText(body, 'html', 'utf-8'))

        if attachment:
            attachment = Path(attachment)
            if attachment.exists():
                part = MIMEBase(*attachment_mime)
                part.set_payload(attachment.read_bytes())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=attachment.name)
                root.attach(part)
            else:
                logger.warning(f"Attachment not found, skipping: {attachment}")

        recipients = to + cc + bcc
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.login(self.user, self.password)
                smtp.sendmail(from_addr, recipients, root.as_string())
            logger.info(f"Mail sent → {recipients}")
            return True
        except Exception as e:
            logger.error(f"Failed to send mail: {e}", exc_info=True)
            return False


def load_sender() -> Optional[EmailSender]:
    """Load SMTP credentials from .env and return an EmailSender instance."""
    load_dotenv(ENV_PATH)
    user = os.environ.get(SMTP_USER_ENV)
    pw   = os.environ.get(SMTP_PASS_ENV)
    if not user or not pw:
        logger.error("SMTP credentials missing")
        return None
    return EmailSender(user, pw)


def send_report(sender: EmailSender, output_file: Path, received_times: dict, is_updated: bool = False, pivot_data: list = None) -> bool:
    pivot_html = pivot_utils.to_html(pivot_data) if pivot_data else ""
    all_targets = [t['name'] for t in EMAIL_TARGETS]
    missing  = [n for n in all_targets if n not in received_times]
    received = [n for n in all_targets if n in received_times]

    file_lines = []
    if missing:
        file_lines += ["The following files were not received today:"]
        file_lines += [f"ㆍ {n}" for n in missing]
        file_lines += [""]
    file_lines += ["The following files were received today:"]
    file_lines += [f"ㆍ {n}" for n in received]

    body = (
        '<BODY style="font-size:11pt;font-family:Calibri">'
        + "Dear All,<br><br>"
        + f"Please find the consolidated Daily Outbound Report for {today_str('h')} Attached.<br><br>"
        + pivot_html
        + "<br>" + "<br>".join(file_lines) + "<br><br>"
        + "Thank you"
        + "</BODY>"
    )

    subject = REPORT_SUBJECT.format(today_str('e'))
    if is_updated:
        subject = "(Updated) " + subject

    return sender.send(
        subject=subject,
        body=body,
        to=REPORT_TO,
        cc=REPORT_CC,
        bcc=REPORT_BCC,
        attachment=output_file,
        attachment_mime=ATTACHMENT_MIME,
    )


def send_failure_alert(sender: EmailSender, received_times: dict = None) -> bool:
    received_times = received_times or {}
    all_targets = [t['name'] for t in EMAIL_TARGETS]
    missing  = [n for n in all_targets if n not in received_times]
    received = [n for n in all_targets if n in received_times]

    lines = ["Dear All,", ""]
    if missing:
        lines += ["The following files were not received today:"]
        lines += [f"ㆍ {n}" for n in missing]
        lines += [""]
    if received:
        lines += ["The following files were received today:"]
        lines += [f"ㆍ {n} (Received Date: {received_times[n]})" for n in received]
        lines += [""]
    lines.append("Thank you")

    body = '<BODY style="font-size:11pt;font-family:Calibri">' + "<br>".join(lines) + "</BODY>"
    return sender.send(
        subject=FAILURE_SUBJECT,
        body=body,
        to=NOTIFICATION_TO,
        cc=NOTIFICATION_CC,
        bcc=NOTIFICATION_BCC,
    )


def send_skip(sender: EmailSender, downloaded_files: list) -> bool:
    file_names = [Path(f).name for f in downloaded_files]
    lines = (
        ["Dear All,", ""]
        + ["The following files from the second run were identical to the first run.",
           "No updated report has been sent.", ""]
        + [f"ㆍ {n}" for n in file_names]
        + ["", "Thank you"]
    )
    body = '<BODY style="font-size:11pt;font-family:Calibri">' + "<br>".join(lines) + "</BODY>"
    return sender.send(
        subject=SKIP_SUBJECT.format(today_str('e')),
        body=body,
        to=SKIP_TO,
        cc=SKIP_CC,
        bcc=SKIP_BCC,
    )
