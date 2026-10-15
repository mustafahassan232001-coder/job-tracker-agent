import os
import smtplib
import logging
import mimetypes
from pathlib import Path
from html import escape
from email.message import EmailMessage
from typing import Any

from dotenv import load_dotenv

load_dotenv()
logger = logging.getLogger(__name__)


def send_new_jobs_email(
    jobs: list[dict[str, Any]],
    recipients: list[str] | None = None,
    subject: str | None = None,
    attachment_paths: list[str | Path] | None = None,
) -> bool:
    """Send one digest email containing jobs to the supplied recipients.

    If recipients is omitted, NOTIFICATION_EMAILS from .env is used for
    compatibility with the earlier single/multiple-recipient notifier.
    """
    if not jobs:
        logger.info("No jobs to email.")
        return False

    sender = os.getenv("SMTP_EMAIL", "").strip()
    app_password = os.getenv("SMTP_APP_PASSWORD", "").strip()
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com").strip()
    try:
        smtp_port = int(os.getenv("SMTP_PORT", "587"))
    except ValueError as exc:
        raise ValueError("SMTP_PORT must be a number, such as 587.") from exc

    if recipients is None:
        recipients = [
            address.strip()
            for address in os.getenv("NOTIFICATION_EMAILS", "").split(",")
            if address.strip()
        ]
    else:
        recipients = list(dict.fromkeys(
            address.strip() for address in recipients if address.strip()
        ))

    if not sender or not app_password:
        raise ValueError("Configure SMTP_EMAIL and SMTP_APP_PASSWORD in .env.")
    if not recipients:
        raise ValueError(
            "No recipients supplied. Configure the appropriate recipient variable in .env."
        )

    if not subject:
        subject = f"Job Search Agent: {len(jobs)} Matching Job(s)"

    text_lines = [f"Your job-search agent found {len(jobs)} matching job(s).", ""]
    html_items = []

    for index, job in enumerate(jobs, start=1):
        title = str(job.get("title") or job.get("Job Title") or "Untitled job")
        company = str(job.get("company") or job.get("Company") or "Unknown company")
        location = str(job.get("location") or job.get("Location") or "Location not specified")
        url = str(job.get("url") or job.get("Job URL") or job.get("link") or "")
        score = job.get("fit_score", job.get("match_score", job.get("Match Score", job.get("score"))))
        reason = str(job.get("reason") or job.get("Reason") or job.get("summary") or "")
        role_category = str(job.get("job_category") or job.get("category") or "")
        verification_status = str(job.get("verification_status") or "not provided")
        verification_reason = str(job.get("verification_reason") or "")

        text_lines.extend([f"{index}. {title}", f"Company: {company}", f"Location: {location}"])
        if role_category:
            text_lines.append(f"Job category: {role_category}")
        text_lines.append(f"URL verification: {verification_status}")
        if verification_reason:
            text_lines.append(f"Verification notes: {verification_reason}")
        if score is not None:
            text_lines.append(f"Match score: {score}/100")
        if url:
            text_lines.append(f"Job link: {url}")
        if reason:
            text_lines.append(f"Why it matches: {reason}")
        text_lines.append("")

        item = (
            f"<li><h3>{escape(title)}</h3>"
            f"<p><strong>Company:</strong> {escape(company)}</p>"
            f"<p><strong>Location:</strong> {escape(location)}</p>"
        )
        if score is not None:
            item += f"<p><strong>Match score:</strong> {escape(str(score))}/100</p>"
        if role_category:
            item += f"<p><strong>Category:</strong> {escape(role_category)}</p>"
        item += f"<p><strong>URL verification:</strong> {escape(verification_status)}</p>"
        if verification_reason:
            item += f"<p><strong>Verification notes:</strong> {escape(verification_reason)}</p>"
        if url:
            item += f'<p><a href="{escape(url, quote=True)}">View job</a></p>'
        if reason:
            item += f"<p><strong>Why it matches:</strong> {escape(reason)}</p>"
        item += "</li>"
        html_items.append(item)

    # Validate paths before rendering the message so its body explicitly states
    # whether job-specific documents are attached.
    valid_attachments: list[Path] = []
    for raw_path in attachment_paths or []:
        path = Path(raw_path)
        if path.is_file():
            valid_attachments.append(path)
        else:
            logger.warning("Job email attachment does not exist and was omitted: %s", path)
    attachment_html_note = ""
    if valid_attachments:
        text_lines.extend([
            "",
            "ATTACHED APPLICATION DOCUMENTS",
            "The attached category-specific ZIP contains tailored CV and cover-letter files for qualifying jobs in this category.",
            *[f"Attachment: {path.name}" for path in valid_attachments],
        ])
        attachment_html_note = (
            "<p><strong>Application documents attached.</strong> The category-specific ZIP contains tailored CV and cover-letter files for qualifying jobs in this category: "
            + ", ".join(escape(path.name) for path in valid_attachments)
            + "</p>"
        )
    elif attachment_paths is not None:
        text_lines.extend([
            "",
            "NO CATEGORY DOCUMENT ZIP WAS ATTACHED",
            "No generated CV/cover-letter files were available for this category. Check the application report and the MAX_APPLICATION_PACKAGES_PER_RUN limit.",
        ])
        attachment_html_note = (
            "<p><strong>No application documents attached.</strong> No generated CV/cover-letter files were available for this category. Check the application report and MAX_APPLICATION_PACKAGES_PER_RUN limit.</p>"
        )

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message.set_content("\n".join(text_lines))
    message.add_alternative(
        "<html><body><ol>" + "".join(html_items) + "</ol>" + attachment_html_note + "</body></html>",
        subtype="html",
    )

    # Attach a category-specific ZIP while keeping this notifier backward compatible.
    for path in valid_attachments:
        content_type, _ = mimetypes.guess_type(path.name)
        if content_type and "/" in content_type:
            maintype, subtype = content_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        message.add_attachment(
            path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name
        )

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(sender, app_password)
            smtp.send_message(message, to_addrs=recipients)
        logger.info("Email sent to %d recipient(s): %s", len(recipients), subject)
        return True
    except (smtplib.SMTPException, OSError) as exc:
        logger.exception("Email delivery failed.")
        raise RuntimeError(f"Could not send job notification: {exc}") from exc


def send_feedback_email(
    subject: str,
    text_body: str,
    html_body: str | None = None,
    recipients: list[str] | None = None,
) -> bool:
    """Send a plain-text/HTML career feedback email."""
    sender = os.getenv("SMTP_EMAIL", "").strip()
    app_password = os.getenv("SMTP_APP_PASSWORD", "").strip()
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com").strip()
    try:
        smtp_port = int(os.getenv("SMTP_PORT", "587"))
    except ValueError as exc:
        raise ValueError("SMTP_PORT must be a number, such as 587.") from exc

    if recipients is None:
        recipients = [
            address.strip()
            for address in os.getenv("NOTIFICATION_EMAILS", "").split(",")
            if address.strip()
        ]
    else:
        recipients = list(dict.fromkeys(
            address.strip() for address in recipients if address.strip()
        ))

    if not sender or not app_password:
        raise ValueError("Configure SMTP_EMAIL and SMTP_APP_PASSWORD in .env.")
    if not recipients:
        raise ValueError("No recipients supplied for career feedback email.")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message.set_content(text_body)
    if html_body:
        message.add_alternative(html_body, subtype="html")

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(sender, app_password)
            smtp.send_message(message, to_addrs=recipients)
        logger.info("Career feedback email sent to %d recipient(s).", len(recipients))
        return True
    except (smtplib.SMTPException, OSError) as exc:
        logger.exception("Career feedback email delivery failed.")
        raise RuntimeError(f"Could not send career feedback email: {exc}") from exc


def send_application_report_email(
    subject: str,
    text_body: str,
    html_body: str | None = None,
    recipients: list[str] | None = None,
    attachment_path: str | Path | None = None,
) -> bool:
    """Send application-package status and attach the generated ZIP when present."""
    sender = os.getenv("SMTP_EMAIL", "").strip()
    app_password = os.getenv("SMTP_APP_PASSWORD", "").strip()
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com").strip()
    try:
        smtp_port = int(os.getenv("SMTP_PORT", "587"))
    except ValueError as exc:
        raise ValueError("SMTP_PORT must be a number, such as 587.") from exc

    if recipients is None:
        recipients = [x.strip() for x in os.getenv("NOTIFICATION_EMAILS", "").split(",") if x.strip()]
    else:
        recipients = list(dict.fromkeys(x.strip() for x in recipients if x.strip()))
    if not sender or not app_password:
        raise ValueError("Configure SMTP_EMAIL and SMTP_APP_PASSWORD in .env.")
    if not recipients:
        raise ValueError("No recipients supplied for application report email.")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message.set_content(text_body)
    if html_body:
        message.add_alternative(html_body, subtype="html")

    if attachment_path:
        path = Path(attachment_path)
        if path.is_file():
            content_type, _ = mimetypes.guess_type(path.name)
            if content_type and "/" in content_type:
                maintype, subtype = content_type.split("/", 1)
            else:
                maintype, subtype = "application", "octet-stream"
            message.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)
        else:
            logger.warning("Application report attachment does not exist and was omitted: %s", path)

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(sender, app_password)
            smtp.send_message(message, to_addrs=recipients)
        logger.info("Application report email sent to %d recipient(s): %s", len(recipients), subject)
        return True
    except (smtplib.SMTPException, OSError) as exc:
        logger.exception("Application report email delivery failed.")
        raise RuntimeError(f"Could not send application report email: {exc}") from exc
