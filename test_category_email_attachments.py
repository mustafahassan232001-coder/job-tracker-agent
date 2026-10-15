"""Offline test that category job emails really include their ZIP attachment."""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
import ast
import hashlib
import logging
import re
import sqlite3
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import patch
import zipfile
from typing import Any

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import email_notifier  # noqa: E402


class _FakeSMTP:
    def __init__(self, *args, **kwargs):
        self.message = None

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def ehlo(self):
        return None

    def starttls(self):
        return None

    def login(self, *args):
        return None

    def send_message(self, message, to_addrs=None):
        self.message = message
        _FakeSMTP.last_message = message


class CategoryAttachmentTests(unittest.TestCase):
    def test_category_email_attaches_category_zip(self):
        with tempfile.TemporaryDirectory() as tmp:
            zip_path = Path(tmp) / "category_documents_IT_Specialist_run_1.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("Example_1234/tailored_cv.docx", b"test cv")
                archive.writestr("Example_1234/cover_letter.docx", b"test letter")
            env = {
                "SMTP_EMAIL": "sender@example.com",
                "SMTP_APP_PASSWORD": "dummy-test-password",
                "SMTP_HOST": "localhost",
                "SMTP_PORT": "587",
            }
            job = {
                "title": "IT Operations Specialist",
                "company": "Example Company",
                "location": "Riyadh",
                "fit_score": 85,
                "url": "https://www.linkedin.com/jobs/view/1234567890",
                "job_category": "IT Specialist & General IT",
                "verification_status": "likely_job",
            }
            with patch.dict(os.environ, env, clear=False), patch.object(email_notifier.smtplib, "SMTP", _FakeSMTP):
                sent = email_notifier.send_new_jobs_email(
                    [job],
                    recipients=["recipient@example.com"],
                    subject="IT Specialist strong-fit jobs",
                    attachment_paths=[zip_path],
                )
            self.assertTrue(sent)
            message = _FakeSMTP.last_message
            self.assertIsInstance(message, EmailMessage)
            attachments = [part.get_filename() for part in message.iter_attachments()]
            self.assertIn(zip_path.name, attachments, attachments)
            self.assertIn("ATTACHED APPLICATION DOCUMENTS", message.get_body(preferencelist=("plain",)).get_content())

    def test_no_attachment_argument_preserves_legacy_use(self):
        env = {
            "SMTP_EMAIL": "sender@example.com",
            "SMTP_APP_PASSWORD": "dummy-test-password",
            "SMTP_HOST": "localhost",
            "SMTP_PORT": "587",
        }
        with patch.dict(os.environ, env, clear=False), patch.object(email_notifier.smtplib, "SMTP", _FakeSMTP):
            sent = email_notifier.send_new_jobs_email(
                [{"title": "NOC Engineer", "url": "https://example.com/noc"}],
                recipients=["recipient@example.com"],
                subject="NOC jobs",
            )
        self.assertTrue(sent)
        self.assertEqual(list(_FakeSMTP.last_message.iter_attachments()), [])


class CategoryZipBuilderTests(unittest.TestCase):
    def test_category_zip_contains_only_qualifying_jobs_in_that_category(self):
        source = (HERE / "main.py").read_text(encoding="utf-8")
        tree = ast.parse(source)
        function_node = next(
            node for node in tree.body
            if isinstance(node, ast.FunctionDef) and node.name == "build_category_application_zip"
        )
        code = ast.get_source_segment(source, function_node)
        namespace = {
            "Any": Any, "Path": Path, "sqlite3": sqlite3, "re": re,
            "hashlib": hashlib, "zipfile": zipfile, "logger": logging.getLogger("test_zip_builder"),
            "normalize_url": lambda value: str(value or "").strip().rstrip("/").split("#")[0],
            "_safe_attachment_slug": lambda value, limit=70: (re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip()).strip("._-") or "job")[:limit],
            "APPLICATION_MIN_FIT_SCORE": 80,
        }
        exec(code, namespace)
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            db_path = tmp_path / "jobs.db"
            out_dir = tmp_path / "output"
            out_dir.mkdir()
            cv_file = tmp_path / "candidate.pdf"
            cv_file.write_bytes(b"dummy")
            high_cv = tmp_path / "cv_high.docx"
            high_letter = tmp_path / "letter_high.docx"
            low_cv = tmp_path / "cv_low.docx"
            low_letter = tmp_path / "letter_low.docx"
            for path in (high_cv, high_letter, low_cv, low_letter):
                path.write_bytes(b"test document")
            with sqlite3.connect(db_path) as con:
                con.execute("CREATE TABLE applications (url TEXT PRIMARY KEY, cv_path TEXT, cover_letter_path TEXT)")
                con.execute("INSERT INTO applications VALUES (?, ?, ?)", ("https://example.com/job/high", str(high_cv), str(high_letter)))
                con.execute("INSERT INTO applications VALUES (?, ?, ?)", ("https://example.com/job/low", str(low_cv), str(low_letter)))
            namespace.update({"DB_PATH": db_path, "OUTPUT_DIR": out_dir, "CV_PATH": cv_file})
            result = namespace["build_category_application_zip"]("Network & NOC Engineering", [
                {"title": "NOC Engineer", "company": "Example", "fit_score": 85, "url": "https://example.com/job/high"},
                {"title": "Network Trainee", "company": "Example", "fit_score": 75, "url": "https://example.com/job/low"},
            ], 1)
            self.assertIsNotNone(result)
            with zipfile.ZipFile(result) as archive:
                names = archive.namelist()
            self.assertEqual(len(names), 2, names)
            self.assertTrue(any("tailored_cv.docx" in name for name in names), names)
            self.assertTrue(any("cover_letter.docx" in name for name in names), names)
            self.assertFalse(any("low" in name.lower() for name in names), names)


if __name__ == "__main__":
    unittest.main(verbosity=2)
