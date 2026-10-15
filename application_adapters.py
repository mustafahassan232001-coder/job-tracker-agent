"""Extension point for platform-authorized application integrations.

This starter deliberately registers no job-board submission adapters. Add an
adapter only after verifying that the platform offers and permits the method.
Never automate around CAPTCHAs, access controls, or platform restrictions.

An adapter can implement:
    supports(job: dict) -> bool
    submit(job: dict, cv_path: Path, cover_letter_path: Path) -> dict

The submit result should be one of:
    {"status": "submitted", "confirmation": "...", "notes": "..."}
    {"status": "submission_failed", "notes": "..."}
"""
from pathlib import Path
from typing import Any

AUTHORIZED_ADAPTERS: list[Any] = []


def try_submit_with_authorized_adapter(job: dict[str, Any], cv_path: Path, cover_letter_path: Path) -> dict[str, Any] | None:
    for adapter in AUTHORIZED_ADAPTERS:
        try:
            if adapter.supports(job):
                return adapter.submit(job, cv_path, cover_letter_path)
        except Exception:
            # Caller logs and routes this application to manual handling.
            raise
    return None
