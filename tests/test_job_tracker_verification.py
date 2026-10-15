"""Offline helper tests for vacancy classification and LinkedIn query planning."""
import ast
import logging
import re
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse
import unittest

SOURCE_PATH = Path(__file__).resolve().parent.parent / "main.py"
SOURCE = SOURCE_PATH.read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)

assignments = {}
for node in TREE.body:
    if isinstance(node, ast.Assign):
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id in {"DEFAULT_REGIONS", "JOB_SITES"}:
                assignments[target.id] = ast.literal_eval(node.value)

# The production module loads region overrides from .env; this offline unit test
# uses the generic defaults, which are parsed as a literal from the source.
assignments["REGIONS"] = assignments["DEFAULT_REGIONS"]

wanted = {
    "looks_like_listing_page",
    "_has_role_title_signal",
    "_is_explicit_individual_job_url",
    "_provisional_job_fallback",
    "_is_obvious_learning_or_ad_page",
    "classify_opened_job_pages",
    "plan_search",
}
segments = [ast.get_source_segment(SOURCE, node) for node in TREE.body
            if isinstance(node, ast.FunctionDef) and node.name in wanted]
ns = {
    "Any": Any,
    "JobSearchState": dict,
    "re": re,
    "parse_qs": parse_qs,
    "unquote": unquote,
    "urlparse": urlparse,
    "normalize_url": lambda value: str(value or "").strip().rstrip("/").split("#")[0],
    "REGIONS": assignments["REGIONS"],
    "JOB_SITES": assignments["JOB_SITES"],
    "logger": logging.getLogger("test_job_tracker_verification"),
}
exec("\n\n".join(segments), ns)


class VerificationTests(unittest.TestCase):
    def test_linkedin_single_posting_url_is_not_listing(self):
        job = {
            "title": "NOC Engineer at Example Company",
            "url": "https://www.linkedin.com/jobs/view/noc-engineer-at-example-1234567890",
            "description": "Monitor network alerts; investigate incidents; requirements include TCP/IP and Linux.",
        }
        self.assertFalse(ns["looks_like_listing_page"](job))
        self.assertTrue(ns["_is_explicit_individual_job_url"](job))
        self.assertTrue(ns["_has_role_title_signal"](job))

    def test_linkedin_search_and_career_pages_are_not_individual_jobs(self):
        examples = [
            {"title": "Jobs | LinkedIn", "url": "https://www.linkedin.com/jobs/search/?keywords=cloud"},
            {"title": "LinkedIn Careers", "url": "https://www.linkedin.com/company/linkedin/"},
            {"title": "LinkedIn Careers", "url": "https://www.linkedin.com/careers/"},
        ]
        for job in examples:
            with self.subTest(url=job["url"]):
                self.assertTrue(ns["looks_like_listing_page"](job))
                self.assertFalse(ns["_is_explicit_individual_job_url"](job))

    def test_blocked_linkedin_individual_posting_is_provisionally_scoreable(self):
        job = {
            "title": "Junior Cloud Engineer at Example Company",
            "url": "https://www.linkedin.com/jobs/view/1234567890",
            "description": "Junior Cloud Engineer. Azure, Linux, and CI/CD are preferred.",
        }
        result = ns["_provisional_job_fallback"](job, job["url"], "HTTP 403")
        self.assertEqual(result["verification_status"], "likely_job")
        self.assertIn("not confirmation", result["verification_reason"])

    def test_course_and_training_pages_are_excluded(self):
        job = {
            "title": "DevOps Engineer Certification Course",
            "url": "https://www.linkedin.com/jobs/view/1234567890",
            "description": "Enroll in this course and receive a certificate of completion.",
        }
        rejected, reason = ns["_is_obvious_learning_or_ad_page"](job)
        self.assertTrue(rejected, reason)
        fallback = ns["_provisional_job_fallback"](job, job["url"], "HTTP 403")
        self.assertEqual(fallback["verification_status"], "review_required")

    def test_classifier_accepts_explicit_blocked_job_url_but_not_course(self):
        linkedin_job = {
            "title": "NOC Analyst at Example Company",
            "url": "https://www.linkedin.com/jobs/view/9876543210",
            "description": "NOC Analyst: monitor incidents, escalate faults, manage network alerts.",
            "verification_status": "likely_job",
            "verification_reason": "Explicit individual-vacancy URL pattern and role-specific title/snippet; automated page content was unavailable (HTTP 403).",
        }
        course = {
            "title": "DevOps Training Course",
            "url": "https://www.linkedin.com/jobs/view/1122334455",
            "description": "Enroll in this training course to prepare for certifications.",
            "verification_status": "review_required",
            "verification_reason": "HTTP 403",
        }
        result = ns["classify_opened_job_pages"]({"opened_pages": [linkedin_job, course], "not_opened_jobs": []})
        self.assertEqual([item["url"] for item in result["verified_jobs"]], [linkedin_job["url"]])
        self.assertEqual([item["url"] for item in result["excluded_pages"]], [course["url"]])

    def test_linkedin_searches_are_split_by_region_and_role_family(self):
        profile = {
            "target_roles": ["Cloud Engineer", "DevOps Engineer", "Network Engineer", "IT Specialist", "Systems Administrator"],
            "keywords": ["Azure", "Terraform", "Linux", "GitHub"],
        }
        tasks = ns["plan_search"]({"profile": profile})["search_tasks"]
        linkedin = [item for item in tasks if item.get("site_domain") == "linkedin.com"]
        self.assertEqual(len(linkedin), len(assignments["REGIONS"]) * 3)
        self.assertTrue(all("site:linkedin.com/jobs/view" in item["query"] for item in linkedin))
        queries = "\n".join(item["query"].lower() for item in linkedin)
        self.assertIn("noc engineer", queries)
        self.assertIn("network operations engineer", queries)
        self.assertIn("junior cloud engineer", queries)
        # Broad-web queries and additional targeted searches for other preferred sites remain enabled.
        self.assertTrue(any(not item.get("site_domain") for item in tasks))
        for site in assignments["JOB_SITES"]:
            if site["domain"] != "linkedin.com":
                self.assertTrue(any(item.get("site_domain") == site["domain"] for item in tasks), site["domain"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
