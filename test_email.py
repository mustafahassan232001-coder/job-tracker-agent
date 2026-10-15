
from email_notifier import send_new_jobs_email

test_jobs = [
    {
        "title": "Junior Cloud Engineer - Test",
        "company": "Example Company",
        "location": "Brno, Czech Republic",
        "url": "https://example.com/job",
        "match_score": 85,
        "reason": "Test notification to verify email delivery.",
    }
]

send_new_jobs_email(test_jobs)
print("Test email sent successfully.")
