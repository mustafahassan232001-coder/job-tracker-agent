# Security Policy

## Handling secrets and personal data

- Never commit `.env`, API keys, SMTP passwords, CVs, generated application documents, spreadsheets, databases, logs, or local SearXNG settings.
- `.env.example` and `searxng/settings.example.yml` are templates only. Replace placeholders locally and keep the resulting files out of version control.
- If a credential is exposed in Git history, revoke or rotate it immediately. Removing it in a later commit does not make the old credential safe.
- Obtain permission before processing another person's CV. Treat CVs, generated documents, and job-search history as personal data.

## Reporting a security issue

Please do not post working credentials or private candidate data in a public issue. Report suspected issues privately to the repository maintainer using the contact method configured for the GitHub repository.
