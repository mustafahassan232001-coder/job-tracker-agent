# Job Search & Application Assistant

A configurable Python/LangGraph project that searches for job leads, opens and classifies candidate URLs, assesses fit against a CV, categorizes vacancies by role family, saves the results in SQLite and Excel, sends separate category emails for stronger matches, and sends a final lower-fit analysis with a learning plan.

This repository is a generic template. It intentionally contains no personal CVs, real email addresses, API keys, SMTP passwords, databases, or generated application documents.

## Features

- Extracts text from one or more selectable-text PDF CVs and creates a separate candidate profile for each configured CV.
- Runs broad regional searches and additional targeted searches for preferred sites such as LinkedIn Jobs, Indeed, Bayt, GulfTalent, Naukrigulf, Laimoon, Foundit Gulf, Dubizzle, Oliv, and Jora.
- Uses SearXNG's JSON endpoint with Yahoo configured as the selected search engine. Search coverage depends on the search engine's index and rate limits; there is no guarantee that all jobs or websites are discoverable.
- Verifies URLs where possible and distinguishes individual vacancies from search pages, training/course content, informational resources, advertising/service pages, closed pages, and ambiguous or blocked pages.
- Categorizes confirmed/provisional vacancy pages by job-title family, then assesses fit against the active candidate profile.
- Sends one email per strong-fit role category, with a category-specific ZIP of available tailored CV/cover-letter files.
- Sends a separate final career-gap email for lower-fit vacancy leads and places uncertain URLs in a manual-review section.
- Creates a cumulative Excel report with `All Discovered Leads`, `Verified & Scored`, `Needs Manual Review`, and `Excluded Non-Jobs` sheets. Job URLs are clickable.
- Stores search and application tracking data in SQLite.
- Keeps automatic application submission disabled by default. Preparing a document is not the same as submitting an application.

## Requirements

- Python 3.11 or newer
- Docker Desktop (for the optional local SearXNG service)
- A Google Gemini API key
- SMTP credentials if you want email delivery

## Setup (Windows PowerShell)

1. Clone/download this repository and open PowerShell in its folder.
2. Create and activate a virtual environment:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

3. Copy `.env.example` to `.env`. Add your Gemini key, SMTP credentials, CV file paths, recipients, and desired region searches. Do not commit `.env`.

4. Put your own CV PDFs in the project folder, or set paths in `CV_PATHS`. Do not put private CVs in source control.

5. Configure SearXNG:

   ```powershell
   Copy-Item .\searxng\settings.example.yml .\searxng\settings.yml
   ```

   Replace `REPLACE_WITH_A_RANDOM_PRIVATE_VALUE` with a randomly generated value and keep `settings.yml` out of Git. Start SearXNG:

   ```powershell
   docker compose -f docker-compose.searxng.yml up -d
   ```

   Check that `http://localhost:8888/search?q=test&format=json` returns JSON before running the tracker. SearXNG engines may fail, rate-limit or return CAPTCHA pages.

6. Check syntax, run the offline tests, then start the workflow:

   ```powershell
   .\.venv\Scripts\python.exe -m py_compile main.py application_manager.py email_notifier.py
   .\.venv\Scripts\python.exe -m unittest discover -s tests -v
   .\.venv\Scripts\python.exe main.py
   ```

## Configuration

Most settings are in `.env.example`. Important ones:

- `CV_PATHS`: comma-separated PDF paths. The workflow processes each CV independently.
- `MY_CV_EMAILS`: email recipients for the primary CV.
- `SECOND_CV_EMAIL`: recipients for the CV whose filename matches any keyword in `SECOND_CV_KEYWORDS`.
- `SEARCH_REGIONS_JSON`: JSON object mapping region labels to search terms. Example:

  ```dotenv
  SEARCH_REGIONS_JSON='{"Central Europe":"Czechia OR Germany OR Poland","Gulf":"UAE OR Saudi Arabia OR Qatar"}'
  ```

- `MAX_RESULTS_PER_LINKEDIN_QUERY`: requested result count for each LinkedIn targeted search query. This is not a guaranteed number of unique jobs.
- `STRONG_FIT_MIN_SCORE`: minimum score for strong-fit category emails.
- `APPLICATION_MIN_FIT_SCORE`: minimum score for preparing tailored application documents.
- `MAX_APPLICATION_PACKAGES_PER_RUN`: cap on new application packages generated per run.
- `VERIFY_JOB_URLS`: enables URL verification.
- `AUTO_APPLY_ENABLED`: default is `false`; only explicitly implemented and authorized adapters can submit applications.

If you use more than two CVs, review the candidate recipient routing in `main.py` and `application_manager.py` before using it. This template exposes separate primary and secondary recipient settings and isn't a full multi-user account system.

## Output

The default `OUTPUT_DIR=output` contains candidate-specific SQLite databases where configured, cumulative Excel files, application folders, document metadata, and application-package ZIPs. The entire output folder is ignored by Git because it can contain personal details.

## Caveats

- URL verification is evidence-based, not a guarantee that a vacancy is still open. Sites may block requests or require JavaScript/login; such pages should be reviewed manually.
- Match scores are estimates based on the text available to the model. Search snippets may be incomplete or stale.
- Some job boards prohibit automation or restrict access. This project uses indexed search results and a visible-page check; it does not bypass logins or access controls.
- Review generated CVs and cover letters before sending them. The model should not invent skills, certifications, employers, dates or achievements.
- Ensure you are allowed to process each person's CV and keep candidate records, API keys, SMTP settings and generated documents private.

## Security before publishing

- Copy `.env.example` to `.env` locally; never commit `.env`.
- Never commit real CVs, application files, spreadsheets, databases, screenshots of credentials, or SearXNG's private `settings.yml`.
- If a real secret was ever committed, deleting it from a later commit is not enough: rotate/revoke it and clean the repository history before publishing.
