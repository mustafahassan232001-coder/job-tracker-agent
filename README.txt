AI JOB TRACKER + APPLICATION DOCUMENT PREPARATION
=================================================

WHAT THIS PACKAGE DOES
----------------------
1. Searches broadly through DuckDuckGo-indexed websites, plus targeted queries for:
   LinkedIn, Indeed, Bayt, GulfTalent, Naukrigulf, Laimoon, Foundit Gulf, Dubizzle,
   Oliv, and Jora.
2. Analyzes each configured CV separately and scores discovered job leads.
3. Keeps each CV's job database and Excel report separate.
4. For each unprocessed, recommended job scoring at least APPLICATION_MIN_FIT_SCORE,
   generates a job-specific ATS-friendly CV (DOCX) and cover letter (DOCX).
5. Saves each application package locally and attaches a ZIP of newly generated
   documents to a combined application report email for that candidate.
6. Reports confirmed submissions separately from jobs that need manual application.
7. Sends a separate CV-specific career feedback email with job leads, recurring
   missing skills, and a learning plan.

IMPORTANT AUTOMATION LIMITATION
-------------------------------
This package does NOT yet contain a working submission adapter for LinkedIn or any
other job board. The application_adapters.py file is an extension point and starts
with an empty adapter list. Therefore, in this version, jobs are prepared for manual
submission; the code does not claim to have applied automatically.

AUTO_APPLY_ENABLED=true does not magically enable applications. A real, platform-
authorized adapter must first be implemented and registered in application_adapters.py.
Only use integrations and methods that the platform supports and permits. Do not
attempt to bypass login protections, CAPTCHAs, access controls, or site restrictions.
The tracker records a job as submitted only when a configured adapter explicitly
returns status="submitted".

DOCUMENT SAFETY
---------------
The model is instructed to tailor documents using only facts in the original CV.
However, AI-generated documents must still be reviewed for factual accuracy before
submission. The code will not invent skills intentionally, but a model can make
mistakes. Review the CV and cover letter before using them.

APPLICATION WORKFLOW
--------------------
- Each candidate has an independent SQLite database and applications folder.
- By default, only recommended jobs with a match score of 80+ are prepared.
- MAX_APPLICATION_PACKAGES_PER_RUN limits new packages to 10 per candidate per run
  by default, to control API usage. Remaining eligible jobs stay queued for later runs.
- A job is processed once after documents are prepared; preparation failures are retried.
- A combined application report email is sent for newly processed jobs. It includes
  confirmed submissions, jobs needing manual action, and preparation failures.
- The tailored CV and cover letter are attached in a ZIP when files were generated.
- Career feedback is a separate email for each CV.
- The older per-category job-digest emails are disabled by default to avoid duplicate
  job alerts. Set SEND_JOB_DIGEST_EMAILS=true to re-enable them.

INSTALL / UPDATE ON WINDOWS
---------------------------
1. Back up your current project folder.
2. Extract this package.
3. Copy main.py, email_notifier.py, application_manager.py, application_adapters.py,
   and requirements.txt into your existing Job_tracker_project folder.
4. Do NOT replace your existing .env. Keep secrets in that local file only.
5. In the activated virtual environment, install any missing dependencies:
      pip install -r requirements.txt
6. Add or review the settings below in your existing .env.
7. Run syntax checks:
      python -m py_compile main.py email_notifier.py application_manager.py application_adapters.py
8. Run the tracker:
      python main.py

RECOMMENDED .ENV SETTINGS
-------------------------
# Keep your existing GOOGLE_API_KEY, SMTP_EMAIL, SMTP_APP_PASSWORD, SMTP_HOST,
# SMTP_PORT, CV_PATH/CV_PATHS, OUTPUT_DIR, and recipient settings.

CV_PATH=my_cv.pdf
CV_PATHS=my_cv.pdf,Ahmed_Mohamed_Hassan.pdf
OUTPUT_DIR=output
GEMINI_MODEL=gemini-3.1-flash-lite
MAX_RESULTS_PER_REGION=10
MAX_RESULTS_PER_SITE=10
EMAIL_MIN_FIT_SCORE=85
EMAIL_EVERY_N_RUNS=1
SEND_CAREER_FEEDBACK_EMAIL=true
FEEDBACK_MIN_FIT_SCORE=60
SEND_JOB_DIGEST_EMAILS=false
APPLICATION_MIN_FIT_SCORE=80
MAX_APPLICATION_PACKAGES_PER_RUN=10
AUTO_APPLY_ENABLED=false
MY_CV_EMAILS=your_email@example.com
AHMED_CV_EMAIL=ahmedzidan1281993@gmail.com

Set MY_CV_EMAILS to your own recipient address(es), comma-separated. Configure
AHMED_CV_EMAIL for Ahmed. Use the actual filenames in CV_PATHS; the Ahmed detection
uses AHMED_CV_KEYWORDS if configured. Do not put API keys or passwords in this README.

OUTPUTS
-------
- output/job_history_<cvname>.xlsx
- output/jobs_<cvname>.db (when multiple CVs are configured)
- output/applications/<cvname>/<company>_<job>_<id>/tailored_cv.docx
- output/applications/<cvname>/<company>_<job>_<id>/cover_letter.docx
- output/application_packages_<cvname>_run_<n>.zip

SEARCH LIMITATIONS
------------------
“All websites” means broad searches across websites indexed by DuckDuckGo plus
additional targeted domain searches. It does not guarantee every website or vacancy
is found. LinkedIn is searched through publicly indexed pages only. Search results may
be stale, may lead to search pages rather than individual vacancies, and do not prove
a vacancy is still open.

LIVE TESTING
------------
Python syntax and document-writing helpers can be tested locally, but live job search,
Gemini calls, and SMTP delivery require your own working environment and credentials.
Never paste your .env secrets into chat or commit them to a repository.
