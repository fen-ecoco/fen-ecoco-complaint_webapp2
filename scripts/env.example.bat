@echo off
rem ============================================================
rem  Environment template - copy to env.local.bat and fill in.
rem  env.local.bat is gitignored and never committed.
rem  Safer alternative: set these as Windows system environment
rem  variables so no plaintext credentials sit on disk.
rem  Keep this file ASCII-only (see run_analysis.bat).
rem ============================================================

rem -- Google service account JSON, full content on one line --
rem set GOOGLE_CREDENTIALS_JSON={"type":"service_account", ...}

rem -- or point at a JSON file instead --
rem set GOOGLE_APPLICATION_CREDENTIALS=D:\info\0318_test\service-account.json

rem -- History spreadsheet (blank = built-in default) --
rem set HISTORY_SHEET_ID=1Sqh_8bXtFw7jvmCPufTpStKxfIafDzwYJRlgc0HFBSs

rem -- Source spreadsheet the scheduler reads.
rem    Leave SOURCE_SHEET_ID blank to watch a folder instead (see below). --
rem set SOURCE_SHEET_ID=
rem set SOURCE_WORKSHEET=

rem -- Folder watched when SOURCE_SHEET_ID is blank (default: inbox) --
rem set ECOCO_INBOX=\\server\share\complaints_inbox

rem -- Optional: L2 LLM classification and AI reports --
rem set ANTHROPIC_API_KEY=

rem -- Behaviour knobs (all have defaults) --
rem set REVIEW_CONFIDENCE_THRESHOLD=0.75
rem set AUDIT_SAMPLE_RATE=0.03
rem set HISTORY_READONLY=false
