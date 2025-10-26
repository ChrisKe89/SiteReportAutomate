# EP Automate

Automation utilities for EP Gateway, firmware scheduling, and related toner/reporting workflows. Every tool is designed to run from this repository root, reads configuration from `.env`, and relies on Playwright for browser automation.

## Setup
1. **Clone & enter the repo**
   ```powershell
   git clone https://github.com/ChrisKe89/SiteReportAutomate.git
   cd .\SiteReportAutomate
   ```
2. **Create a virtual environment & install deps**
   ```powershell
   python -m venv .venv
   . .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   python -m playwright install msedge
   ```
3. **Sync environment files**
   ```powershell
   pwsh -File .\sync-dotenv.ps1
   ```
   Rerun the script whenever `.env.example` gains new keys (use `-DryRun` first if you only want to diff).

## .env Variables
- `.env.example` documents every supported key. Copy/merge it into `.env` via `sync-dotenv.ps1`.
- Paths accept absolute, relative, or UNC locations. Defaults point inside `data/`, `logs/`, or `downloads/`.
- Core sections:
  - **Report fetch (`FETCH_*`, `REPORT_OUTPUT_XLSX`)** – controls EP report scraping, download directories, and browser behavior.
  - **Firmware scheduler (`FIRMWARE_*`)** – inputs, outputs, concurrency, and browser storage for firmware automation.
  - **AST toner (`AST_*`)** – workbook locations plus optional RDHC snapshot overrides.
- Update keys directly in `.env`; scripts call `load_dotenv()` so no code changes are required.

## Login Capture
Use these interactive helpers once per account (or whenever NTLM/SSO cookies expire):

### EP Firmware Login Capture
`python scripts\login_capture\login_capture_remote_firmware.py`

- Opens Edge/Chromium via Playwright, navigates to `SingleRequest.aspx`, and pauses so you can authenticate.
- Saves cookies to `storage_state.json` (path configurable via `--storage-state` or `FIRMWARE_STORAGE_STATE`).
- Reuse the same state with firmware scheduling and AST toner scripts.

### EP Gateway Login Capture
`python scripts\login_capture\login_capture_epgw.py`

- Targets the EP Gateway warm-up URL defined in the script (or override via CLI).
- Follow the prompts, finish any MFA, then press **Enter** so the context can be persisted for downstream tasks.

## EP Report
`python scripts\ep_report\fetch_and_clean.py`

- Downloads the device list report, converts the HTML-in-XLS payload into a clean XLSX, and moves it to `REPORT_OUTPUT_XLSX`.
- Key environment toggles:
  - `FETCH_BASE_URL`, `FETCH_REPORT_URL` – endpoints to visit.
  - `FETCH_DOWNLOAD_DIR`, `FETCH_USER_DATA_DIR` – working folders for Playwright.
  - `FETCH_HEADLESS`, `FETCH_AUTH_ALLOWLIST`, `FETCH_NAV_TIMEOUT_MS`, `FETCH_AFTER_SEARCH_WAIT_MS` – browser/session behavior.
- Resulting spreadsheet feeds other automations such as AST toner or firmware scheduling.

## EP Firmware
`python scripts\schedule_firmware\firmware_webforms_replay_playwright.py`

- Reads rows from `FIRMWARE_INPUT_XLSX` (CSV/XLSX), performs Search + Schedule inside the firmware portal, and writes per-row outcomes.
- Environment highlights:
  - `FIRMWARE_OPCO`, `FIRMWARE_INPUT_XLSX`, `FIRMWARE_STORAGE_STATE`, `FIRMWARE_BROWSER_CHANNEL`, `FIRMWARE_AUTH_ALLOWLIST`, `FIRMWARE_HEADLESS`.
  - Scheduling knobs: `FIRMWARE_TIME_VALUE`, `FIRMWARE_DAYS_MIN`, `FIRMWARE_DAYS_MAX`, `FIRMWARE_DEBUG_TZ`.
  - Throughput: `FIRMWARE_CONCURRENCY`.
  - **Output control:** `FIRMWARE_OUTPUT_CSV` (default `data/firmware_schedule_out.csv`). The script also creates a timestamped copy (e.g. `firmware_schedule_out_20250130-103000.csv`).
- Behavior:
  - Each worker removes its completed/skipped row from `FIRMWARE_INPUT_XLSX` when the source is CSV.
  - `run_started_at` / `run_completed_at` columns mark the execution window.
  - Errors are logged inline per device row for downstream triage.

## AST Toner
`python scripts\ast_toner\fetch_ast_toner.py`

- Uses report data (`AST_INPUT_XLSX`) to query the RDHC toner portal and exports summaries to `AST_OUTPUT_CSV`.
- Column mapping is configurable through env vars (`PRODUCT_FAMILY_COLUMN`, etc.).
- `RDHC.html` is loaded from the repo root by default (or override with `RDHC_HTML_PATH`) to map product families to dropdown values.
- Supply `AST_TONER_STORAGE_STATE`/`AST_BROWSER_CHANNEL`/`AST_HEADLESS` as needed; failures are logged with helpful context.

## EP Business Rule
TBA – this section will be populated once the business rule automation is reinstated.
