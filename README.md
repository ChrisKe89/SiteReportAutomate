# EP Gateway Business Rule Bot

## Version

Current release: **0.1.7** (2025-10-22)


> [!NOTE]  
>
>This tool reads rows from an **Excel workbook** and creates Business Rules in the EPGW web app using Playwright (Microsoft Edge automation).
>
>Each row is validated; results are written to a **Completed** CSV/XLSX ledger; detailed logs go to `logs/`.
>
>If you’re not logged in, the run stops and tells you why.


### **First Time Use**
---
> [!IMPORTANT]
> 
> **Install**
> * **[VS Code](https://code.visualstudio.com/download)** (then open it)
> * **[Python 3.12+](https://www.python.org/downloads/)** (tick “**Add Python to PATH**” during install)
> * **[Git for Windows](https://git-scm.com/downloads/win)**
> 
> 1. **Create a working folder**
> 
> ```powershell
> New-Item -ItemType Directory -Path C:\Dev\sitereportautomate -Force
> Set-Location C:\Dev
> ```
> 
> 3. **Clone the repo into that folder**
> 
> ```powershell
> git clone https://github.com/ChrisKe89/SiteReportAutomate.git
> cd .\sitereportautomate
> ```
> 
> 4. **Open the project in VS Code**
> 
> ```powershell
> code .
> ```

---

### 1) Create & select a Python environment

In VS Code: **Terminal → New Terminal** (PowerShell), then:

```powershell
python -m venv .venv
. .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

Then select the interpreter: **Ctrl + Shift + P → “Python: Select Interpreter” → .venv**.

---

### 2) Install dependencies

```bash
pip install -r requirements.txt
python -m playwright install
python -m playwright install msedge
```

> [!TIP]
> If activation is blocked:
> `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`

---

## 3) Configure environment variables

The script reads its input/output file paths from a **.env** file.

1. **Review `.env.example`**
   It contains sample keys like:

   ```
   INPUT_XLSX=.\EPGW-BusinessRules-template1.xlsx
   COMPLETED_XLSX=.\Completed.xlsx
   COMPLETED_CSV=.\RulesCompleted.csv
   ```

   Edit these defaults if you need different locations (UNC paths like `\\server\share\Rules.xlsx` work fine).

2. **Create or update `.env` from the example**
   The repo includes **sync-dotenv.ps1** to handle this.

   * First-time create:

     ```powershell
     pwsh -File .\sync-dotenv.ps1
     ```
   * Update existing `.env` with any new keys (keeping your values):

     ```powershell
     pwsh -File .\sync-dotenv.ps1
     ```
   * Force full overwrite (backs up current file):

     ```powershell
     pwsh -File .\sync-dotenv.ps1 -ForceCopy
     ```
   * Preview changes without writing:

     ```powershell
     pwsh -File .\sync-dotenv.ps1 -DryRun
     ```

   Your final `.env` should live at the repo root.
   The main script automatically loads it (no code edits required).

---

## 4) One-time login capture (stores your session)

```powershell
python .\login_capture_epgw.py
```

An Edge window opens on the EPGW page. Sign in (SSO/NTLM/MFA).
When the Business Rules page is fully loaded, return to the terminal and press **Enter**.
A `storage_state.json` file is created and reused on later runs.

---

## 5) Prepare your Excel input

* Use the provided template (recommended) or your own workbook.
* **Single worksheet** is assumed.
* The file may be local or on a network share.

**Required headers (case-sensitive):**

* Core: `BusinessRuleName`, `AlertType`, `OpCo`
* Status: `BusinessRuleStatus` (`Enabled` / `Disabled`)
* Additional Conditions:

  * `AdditionalConditionsMode`: `DayOfMonth` | `DayOfWeek` | `FromDate` (blank = Any Day/Time)
  * If `DayOfMonth`: `DayOfMonth`
  * If `FromDate`: `FromDate`, `ToDate`, `FromHour`, `FromMinute`, `ToHour`, `ToMinute`
  * If `DayOfWeek`: `WeekDay::Monday` … `WeekDay::Sunday` with `Yes` for selected days
* Consolidation:

  * `ConsolidationMode`: `ByDays` | `ByWeek`
  * If `ByDays`: `ConsolidateDays` (positive integer)
  * If `ByWeek`: `ConsolidationWeek::Monday` … `::Sunday` with `Yes` for selected days
* Toner (when `AlertType` is `EXCH`, `SE/F`, `SNE/F`, or `SNEFGC`):

  * Columns like `TonerMain::BlackTonerCartridge` or `TonerOther::WasteTonerContainer` with `Yes`

Rows that miss required combinations are **SKIPPED** and the reason is recorded in the Completed ledger.

---

## Download the firmware device list

Use `fetch_and_clean.py` to download the Device List report, clean the HTML
spreadsheet, and place the result where the other scripts expect it.

* The cleaned workbook is written to the path defined by
  `REPORT_OUTPUT_XLSX` (default `data/EPFirmwareReport.xlsx`). The script
  always moves the latest download into this location after conversion.
* The raw HTML-in-XLS downloads remain in `downloads/` with timestamped names
  so you can archive or audit them later.

```bash
python fetch_and_clean.py
```

## Firmware scheduling automation

The repository now includes **`schedule_firmware.py`** to automate software upgrade
requests in the Fuji Xerox Single Request portal.

### Configure inputs

Add the following keys to your `.env` (or rely on the defaults shown):

```env
FIRMWARE_INPUT_XLSX=data/firmware_schedule.csv
FIRMWARE_LOG_XLSX=logs/fws_log.json
FIRMWARE_STORAGE_STATE=storage_state.json
FIRMWARE_BROWSER_CHANNEL=msedge
FIRMWARE_ERRORS_JSON=logs/fws_error_log.json
FIRMWARE_HTTP_USERNAME=
FIRMWARE_HTTP_PASSWORD=
FIRMWARE_AUTH_ALLOWLIST=*.fujixerox.net,*.xerox.com
FIRMWARE_WARMUP_URL=http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx
```

* `FIRMWARE_INPUT_XLSX` may reference either a `.xlsx` workbook or a `.csv`
  file. It must contain the columns **`SerialNumber`**, **`Product_Code`**,
  **`OpcoID`**, and **`State`**.
* `FIRMWARE_LOG_XLSX` accepts any path. When the extension is `.json` the
  scheduler appends structured log entries to that JSON array. Otherwise an
  Excel workbook is created and updated using the same columns.
* `FIRMWARE_STORAGE_STATE` should reference a login state captured via
  `login_capture_epgw.py` so the script can reuse your authenticated session.
  The scheduler now **requires** that this file exists and will stop with a
  helpful error if it cannot be found.
* `FIRMWARE_BROWSER_CHANNEL` defaults to the bundled Chromium build (empty
  value). Set it to `msedge`, `chromium`, or `chrome` if you prefer another
  locally installed browser.
* `FIRMWARE_HEADLESS` controls whether the scheduler runs without a visible
  window. It now defaults to `true` so unattended automation remains silent.
* `FIRMWARE_ERRORS_JSON` is a JSON ledger that captures failed or skipped rows
  (for example, when a device table is missing). Leave it at the default or
  point it to a shared location if multiple team members are running the bot.
* `FIRMWARE_HTTP_USERNAME` / `FIRMWARE_HTTP_PASSWORD` are optional HTTP Auth
  credentials for environments where the Single Request portal prompts for a
  login dialog before cookies are considered valid.
* `FIRMWARE_AUTH_ALLOWLIST` controls which hosts Microsoft Edge will
  automatically challenge with your Windows credentials. The default now covers
  both the firmware portal (`*.fujixerox.net`) and the gateway warm-up hop
  (`*.xerox.com`). Extend it if your infrastructure fronts additional domains.
* `FIRMWARE_WARMUP_URL` lets the scheduler visit a gateway page (captured with
  `login_capture_epgw.py`) before loading the firmware scheduling form. Keep the
  default or clear it if your environment does not require the warm-up hop.

### Run the scheduler

```bash
python schedule_firmware.py
```

To run the automation in a forced headless browser session, use the dedicated
entry point:

```bash
python scripts/schedule_firmware_headless.py
```

For each Excel row the bot:

1. Selects **FBAU** in the OpCo dropdown.
2. Submits the product code and serial number.
3. Skips rows that report `already upgraded` in the eligibility grid.
4. When eligible, chooses a random schedule date within the next six days,
   selects an allowed time slot (00:00–07:59 or 18:00–23:59), and picks the
   correct timezone based on the State column (NT → Darwin, SA → Adelaide,
   ACT/VIC/NSW → Canberra/Melbourne/Sydney, QLD → Brisbane, TAS → Hobart).
5. Clicks **Schedule** and records the portal’s confirmation message in the
   firmware log workbook.

Any exceptions encountered while processing a row are also logged so you can
review and retry later.

Failures and critical skips are additionally written to `errors.json` (or your
custom `FIRMWARE_ERRORS_JSON` path) to make quick triage easier.

> **Tip:** If the script reports `ERR_INVALID_AUTH_CREDENTIALS`, re-run
> `login_capture_epgw.py` to refresh `storage_state.json` or configure the HTTP
> credentials in `.env` so the browser can satisfy the portal's authentication
> challenge automatically.

## 6) Run it (normal usage from VS Code)

With `.venv` active:

---

## Toner status automation companion script

`fetch_ast_toner.py` drives the RDHC parts status portal using Playwright.

* **Input**: `AST_INPUT_XLSX` (defaults to `data/EPFirmwareReport.xlsx`). The
  script reads the serial number from column **A**, the product code from
  column **B**, and the product family from column **G**.
* **Dropdown mapping**: the product family string is matched against the
  options in `RDHC.html` so the correct value is selected when submitting the
  form. Update the HTML snapshot if the portal introduces new families.
* **Output**: `AST_OUTPUT_CSV` receives one row per lookup with the serial,
  product code, product family, and the text extracted from
  `MainContent_UpdatePanelResult`.
* **Authentication**: set `AST_TONER_STORAGE_STATE` to point at the same
  `storage_state.json` captured for the other scripts. When the file is missing
  the script still launches but warns that a fresh login may be required.

## Firmware HAR utilities

Use the helper scripts at the repository root to capture and analyse the
network traffic generated during a manual firmware lookup:

```bash
python capture_har.py
```

Follow the prompts to perform a single lookup in the browser window. The HAR is
saved to `logs/firmware_lookup.har.zip`. To review likely JSON endpoints inside
the capture, run:

```bash
python scan_har.py
```

The scanner prints a sorted list of candidate API calls, including preview
snippets for any POST bodies that were recorded.

Run it with the same authenticated storage state created for the main bot:

```bash
python fetch_ast_toner.py
```

```powershell
python .\EPGW-BusinessRules.py `
  --input templates\EPGW-BusinessRules-template.xlsx `
  --completed .\Completed.xlsx `
  --completed-csv .\RulesCompleted.csv `
  --mutate-input
  ```

Because `.env` supplies all paths, no other arguments are needed.

What happens:

* Reads `INPUT_XLSX` from `.env`
* Writes successes to `COMPLETED_XLSX` / `COMPLETED_CSV`
* Removes every **successful** row from the input after the run (`--mutate-input`)
* Logs full details to `logs/run-YYYYMMDD-HHMMSS.log`

Optional extras:

* `--max-rows 5` → test a few rows
* `-v` → more detailed logging (DEBUG) in terminal and log file

---

## 7) Refresh the repo (pull updates)

In VS Code **Source Control** view:

* **… → Pull**
  or run:

```powershell
git pull origin main
```

If requirements changed:

```powershell
. .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install
```

---

## 8) Schedule it (optional)

Create a Basic Task in **Windows Task Scheduler**:

```powershell
-NoProfile -ExecutionPolicy Bypass -Command "Set-Location 'C:\EPGW_Automation\gwbusinessrules'; . .\.venv\Scripts\Activate.ps1; python .\epgw-businessrules.py --mutate-input"
```

If cookies expire, just re-run **login_capture_epgw.py** to refresh `storage_state.json`.

---

## 9) Troubleshooting

* **Input not found** → check the path in `.env`.
* **Login not established / Aborting** → session expired or selector changed; re-run login capture.
* **Rows SKIPPED** → a required combo is missing (see headers above).
* **Edge not found** → run `python -m playwright install msedge`.

---

## Changing file locations later

1. Edit `.env` in VS Code and update:

   ```
   INPUT_XLSX=\\newserver\newshare\NewRules.xlsx
   COMPLETED_XLSX=\\newserver\newshare\Completed.xlsx
   COMPLETED_CSV=\\newserver\newshare\Completed.csv
   ```
2. Save and rerun `sync-dotenv.ps1` (optional) to ensure any new keys are present.

No code changes are needed—the script always reads paths from `.env`.

---
