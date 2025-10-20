# EP Gateway Business Rule Bot

## Version

Current release: **0.1.3** (2024-05-23)


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

## Firmware scheduling automation

The repository now includes **`schedule_firmware.py`** to automate software upgrade
requests in the Fuji Xerox Single Request portal.

### Configure inputs

Add the following keys to your `.env` (or rely on the defaults shown):

```env
FIRMWARE_INPUT_XLSX=downloads/VIC.xlsx
FIRMWARE_LOG_XLSX=downloads/FirmwareLog.xlsx
FIRMWARE_STORAGE_STATE=storage_state.json
FIRMWARE_BROWSER_CHANNEL=msedge
FIRMWARE_ERRORS_JSON=errors.json
FIRMWARE_HTTP_USERNAME=
FIRMWARE_HTTP_PASSWORD=
FIRMWARE_WARMUP_URL=http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx
```

* `FIRMWARE_INPUT_XLSX` must point to a worksheet that contains the columns
  **`SerialNumber`**, **`Product_Code`**, **`OpcoID`**, and **`State`**.
* `FIRMWARE_LOG_XLSX` receives the results for each processed row. Successful
  schedules, skips (already upgraded/not eligible), and failures are all
  appended with timestamps.
* `FIRMWARE_STORAGE_STATE` should reference a login state captured via
  `login_capture_epgw.py` so the script can reuse your authenticated session.
  The scheduler now **requires** that this file exists and will prompt you to
  capture a session if it cannot be found.
* `FIRMWARE_BROWSER_CHANNEL` defaults to Microsoft Edge. Change it to `chromium`
  or `chrome` if you prefer another browser build installed on your system.
* `FIRMWARE_ERRORS_JSON` is a JSON ledger that captures failed or skipped rows
  (for example, when a device table is missing). Leave it at the default or
  point it to a shared location if multiple team members are running the bot.
* `FIRMWARE_HTTP_USERNAME` / `FIRMWARE_HTTP_PASSWORD` are optional HTTP Auth
  credentials for environments where the Single Request portal prompts for a
  login dialog before cookies are considered valid.
* `FIRMWARE_WARMUP_URL` lets the scheduler visit a gateway page (captured with
  `login_capture_epgw.py`) before loading the firmware scheduling form. Keep the
  default or clear it if your environment does not require the warm-up hop.

### Run the scheduler

```bash
python schedule_firmware.py
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

`ast_toner.py` drives the Fuji Xerox parts status portal using Playwright.

* **Input**: The newest `.xlsx` file inside `downloads/` (skipping temporary `~$` files). Override with the `AST_TONER_INPUT` environment variable if you need a specific workbook.
* **Output**: `output.xlsx` populated with the input columns and fetched status rows.
* **Resilience**: missing columns are treated as blanks, empty rows are skipped, and
  timeouts/no-table responses are logged to both the console and output workbook.
* **Authentication**: set `AST_TONER_STORAGE_STATE` to point at a specific `storage_state.json`. If it is omitted the script falls back to a local `storage_state.json` when present, otherwise it launches without persisted cookies and you will need to log in interactively.

Run it with the same authenticated `storage_state.json` created for the main bot:

```bash
python ast_toner.py
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
