# EPGW Business Rules Bot — Windows + VS Code

This tool reads rows from an **Excel workbook** and creates Business Rules in the EPGW web app using Playwright (Microsoft Edge automation).
Each row is validated; results are written to a **Completed** CSV/XLSX ledger; detailed logs go to `logs/`.
If you’re not logged in, the run stops and tells you why.

---

## 0) Install VS Code and basics (one-time)

1. **Install**

* **[VS Code](https://code.visualstudio.com/download)** (then open it)
* **[Python 3.12+](https://www.python.org/downloads/)** (tick “**Add Python to PATH**” during install)
* **[Git for Windows](https://git-scm.com/downloads/win)**

2. **Create a working folder**

```powershell
New-Item -ItemType Directory -Path C:\EPGW_Automation -Force
Set-Location C:\EPGW_Automation
```

3. **Clone the repo into that folder**

```powershell
git clone https://github.com/ChrisKe89/gwbusinessrules.git
cd .\gwbusinessrules
```

4. **Open the project in VS Code**

```powershell
code .
```

---

## 1) Create & select a Python environment

In VS Code: **Terminal → New Terminal** (PowerShell), then:

```powershell
python -m venv .venv
. .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

Then select the interpreter: **Ctrl + Shift + P → “Python: Select Interpreter” → .venv**.

---

## 2) Install dependencies

```powershell
pip install -r requirements.txt
python -m playwright install
python -m playwright install msedge
```

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

## 6) Run it (normal usage from VS Code)

With `.venv` active:

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