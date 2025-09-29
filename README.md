# EPGW-BusinessRules

Row-by-row rule creator driven by an Excel sheet.
Automates login monitoring, creates rules in the UI, and writes a Completed ledger with timestamps and status.

Script: EPGW-BusinessRules.py

Works best on Windows with Edge/Chromium + Playwright persistent profile

Input: RulesToCreate.xlsx (one row = one rule)

Output: RulesCompleted.xlsx (append-only), optional mutation of input

Logs: .\logs\, Auth snapshots: auth_status.json + auth-*.png

Site: http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx

---

0) Prereqs (Windows)

1. Python 3.10+ in PATH (python --version).


2. VS Code (optional, recommended).


3. Edge installed (or Chromium; script prefers Edge).




---

1) First-time Setup

> Run these in PowerShell from the project folder (e.g., C:\Dev\EPGW-BusinessRules).



1. Create venv + install deps:

python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install playwright openpyxl
python -m playwright install chromium

(Optional for lint/type hints): pip install types-openpyxl


2. Open the script and set site details (top of EPGW-BusinessRules.py):

BASE_URL, RULES_URL

selectors like SEL_LOGIN_SENTINEL, SEL_NEW_RULE_BTN, SEL_RULE_NAME, etc.

If your environment uses Windows Integrated Auth (IWA), keep the ALLOWLIST domain pattern (e.g., *.yourcorp.local).
If it uses form login, you can implement the form flow later (see “When Credentials Change”, #6).



3. Prepare your input workbook:

Create RulesToCreate.xlsx with a header row. Minimum columns:

RuleName, RuleType

Optional: Extra1, Extra2 (or any others you’ll map in the script)


Example:

RuleName	RuleType	Extra1	Extra2

Block colour copy	Policy	A	B
Enforce stapling	Action		C






---

2) Test a Dry Run (headed)

> Headed mode is friendliest for IWA and selector tuning.



.\.venv\Scripts\Activate.ps1
python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx

What you should see:

Browser launches (Edge), navigates to RULES_URL

Script checks for a login sentinel (SEL_LOGIN_SENTINEL); writes:

auth_status.json (logged_in true/false + URL + timestamp)

auth-YYYYMMDD-HHMMSS.png (screenshot)


For each row, it clicks New, fills fields, saves, waits for success indicator.

A row is appended to RulesCompleted.xlsx with ProcessedAt + Status.

Logs are written under .\logs\run-*.log.


If the page keeps some connections open (ASP.NET, SignalR), the script tolerates it—no tweaks needed.


---

3) Daily Use

Normal run (append results, keep input untouched):

python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx

Limit to first N rows (testing):

python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx --max-rows 3

Mutate input: remove successful rows from RulesToCreate.xlsx after processing:

python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx --mutate-input

Headless (only if IWA/form auth is solid in headless):

python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx --headless



---

4) Folder Outputs (what appears)

.\logs\run-*.log – step-by-step log per run

.\auth_status.json – last login check result

.\auth-*.png – screenshots proving session state

RulesCompleted.xlsx – append-only ledger (ProcessedAt, Status columns auto-added)

(if you use the report script too) .\downloads\ and .\user-data\ (browser profile)



---

5) Scheduling (Task Scheduler)

Create a tiny wrapper run_rules.ps1:

$ErrorActionPreference = 'Stop'
$AppDir = 'C:\Dev\EPGW-BusinessRules'   # <- your folder
$Py     = Join-Path $AppDir '.venv\Scripts\python.exe'
$Script = Join-Path $AppDir 'EPGW-BusinessRules.py'
$LogDir = Join-Path $AppDir 'logs'
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$ts  = Get-Date -Format 'yyyyMMdd-HHmmss'
$log = Join-Path $LogDir "rules-$ts.log"

Set-Location $AppDir
& $Py $Script --input "$AppDir\RulesToCreate.xlsx" --completed "$AppDir\RulesCompleted.xlsx" *>> $log

Task Scheduler → Create Task…

General:

Name: EPGW-BusinessRules

Run only when user is logged on (recommended if using IWA/headed)

Run with highest privileges


Triggers: your schedule (e.g., Daily 07:30).

Actions:

Program: powershell.exe

Args: -NoLogo -NoProfile -ExecutionPolicy Bypass -File "C:\Dev\EPGW-BusinessRules\run_rules.ps1"

Start in: C:\Dev\EPGW-BusinessRules


Settings: stop after 15 min; “If already running: Do not start a new instance”.


If you confirm headless auth is reliable, you can flip to Run whether user is logged on or not and add --headless.


---

6) When Credentials Change (read this first)

There are two common auth modes. The script supports both patterns.

A) Windows Integrated Auth (Kerberos/NTLM)

Used for intranet sites. Auth piggybacks your Windows session.

The script launches a persistent Edge profile (.\user-data) with:

--auth-server-allowlist=*.yourcorp.local
--auth-negotiate-delegate-allowlist=*.yourcorp.local

When your Windows password changes: nothing to do in the script. Your Windows login handles it.

If the site changes domain/host: update ALLOWLIST pattern and RULES_URL.

If the stored browser profile is stale (odd cached state): stop all runs, delete .\user-data folder, re-run (headed) once to rebuild the session.


B) Form Login (username/password, maybe MFA)

Add a small login flow to ensure_logged_in():

# Pseudocode inside ensure_logged_in, before checking SEL_LOGIN_SENTINEL
await page.goto(RULES_URL, wait_until="networkidle")
if await page.query_selector(SEL_LOGIN_SENTINEL):
    # already logged in
    ...
else:
    await page.fill(SEL_USERNAME, os.environ["APP_USER"])
    await page.fill(SEL_PASSWORD, os.environ["APP_PASS"])
    await page.click(SEL_LOGIN_BTN)
    # handle MFA here if needed (pause, wait for code field, etc.)
    await page.wait_for_selector(SEL_LOGIN_SENTINEL, timeout=20000)

Store creds in Windows Credential Manager or .env (with care):

Windows CredMan (recommended): use pywin32/keyring or just mandate an interactive login once and rely on a persistent cookie.

.env (quick):

APP_USER=someone
APP_PASS=secret

and read with os.getenv.


When credentials rotate:

Update the source of truth (CredMan or .env).

Re-run the script headed once to refresh any cookies.

If the login page changed IDs or flow, update SEL_* selectors.




---

7) Mapping Spreadsheet → Form

By default, the script expects columns:

RuleName → SEL_RULE_NAME

RuleType → SEL_RULE_TYPE

Optionals: Extra1, Extra2 (add your own fields)


Add any additional field mappings in create_rule_from_row():

await page.fill(SEL_RULE_NAME, name)
await page.fill(SEL_RULE_TYPE, rtype)
# Example extras:
# await page.fill(SEL_RULE_SCOPE, str(row.get("Scope", "")))
# await page.select_option(SEL_RULE_PRIORITY, str(row.get("Priority", "Medium")))


---

8) Troubleshooting

Login fails: check auth_status.json + the latest auth-*.png.

For IWA: confirm you can browse the URL manually in Edge as the same Windows user.

If still failing headless, run headed and schedule “Run only when user is logged on”.


Selectors moved: open DevTools, re-grab IDs, update SEL_*.

Workbook issues:

RulesCompleted.xlsx is append-only. If you want a fresh ledger, archive or delete it.

--mutate-input only removes rows with Status == "SUCCESS".


Long-running pages (SignalR/WebSockets): the script already tolerates missing full idle by adding short waits after actions.



---

9) Quick Commands (clipboard-friendly)

# Activate env
.\.venv\Scripts\Activate.ps1

# Run (no input mutation)
python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx

# Run and remove successful rows from input
python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx --mutate-input

# Limit to N rows for testing
python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx --max-rows 5

# Headless (only if auth is solid)
python .\EPGW-BusinessRules.py --input .\RulesToCreate.xlsx --completed .\RulesCompleted.xlsx --headless


---

10) What to send me next

When you’re ready:

The real RULES_URL

The final selectors for: login sentinel, new rule button, each field, save button, success toast

A sample row (or the live workbook headers) if you want me to wire the exact mappings


I’ll plug those in and hand back the ready-to-run version.

