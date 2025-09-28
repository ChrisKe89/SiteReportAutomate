# Site Report Bot

## Automate any report

1. **Download GIT REPO**

```bash
git clone [text](https://github.com/ChrisKe89/SITEREPORTBOT.git)
```
2. **Install Requirments**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install playwright python-dotenv keyring
python -m playwright install chromium
```

3. **Run it**

```powershell
python login_capture.py
```
- got to website and login
- generate your report 
- export


4. **Run
```powershell
python download_report.py

```
- You want to see it land a file in .\downloads\.
