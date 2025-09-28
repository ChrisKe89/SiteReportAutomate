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

<select name="ctl00$MainContent$ddlOpCoCode" onchange="javascript:setTimeout('__doPostBack(\'ctl00$MainContent$ddlOpCoCode\',\'\')', 0)" id="MainContent_ddlOpCoCode" class="yellowBg">
		<option selected="selected" value="">ALL</option>
		<option value="FXAU">FBAU</option>
		<option value="FXCA">FBCA</option>
		<option value="FXID">FBCN</option>
		<option value="FXHK">FBHK</option>
		<option value="FXKR">FBKR</option>
		<option value="FXMM">FBMM</option>
		<option value="FXMY">FBMY</option>
		<option value="FXNZ">FBNZ</option>
		<option value="FXPH">FBPH</option>
		<option value="FXSG">FBSG</option>
		<option value="THFX">FBTH</option>
		<option value="FXTW">FBTW</option>
		<option value="FXVN">FBVN</option>
		<option value="TWSI">TWSI</option>

	</select>
js document.querySelector("#MainContent_ddlOpCoCode")
//*[@id="MainContent_ddlOpCoCode"]

<input type="submit" name="ctl00$MainContent$btnSearch" value="Search" id="MainContent_btnSearch" class="btn btn-small btn-info button-small">
document.querySelector("#MainContent_btnSearch")
//*[@id="MainContent_btnSearch"]
/html/body/form/div[3]/div[3]/div[3]/div/div/div[2]/table/tbody/tr[8]/td[2]/input[1]


element <input type="image" name="ctl00$MainContent$btnExport" id="MainContent_btnExport" src="images/xls.png" style="border: none; background: inherit;">
js document.querySelector("#MainContent_btnExport")
//*[@id="MainContent_btnExport"]
Full xpath /html/body/form/div[3]/div[3]/div[4]/input