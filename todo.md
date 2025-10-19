(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/schedule_firmware.py        
Warm-up navigation to http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx failed: Page.goto: net::ERR_INVALID_AUTH_CREDENTIALS at http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx
Call log:
  - navigating to "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx", waiting until "domcontentloaded"

Authentication failed when opening the firmware scheduling page. Re-capture storage_state.json after a successful manual login or set FIRMWARE_HTTP_USERNAME/FIRMWARE_HTTP_PASSWORD in your .env.