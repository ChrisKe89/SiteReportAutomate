(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/tools/diag_auth.py
Cookies loaded for host:
  sgpaphq-epbbcs3.dc01.fujixerox.net ASP.NET_SessionId=h20ayg…; path=/
  sgpaphq-epbbcs3.dc01.fujixerox.net __AntiXsrfToken=195e7c…; path=/

GET https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx -> 401
WWW-Authenticate: 'Negotiate, NTLM'
Set-Cookie: <none>


(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/tools/diag_auth_insecure.py
Cookies loaded for host:
  sgpaphq-epbbcs3.dc01.fujixerox.net ASP.NET_SessionId=h20ayg…; path=/
  sgpaphq-epbbcs3.dc01.fujixerox.net __AntiXsrfToken=195e7c…; path=/

GET https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx -> 401
WWW-Authenticate: 'Negotiate, NTLM'
Set-Cookie: <none>