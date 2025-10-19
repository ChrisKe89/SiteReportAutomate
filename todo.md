(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/login_capture_epgw.py
Traceback (most recent call last):
  File "c:\Dev\SiteReportAutomate\login_capture_epgw.py", line 26, in <module>
    asyncio.run(main())
    ~~~~~~~~~~~^^^^^^^^
  File "C:\Users\au016207\AppData\Local\Programs\Python\Python313\Lib\asyncio\runners.py", line 195, in run
    return runner.run(main)
           ~~~~~~~~~~^^^^^^
  File "C:\Users\au016207\AppData\Local\Programs\Python\Python313\Lib\asyncio\runners.py", line 118, in run
    return self._loop.run_until_complete(task)
           ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~^^^^^^
  File "C:\Users\au016207\AppData\Local\Programs\Python\Python313\Lib\asyncio\base_events.py", line 725, in run_until_complete
    return future.result()
           ~~~~~~~~~~~~~^^
  File "c:\Dev\SiteReportAutomate\login_capture_epgw.py", line 15, in main
    await page.goto(LOGIN_URL, wait_until="networkidle")
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py", line 8992, in goto
    await self._impl_obj.goto(
        url=url, timeout=timeout, waitUntil=wait_until, referer=referer
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_page.py", line 556, in goto
    return await self._main_frame.goto(**locals_to_params(locals()))
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_frame.py", line 153, in goto
    await self._channel.send(
        "goto", self._navigation_timeout, locals_to_params(locals())
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 69, in send
    return await self._connection.wrap_api_call(
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<3 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 558, in wrap_api_call
    raise rewrite_error(error, f"{parsed_st['apiName']}: {error}") from None
playwright._impl._errors.Error: Page.goto: net::ERR_INVALID_AUTH_CREDENTIALS at http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx
Call log:
  - navigating to "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx", waiting until "networkidle"
