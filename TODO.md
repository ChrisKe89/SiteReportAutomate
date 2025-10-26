.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/scripts/firmware_webforms_replay_playwright.py
Traceback (most recent call last):
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 602, in <module>
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
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 569, in main
    code_c, status_c = await dom_submit_schedule(
                       ^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<7 lines>...
    )
    ^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 489, in dom_submit_schedule
    msg = await read_status()
          ^^^^^^^^^^^^^^^^^^^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 480, in read_status
    el = await page.query_selector(sel)
         ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py", line 8087, in query_selector
    await self._impl_obj.query_selector(selector=selector, strict=strict)
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_page.py", line 411, in query_selector
    return await self._main_frame.query_selector(selector, strict)
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_frame.py", line 348, in query_selector
    await self._channel.send("querySelector", None, locals_to_params(locals()))
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 69, in send
    return await self._connection.wrap_api_call(
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<3 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 558, in wrap_api_call
    raise rewrite_error(error, f"{parsed_st['apiName']}: {error}") from None
playwright._impl._errors.Error: Page.query_selector: Execution context was destroyed, most likely because of a navigation