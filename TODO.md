(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/scripts/firmware_webforms_replay_playwright.py
Traceback (most recent call last):
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 506, in <module>
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
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 476, in main
    code_c, status_c = await dom_submit_schedule(
                       ^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<7 lines>...
    )
    ^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 393, in dom_submit_schedule
    await page.evaluate("""
    ...<8 lines>...
    """, SCHEDULE_TRIGGER)
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py", line 8514, in evaluate
    await self._impl_obj.evaluate(
        expression=expression, arg=mapping.to_impl(arg)
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_page.py", line 468, in evaluate
    return await self._main_frame.evaluate(expression, arg)
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_frame.py", line 320, in evaluate
    await self._channel.send(
    ...<6 lines>...
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 69, in send
    return await self._connection.wrap_api_call(
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<3 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 558, in wrap_api_call
    raise rewrite_error(error, f"{parsed_st['apiName']}: {error}") from None
playwright._impl._errors.Error: Page.evaluate: TypeError: 'caller', 'callee', and 'arguments' properties may not be accessed on strict mode functions or the arguments objects for calls to them
    at get arguments (<anonymous>)
    at Sys.WebForms.PageRequestManager._doPostBack (https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/ScriptResource.axd?d=dwY9oWetJoJoVpgL6Zq8OKdTQFwDxS6qr-AdfsdxzUjg4Pm0BKWaIyx4BofwpYTAdhikyr4IEVBSktLaR4w5GD8DMjb9ihUuUoehrfwfi1bBeHGgGSnz6aHj9YbbF8B8o91vQDktWb7ZjE_3F0TXH91LngXlNpsUPsKA5_6-8yI1&t=32e5dfca:5:13771)
    at https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/ScriptResource.axd?d=NJmAwtEo3Ipnlaxl6CMhvqU_Xi_L5WBslZueKMd7KjK2d1bYvA2kvR82kfMfyVBUeOIL2OUJdQPMFNeBuODHdYTxYomuSHTEgc82IISyTMV5pIdZ4iTGOVCz9AdLqb7h_4OQ1roKOOdcdJXKzLHpfUMebIhnPsGSFhh_RbYDXgI1&t=32e5dfca:5:307
    at eval (eval at evaluate (:291:30), <anonymous>:3:24)
    at UtilityScript.evaluate (<anonymous>:298:18)
    at UtilityScript.<anonymous> (<anonymous>:1:44)