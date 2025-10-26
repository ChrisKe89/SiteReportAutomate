(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/scripts/firmware_webforms_replay_playwright.py
Done. Wrote: data\firmware_schedule_out.csv
Task exception was never retrieved
future: <Task finished name='Task-5' coro=<Page.wait_for_selector() done, defined at C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py:8110> exception=TargetClosedError('Page.wait_for_selector: Target page, context or browser has been closed\nCall log:\n  - waiting for locator("#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus") to be visible\n')>
Traceback (most recent call last):
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py", line 8181, in wait_for_selector
    await self._impl_obj.wait_for_selector(
        selector=selector, timeout=timeout, state=state, strict=strict
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_page.py", line 423, in wait_for_selector
    return await self._main_frame.wait_for_selector(**locals_to_params(locals()))
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_frame.py", line 369, in wait_for_selector
    await self._channel.send(
        "waitForSelector", self._timeout, locals_to_params(locals())
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 69, in send
    return await self._connection.wrap_api_call(
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<3 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 558, in wrap_api_call
    raise rewrite_error(error, f"{parsed_st['apiName']}: {error}") from None
playwright._impl._errors.TargetClosedError: Page.wait_for_selector: Target page, context or browser has been closed
Call log:
  - waiting for locator("#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus") to be visible

Task exception was never retrieved
future: <Task finished name='Task-7' coro=<Page.wait_for_selector() done, defined at C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py:8110> exception=TargetClosedError('Page.wait_for_selector: Target page, context or browser has been closed\nCall log:\n  - waiting for locator("#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus") to be visible\n')>
Traceback (most recent call last):
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py", line 8181, in wait_for_selector
    await self._impl_obj.wait_for_selector(
        selector=selector, timeout=timeout, state=state, strict=strict
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_page.py", line 423, in wait_for_selector
    return await self._main_frame.wait_for_selector(**locals_to_params(locals()))
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_frame.py", line 369, in wait_for_selector
    await self._channel.send(
        "waitForSelector", self._timeout, locals_to_params(locals())
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 69, in send
    return await self._connection.wrap_api_call(
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<3 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 558, in wrap_api_call
    raise rewrite_error(error, f"{parsed_st['apiName']}: {error}") from None
playwright._impl._errors.TargetClosedError: Page.wait_for_selector: Target page, context or browser has been closed
Call log:
  - waiting for locator("#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus") to be visible
