(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/scripts/firmware_webforms_replay_playwright.py
Traceback (most recent call last):
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 262, in <module>
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
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 238, in main
    code, html = await post_search(
                 ^^^^^^^^^^^^^^^^^^
        page, hidden, item.get("opco") or DEFAULT_OPCO, product, serial
        ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    )
    ^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay_playwright.py", line 182, in post_search
    resp = await page.request.post(URL, headers=HEADERS, data=payload)
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\async_api\_generated.py", line 18613, in post
    await self._impl_obj.post(
    ...<11 lines>...
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_fetch.py", line 280, in post
    return await self.fetch(
           ^^^^^^^^^^^^^^^^^
    ...<12 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_fetch.py", line 319, in fetch
    return await self._inner_fetch(
           ^^^^^^^^^^^^^^^^^^^^^^^^
    ...<13 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_fetch.py", line 410, in _inner_fetch
    response = await self._channel.send(
               ^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<18 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 69, in send
    return await self._connection.wrap_api_call(
           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    ...<3 lines>...
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\playwright\_impl\_connection.py", line 558, in wrap_api_call
    raise rewrite_error(error, f"{parsed_st['apiName']}: {error}") from None
playwright._impl._errors.Error: APIRequestContext.post: unable to get local issuer certificate
Call log:
  - → POST https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx
    - user-agent: Mozilla/5.0
    - accept: */*
    - accept-encoding: gzip,deflate,br
    - X-Requested-With: XMLHttpRequest
    - X-MicrosoftAjax: Delta=true
    - Content-Type: application/x-www-form-urlencoded; charset=UTF-8
    - Origin: https://sgpaphq-epbbcs3.dc01.fujixerox.net
    - Referer: https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx
    - content-length: 2148
    - cookie: ASP.NET_SessionId=h20aygi4mvd15kf0htnx4k05; __AntiXsrfToken=195e7c91efde47ba83926573385bdc9d
