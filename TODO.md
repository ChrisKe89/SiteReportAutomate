(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/scripts/firmware_webforms_replay.py
Traceback (most recent call last):
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 266, in <module>
    main()
    ~~~~^^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 246, in main
    code, html = post_search(
                 ~~~~~~~~~~~^
        client, item.get("opco") or DEFAULT_OPCO, product, serial
        ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    )
    ^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 173, in post_search
    hidden = get_state(client)
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 166, in get_state
    r.raise_for_status()
    ~~~~~~~~~~~~~~~~~~^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_models.py", line 829, in raise_for_status
    raise HTTPStatusError(message, request=request, response=self)
httpx.HTTPStatusError: Client error '401 Unauthorized' for url 'https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx'
For more information check: https://developer.mozilla.org/en-US/docs/Web/HTTP/Status/401