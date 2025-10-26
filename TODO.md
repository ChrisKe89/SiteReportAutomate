(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/scripts/firmware_webforms_replay.py     
Traceback (most recent call last):
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 277, in <module>
    main()
    ~~~~^^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 257, in main
    code, html = post_search(
                 ~~~~~~~~~~~^
        client, item.get("opco") or DEFAULT_OPCO, product, serial
        ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    )
    ^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 185, in post_search
    hidden = get_state(client)
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 177, in get_state
    raise PermissionError("Unauthorized (401): cookies may be expired or invalid.")
PermissionError: Unauthorized (401): cookies may be expired or invalid.