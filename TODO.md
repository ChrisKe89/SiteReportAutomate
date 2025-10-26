(.venv) PS C:\Dev\SiteReportAutomate> & C:/Dev/SiteReportAutomate/.venv/Scripts/python.exe c:/Dev/SiteReportAutomate/scripts/firmware_webforms_replay.py
Traceback (most recent call last):
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_transports\default.py", line 101, in map_httpcore_exceptions
    yield
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_transports\default.py", line 250, in handle_request
    resp = self._pool.handle_request(req)
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpcore\_sync\connection_pool.py", line 256, in handle_request
    raise exc from None
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpcore\_sync\connection_pool.py", line 236, in handle_request
    response = connection.handle_request(
        pool_request.request
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpcore\_sync\connection.py", line 101, in handle_request
    raise exc
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpcore\_sync\connection.py", line 78, in handle_request
    stream = self._connect(request)
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpcore\_sync\connection.py", line 156, in _connect
    stream = stream.start_tls(**kwargs)
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpcore\_backends\sync.py", line 154, in start_tls
    with map_exceptions(exc_map):
         ~~~~~~~~~~~~~~^^^^^^^^^
  File "C:\Users\au016207\AppData\Local\Programs\Python\Python313\Lib\contextlib.py", line 162, in __exit__
    self.gen.throw(value)
    ~~~~~~~~~~~~~~^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpcore\_exceptions.py", line 14, in map_exceptions
    raise to_exc(exc) from exc
httpcore.ConnectError: [SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: unable to get local issuer certificate (_ssl.c:1028)

The above exception was the direct cause of the following exception:

Traceback (most recent call last):
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 256, in <module>
    main()
    ~~~~^^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 236, in main
    code, html = post_search(
                 ~~~~~~~~~~~^
        client, item.get("opco") or DEFAULT_OPCO, product, serial
        ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    )
    ^
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 164, in post_search
    hidden = get_state(client)
  File "c:\Dev\SiteReportAutomate\scripts\firmware_webforms_replay.py", line 156, in get_state
    r = client.get(URL, headers={"User-Agent": HEADERS["User-Agent"]})
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_client.py", line 1053, in get
    return self.request(
           ~~~~~~~~~~~~^
        "GET",
        ^^^^^^
    ...<7 lines>...
        extensions=extensions,
        ^^^^^^^^^^^^^^^^^^^^^^
    )
    ^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_client.py", line 825, in request
    return self.send(request, auth=auth, follow_redirects=follow_redirects)
           ~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_client.py", line 914, in send
    response = self._send_handling_auth(
        request,
    ...<2 lines>...
        history=[],
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_client.py", line 942, in _send_handling_auth
    response = self._send_handling_redirects(
        request,
        follow_redirects=follow_redirects,
        history=history,
    )
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_client.py", line 979, in _send_handling_redirects
    response = self._send_single_request(request)
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_client.py", line 1014, in _send_single_request
    response = transport.handle_request(request)
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_transports\default.py", line 249, in handle_request
    with map_httpcore_exceptions():
         ~~~~~~~~~~~~~~~~~~~~~~~^^
  File "C:\Users\au016207\AppData\Local\Programs\Python\Python313\Lib\contextlib.py", line 162, in __exit__
    self.gen.throw(value)
    ~~~~~~~~~~~~~~^^^^^^^
  File "C:\Dev\SiteReportAutomate\.venv\Lib\site-packages\httpx\_transports\default.py", line 118, in map_httpcore_exceptions
    raise mapped_exc(message) from exc
httpx.ConnectError: [SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: unable to get local issuer certificate (_ssl.c:1028)