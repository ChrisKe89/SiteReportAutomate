# tools/diag_auth_insecure.py
from __future__ import annotations

import json
import os
from pathlib import Path

import httpx

BASE = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"
URL = f"{BASE}/firmware/SingleRequest.aspx"

STORAGE_STATE_PATH = Path(os.getenv("FIRMWARE_STORAGE_STATE", "storage_state.json"))


def import_cookies(client: httpx.Client) -> None:
    if not STORAGE_STATE_PATH.exists():
        raise FileNotFoundError(f"Missing storage_state.json: {STORAGE_STATE_PATH}")
    payload = json.loads(STORAGE_STATE_PATH.read_text(encoding="utf-8"))
    for c in payload.get("cookies", []):
        name = c.get("name")
        value = c.get("value")
        domain = c.get("domain") or "sgpaphq-epbbcs3.dc01.fujixerox.net"
        path = c.get("path") or "/"
        if name and value:
            client.cookies.set(name, value, domain=domain, path=path)


def main() -> None:
    # verify=False is **only** for this diagnostic
    with httpx.Client(
        follow_redirects=False, timeout=30.0, trust_env=True, verify=False
    ) as client:
        import_cookies(client)

        # show which cookies will go out to the host
        jar = [
            c
            for c in client.cookies.jar
            if "fujixerox.net" in (getattr(c, "domain", "") or "")
        ]
        print("Cookies loaded for host:")
        for c in jar:
            domain = getattr(c, "domain", "") or "<unknown>"
            name = getattr(c, "name", "") or "<unnamed>"
            value_preview = (getattr(c, "value", "") or "")[:6]
            path = getattr(c, "path", "") or "/"
            print(f"  {domain} {name}={value_preview}…; path={path}")

        r = client.get(URL, headers={"User-Agent": "Mozilla/5.0"})
        print(f"\nGET {URL} -> {r.status_code}")
        print(f"WWW-Authenticate: {r.headers.get('www-authenticate', '')!r}")
        print(f"Set-Cookie: {'<present>' if r.headers.get('set-cookie') else '<none>'}")


if __name__ == "__main__":
    main()
