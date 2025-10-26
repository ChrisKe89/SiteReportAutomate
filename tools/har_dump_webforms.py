# save as tools\har_dump_webforms.py
import json
import zipfile
from pathlib import Path

HAR_ZIP = Path("logs/firmware_lookup.har.zip")
TARGET = "SingleRequest.aspx"

with zipfile.ZipFile(HAR_ZIP, "r") as z:
    with z.open("har.har") as f:
        har = json.load(f)

for e in har.get("log", {}).get("entries", []):
    req = e.get("request", {})
    res = e.get("response", {})
    url = req.get("url", "")
    if TARGET not in url or req.get("method") != "POST" or res.get("status") != 200:
        continue

    headers = {h["name"].lower(): h["value"] for h in req.get("headers", [])}
    post = req.get("postData", {})
    mime = (post.get("mimeType") or "").lower()
    print("\n=== MATCH ===")
    print("URL:", url)
    print("Content-Type:", mime)
    print("Request Headers (subset):")
    for k in ("cookie", "origin", "referer", "user-agent"):
        if k in headers:
            print(f"  {k}: {headers[k]}")

    if "text" in post:
        print("\nForm Body (raw):")
        print(post["text"][:4000])  # print first 4k chars
    elif "params" in post:
        print("\nForm Params:")
        for p in post["params"]:
            print(f"  {p.get('name')} = {p.get('value')}")
