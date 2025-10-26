"""Scan the captured HAR for JSON-like endpoints."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

HAR_ZIP = Path("logs/firmware_lookup.har.zip")


def is_json_like(content_type: str | None) -> bool:
    if not content_type:
        return False
    ct = content_type.lower()
    return "application/json" in ct or "+json" in ct


def main() -> None:
    if not HAR_ZIP.exists():
        raise FileNotFoundError(
            f"HAR archive not found at {HAR_ZIP}. Run capture_har.py first."
        )

    with zipfile.ZipFile(HAR_ZIP, "r") as archive:
        with archive.open("har.har") as handle:
            har = json.load(handle)

    entries = har.get("log", {}).get("entries", [])
    candidates: list[dict[str, str | int | None]] = []

    for entry in entries:
        request = entry.get("request", {})
        response = entry.get("response", {})
        url = request.get("url", "")
        method = request.get("method", "")
        status = response.get("status", 0)
        content = response.get("content", {}) or {}
        content_type = content.get("mimeType")

        headers = request.get("headers", [])
        x_requested_with = any(
            (header.get("name", "").lower() == "x-requested-with") for header in headers
        )

        if not (is_json_like(content_type) or x_requested_with):
            continue

        body_preview = ""
        post_data = request.get("postData") or {}
        if method.upper() == "POST":
            params = post_data.get("params") or []
            if params:
                body_preview = "&".join(
                    f"{param.get('name')}={param.get('value')}" for param in params
                )
            elif "text" in post_data:
                body_preview = str(post_data.get("text", ""))[:500]

        candidates.append(
            {
                "status": status,
                "method": method,
                "url": url,
                "content_type": content_type,
                "post_preview": body_preview[:200],
            }
        )

    candidates.sort(key=lambda item: (item.get("status") != 200, item.get("url", "")))

    print("\n=== Likely JSON endpoints ===")
    for candidate in candidates:
        status = candidate.get("status")
        method = candidate.get("method")
        url = candidate.get("url")
        print(f"[{status}] {method} {url}")
        preview = candidate.get("post_preview")
        if preview:
            print(f"    body: {preview}")


if __name__ == "__main__":
    main()
