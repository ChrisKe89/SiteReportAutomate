"""Capture POST payloads made during a manual firmware lookup."""

from __future__ import annotations

import asyncio
import json
import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv  # type: ignore[import]
from playwright.async_api import Error as PlaywrightError  # type: ignore[import]
from playwright.async_api import Request, async_playwright  # type: ignore[import]

from scripts.playwright_launch import launch_browser


def _env_path(var_name: str, default: str) -> Path:
    raw_value = os.getenv(var_name, default)
    normalised = raw_value.replace("\\", "/")
    return Path(normalised).expanduser()


load_dotenv()

URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx"
OUT = Path("logs/post_payloads.ndjson")
BROWSER_CHANNEL = os.getenv("FIRMWARE_BROWSER_CHANNEL", "").strip() or None
ALLOWLIST = os.getenv("FIRMWARE_AUTH_ALLOWLIST", "*.fujixerox.net,*.xerox.com")
AUTH_WARMUP_URL = os.getenv(
    "FIRMWARE_WARMUP_URL",
    "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx",
)
STORAGE_STATE_PATH = _env_path("FIRMWARE_STORAGE_STATE", "storage_state.json")
HTTP_USERNAME = os.getenv("FIRMWARE_HTTP_USERNAME")
HTTP_PASSWORD = os.getenv("FIRMWARE_HTTP_PASSWORD")


async def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)

    storage_state = STORAGE_STATE_PATH if STORAGE_STATE_PATH.exists() else None
    context_kwargs: dict[str, Any] = {}
    if HTTP_USERNAME and HTTP_PASSWORD:
        context_kwargs["http_credentials"] = {
            "username": HTTP_USERNAME,
            "password": HTTP_PASSWORD,
        }

    browser_args = [
        f"--auth-server-allowlist={ALLOWLIST}",
        f"--auth-negotiate-delegate-allowlist={ALLOWLIST}",
    ]

    async with async_playwright() as playwright:
        browser, context = await launch_browser(
            playwright,
            headless=False,
            channel=BROWSER_CHANNEL,
            storage_state_path=storage_state,
            browser_args=browser_args,
            context_kwargs=context_kwargs,
        )
        try:
            captured: list[dict[str, Any]] = []

            def on_request(request: Request) -> None:
                try:
                    if request.method == "POST" and "SingleRequest.aspx" in request.url:
                        body = request.post_data or ""
                        captured.append(
                            {
                                "url": request.url,
                                "method": request.method,
                                "headers": request.headers,
                                "post_data": body,
                                "resource_type": request.resource_type,
                            }
                        )
                except Exception:
                    # Ignore serialization issues so recording can continue
                    pass

            context.on("request", on_request)

            page = await context.new_page()
            page.set_default_navigation_timeout(45_000)
            page.set_default_timeout(45_000)

            if AUTH_WARMUP_URL:
                try:
                    await page.goto(AUTH_WARMUP_URL, wait_until="domcontentloaded")
                except PlaywrightError as exc:
                    print(f"Warm-up navigation to {AUTH_WARMUP_URL} failed: {exc}")

            await page.goto(URL, wait_until="domcontentloaded")

            print("\n> Do ONE normal lookup on the page (fill & submit).")
            print("> When the result appears, return here and press Enter.")
            input()

            with OUT.open("a", encoding="utf-8") as fout:
                for item in captured:
                    fout.write(json.dumps(item, ensure_ascii=False) + "\n")
        finally:
            await context.close()
            await browser.close()

    print(f"\nSaved POST bodies to: {OUT.resolve()}")


if __name__ == "__main__":
    asyncio.run(main())
