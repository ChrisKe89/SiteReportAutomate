"""Record a minimal HAR while performing a manual firmware lookup."""

from __future__ import annotations

import asyncio
import os
from pathlib import Path

from dotenv import load_dotenv  # type: ignore[import]
from playwright.async_api import (  # type: ignore[import]
    Error as PlaywrightError,
    async_playwright,
)

from scripts.playwright_launch import launch_browser


def _env_path(var_name: str, default: str) -> Path:
    raw_value = os.getenv(var_name, default)
    normalised = raw_value.replace("\\", "/")
    return Path(normalised).expanduser()

URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx"
HAR_PATH = Path("logs/firmware_lookup.har.zip")
STORAGE_STATE_PATH = _env_path("FIRMWARE_STORAGE_STATE", "storage_state.json")
HTTP_USERNAME = os.getenv("FIRMWARE_HTTP_USERNAME")
HTTP_PASSWORD = os.getenv("FIRMWARE_HTTP_PASSWORD")
ALLOWLIST = os.getenv("FIRMWARE_AUTH_ALLOWLIST", "*.fujixerox.net,*.xerox.com")
AUTH_WARMUP_URL = os.getenv(
    "FIRMWARE_WARMUP_URL",
    "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx",
)

load_dotenv()


async def main() -> None:
    HAR_PATH.parent.mkdir(parents=True, exist_ok=True)

    channel = os.getenv("FIRMWARE_BROWSER_CHANNEL", "").strip() or None

    storage_state = STORAGE_STATE_PATH if STORAGE_STATE_PATH.exists() else None
    if storage_state is None:
        print(
            "\n> No saved credentials were found. Run"
            " 'python scripts/login_capture_remote_firmware.py'"
            " to capture them before recording the HAR."
        )

    context_kwargs = {
        "record_har_path": str(HAR_PATH),
        "record_har_mode": "minimal",
    }
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
            channel=channel,
            storage_state_path=storage_state,
            browser_args=browser_args,
            context_kwargs=context_kwargs,
        )
        try:
            page = await context.new_page()
            page.set_default_navigation_timeout(45_000)
            page.set_default_timeout(45_000)

            if AUTH_WARMUP_URL:
                try:
                    await page.goto(AUTH_WARMUP_URL, wait_until="domcontentloaded")
                except PlaywrightError as exc:
                    print(f"Warm-up navigation to {AUTH_WARMUP_URL} failed: {exc}")

            await page.goto(URL, wait_until="domcontentloaded")

            print("\n> Do ONE normal lookup in the page (fill & submit).")
            print("> When the result appears, return here and press Enter…")
            input()
        finally:
            await context.close()
            await browser.close()

    print(f"\nSaved HAR: {HAR_PATH.resolve()}")


if __name__ == "__main__":
    asyncio.run(main())
