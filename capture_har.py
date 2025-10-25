"""Record a minimal HAR while performing a manual firmware lookup."""

from __future__ import annotations

import asyncio
import os
from pathlib import Path

from dotenv import load_dotenv  # type: ignore[import]
from playwright.async_api import async_playwright  # type: ignore[import]

from scripts.playwright_launch import launch_browser

URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx"
HAR_PATH = Path("logs/firmware_lookup.har.zip")

load_dotenv()


async def main() -> None:
    HAR_PATH.parent.mkdir(parents=True, exist_ok=True)

    channel = os.getenv("FIRMWARE_BROWSER_CHANNEL", "").strip() or None

    async with async_playwright() as playwright:
        browser, context = await launch_browser(
            playwright,
            headless=False,
            channel=channel,
            storage_state_path=None,
            context_kwargs={
                "record_har_path": str(HAR_PATH),
                "record_har_mode": "minimal",
            },
        )
        try:
            page = await context.new_page()
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
