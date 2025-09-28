import asyncio
from datetime import datetime
from pathlib import Path
from playwright.async_api import async_playwright

BASE_URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"  # <-- site base
REPORT_URL = (
    "https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/DeviceList.aspx"  # <-- direct URL to the report page if possible
)
DOWNLOAD_DIR = Path("downloads")  # where files land
DOWNLOAD_DIR.mkdir(exist_ok=True)


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)  # keep visible until verified
        context = await browser.new_context(
            storage_state="storage_state.json", accept_downloads=True
        )
        page = await context.new_page()

        # Go to the report page (or navigate via UI if direct URL doesn’t allow it)
        await page.goto(REPORT_URL, wait_until="networkidle")

        # TODO: apply filters/date range if needed:
        # await page.fill("#dateFrom", "2025-09-01")
        # await page.fill("#dateTo", "2025-09-29")
        # await page.click("text=Apply Filters")
        # await page.wait_for_timeout(1000)

        # Trigger export and wait for download
        # Replace the selector below with the actual Export button
        async with page.expect_download() as dl_info:
            await page.click("text=Export")  # e.g., button with text "Export"
        download = await dl_info.value

        # Give the file a timestamped name
        suggested = download.suggested_filename
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        out_path = DOWNLOAD_DIR / f"{stamp}-{suggested}"
        await download.save_as(out_path)
        print(f"Saved: {out_path.resolve()}")

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
