# download_report.py
import asyncio
from datetime import datetime
from pathlib import Path
from playwright.async_api import async_playwright

BASE_URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"
REPORT_URL = f"{BASE_URL}/firmware/DeviceList.aspx"  # adjust if you navigate via UI
DOWNLOAD_DIR = Path("downloads")
DOWNLOAD_DIR.mkdir(exist_ok=True)

USER_DATA_DIR = Path("user-data")  # persists cookies, windows auth, etc.
USER_DATA_DIR.mkdir(exist_ok=True)


async def main():
    async with async_playwright() as p:
        # Use Edge if available; otherwise drop channel="msedge"
        context = await p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            headless=False,  # keep visible until verified; flip to True later
            channel="msedge",  # comment this if you don't have Edge channel
            args=[
                "--auth-server-allowlist=*.fujixerox.net",
                "--auth-negotiate-delegate-allowlist=*.fujixerox.net",
                "--start-minimized",
            ],
            accept_downloads=True,
        )

        page = await context.new_page()

        # Navigate; with IWA, Chromium should auto-auth using your Windows session
        await page.goto(REPORT_URL, wait_until="networkidle")

        # TODO: apply any filters here, e.g.:
        # await page.fill("#dateFrom", "2025-09-01")
        # await page.fill("#dateTo", "2025-09-29")
        # await page.click("text=Apply Filters")
        # await page.wait_for_load_state("networkidle")

        # Trigger export and wait for the download handle
        async with page.expect_download() as dl_info:
            await page.click("text=Export")  # update selector as needed
        download = await dl_info.value

        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        out_path = DOWNLOAD_DIR / f"{stamp}-{download.suggested_filename}"
        await download.save_as(out_path)
        print(f"Saved: {out_path.resolve()}")

        await context.close()


if __name__ == "__main__":
    asyncio.run(main())
