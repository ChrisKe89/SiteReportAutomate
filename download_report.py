# download_report.py
import asyncio
from datetime import datetime
from pathlib import Path
from playwright.async_api import (
    async_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

# --- Site config ---
BASE_URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"
REPORT_URL = f"{BASE_URL}/firmware/DeviceList.aspx"

# --- Paths ---
DOWNLOAD_DIR = Path("downloads")
DOWNLOAD_DIR.mkdir(exist_ok=True)
USER_DATA_DIR = Path("user-data")  # persists browser profile (cookies, IWA trust, etc.)
USER_DATA_DIR.mkdir(exist_ok=True)

# --- Selectors (from your page snippet) ---
DDL_OPCO = "#MainContent_ddlOpCoCode"
BTN_SEARCH = "#MainContent_btnSearch"
BTN_EXPORT = "#MainContent_btnExport"

# --- Options ---
HEADLESS = False  # flip to True after headed works reliably
ALLOWLIST = "*.fujixerox.net"  # for Windows Integrated Auth
NAV_TIMEOUT_MS = 45000
AFTER_SEARCH_WAIT_MS = 3000  # cushion for slow grids; tweak as needed


async def run_once():
    async with async_playwright() as p:
        # Prefer Edge (uncomment channel="msedge"); comment it out if you only have Chromium.
        context = await p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            headless=HEADLESS,
            channel="msedge",  # comment this line if Edge channel isn't available
            args=[
                f"--auth-server-allowlist={ALLOWLIST}",
                f"--auth-negotiate-delegate-allowlist={ALLOWLIST}",
                "--start-minimized",
            ],
            accept_downloads=True,
        )

        try:
            page = await context.new_page()
            page.set_default_navigation_timeout(NAV_TIMEOUT_MS)
            page.set_default_timeout(NAV_TIMEOUT_MS)

            # 1) Go to report page (IWA should auto-auth if your Windows session has access)
            await page.goto(REPORT_URL, wait_until="networkidle")

            # 2) Set dropdown to FBAU (value="FXAU")
            await page.select_option(DDL_OPCO, "FXAU")

            # 3) Click Search and wait for results to load/settle
            await page.click(BTN_SEARCH)
            # Either wait for network idle or a short cushion for the grid to redraw
            try:
                await page.wait_for_load_state("networkidle")
            except PlaywrightTimeoutError:
                # Some ASP.NET pages keep connections open; use a timed wait as fallback
                pass
            await page.wait_for_timeout(AFTER_SEARCH_WAIT_MS)

            # 4) Click Export and capture the download
            async with page.expect_download() as download_info:
                await page.click(BTN_EXPORT)
            download = await download_info.value

            # 5) Save with a timestamped filename
            suggested = download.suggested_filename
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            out_path = DOWNLOAD_DIR / f"{stamp}-{suggested}"
            await download.save_as(out_path)

            print(f"[OK] Saved: {out_path.resolve()}")

        finally:
            await context.close()


if __name__ == "__main__":
    asyncio.run(run_once())
