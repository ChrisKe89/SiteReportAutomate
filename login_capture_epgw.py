# login_capture_epgw.py
import asyncio

from playwright.async_api import async_playwright  # type: ignore[import]

LOGIN_URL = "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx"  # adjust if your entry point differs


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, channel="msedge")
        context = await browser.new_context()  # ephemeral context just to capture state
        page = await context.new_page()

        await page.goto(LOGIN_URL, wait_until="networkidle")
        print("\n>>> Log in in the Edge window (SSO/NTLM/MFA/etc).")
        input("Press ENTER here once the site shows you're logged in... ")

        await context.storage_state(path="storage_state.json")
        print("Saved storage_state.json")

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
