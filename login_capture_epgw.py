# login_capture_epgw.py
import asyncio

from playwright.async_api import (  # type: ignore[import]
    Error as PlaywrightError,
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
)

LOGIN_URL = "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx"  # adjust if your entry point differs


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, channel="msedge")
        context = await browser.new_context()  # ephemeral context just to capture state
        page = await context.new_page()

        navigation_failed = False
        try:
            await page.goto(LOGIN_URL, wait_until="networkidle")
        except PlaywrightError as exc:  # pragma: no cover - interactive workflow
            navigation_failed = True
            message = str(exc)
            if "ERR_INVALID_AUTH_CREDENTIALS" in message:
                print(
                    "\n>>> The gateway rejected the automatic request because it "
                    "requires manual credentials."
                )
                print(
                    "    Use the Edge window to complete the login (SSO/NTLM/MFA)."
                )
                print("    Once the site finishes loading, return here and continue.")
            else:
                raise

        print("\n>>> Log in in the Edge window (SSO/NTLM/MFA/etc).")
        input("Press ENTER here once the site shows you're logged in... ")

        if navigation_failed:
            try:
                await page.wait_for_load_state("networkidle", timeout=15_000)
            except PlaywrightTimeoutError:  # pragma: no cover - interactive workflow
                pass

        await context.storage_state(path="storage_state.json")
        print("Saved storage_state.json")

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
