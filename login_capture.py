import asyncio
from playwright.async_api import async_playwright  # type: ignore[import]

LOGIN_URL = "https://example.com/login"  # <-- put the real login URL


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False
        )  # show browser so you can do MFA
        context = await browser.new_context()
        page = await context.new_page()

        # Go to login
        await page.goto(LOGIN_URL, wait_until="networkidle")

        # TODO: Update selectors for your site’s login form
        # await page.fill("#username", "YOUR_USERNAME")
        # await page.fill("#password", "YOUR_PASSWORD")
        # await page.click("button[type=submit]")

        print("\n>>> Complete any MFA in the real browser window. <<<")
        print(
            "When you see the site's main/home/report page fully loaded, return here."
        )

        # Pause here so you can do MFA and land on a logged-in page
        # Press Enter in this console once you're fully signed in
        input("\nPress ENTER here once the site shows you're logged in... ")

        # Save cookies/localStorage to reuse later
        await context.storage_state(path="storage_state.json")
        print("Saved storage_state.json")

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
