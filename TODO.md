"""Capture POST payloads made during a manual firmware lookup (WebForms-safe)."""
from __future__ import annotations

import asyncio, json, os
from pathlib import Path
from typing import Any
from dotenv import load_dotenv  # type: ignore[import]
from playwright.async_api import Error as PlaywrightError, Request, Route, async_playwright  # type: ignore[import]
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

        captured: list[dict[str, Any]] = []

        async def handle_route(route: Route, request: Request) -> None:
            try:
                if request.method.upper() == "POST" and "SingleRequest.aspx" in request.url:
                    # Try all ways to get the body
                    body = request.post_data or ""
                    # Some Playwright builds expose a bytes buffer; stringify if present
                    try:
                        if not body and hasattr(request, "post_data_json") and request.post_data_json is not None:
                            body = json.dumps(request.post_data_json, ensure_ascii=False)
                    except Exception:
                        pass

                    captured.append({
                        "url": request.url,
                        "method": request.method,
                        "resource_type": request.resource_type,  # often "document"
                        "headers": request.headers,
                        "post_data": body,
                    })
            finally:
                await route.continue_()

        await context.route("**/firmware/SingleRequest.aspx", handle_route)

        page = await context.new_page()
        page.set_default_navigation_timeout(45_000)
        page.set_default_timeout(45_000)

        # Optional gateway warmup (Basic/NTLM/Kerberos)
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

        await context.close()
        await browser.close()

    print(f"\nSaved POST bodies to: {OUT.resolve()}")

if __name__ == "__main__":
    asyncio.run(main())
