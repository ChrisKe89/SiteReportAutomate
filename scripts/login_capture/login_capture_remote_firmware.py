"""Capture Playwright storage state for the remote firmware site."""

from __future__ import annotations

import argparse
import asyncio
from pathlib import Path
from typing import Iterable, Literal, Sequence, TypedDict, cast

from playwright.async_api import (  # type: ignore[import]
    Error as PlaywrightError,
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
)

from playwright_launch import launch_browser

DEFAULT_STORAGE_STATE = "storage_state.json"
DEFAULT_BROWSER_CHANNEL = ""

WaitUntilState = Literal["commit", "domcontentloaded", "load", "networkidle"]


class TargetConfigBase(TypedDict):
    label: str
    url: str


class TargetConfig(TargetConfigBase, total=False):
    wait_until: WaitUntilState


TARGETS: dict[str, TargetConfig] = {
    "firmware": {
        "label": "Firmware scheduler",
        "url": "https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx",
        "wait_until": "networkidle",
    }
}

DEFAULT_SEQUENCE: Sequence[str] = ("firmware",)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Launch Edge and manually authenticate so Playwright can persist cookies "
            "for subsequent automated runs."
        )
    )
    parser.add_argument(
        "--site",
        action="append",
        choices=sorted(TARGETS),
        help=(
            "Capture login for one or more sites. "
            "Repeat the flag to include multiple entries. "
            "Defaults to the remote firmware scheduler."
        ),
    )
    parser.add_argument(
        "--storage-state",
        default=DEFAULT_STORAGE_STATE,
        help=f"Where to save the captured storage state (default: {DEFAULT_STORAGE_STATE}).",
    )
    parser.add_argument(
        "--browser-channel",
        default=DEFAULT_BROWSER_CHANNEL,
        help=f"Chromium channel to launch (default: {DEFAULT_BROWSER_CHANNEL}).",
    )
    return parser.parse_args()


def dedupe_preserve_order(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for item in items:
        if item not in seen:
            ordered.append(item)
            seen.add(item)
    return ordered


async def capture_logins(
    site_keys: Sequence[str],
    storage_state_path: Path,
    browser_channel: str,
) -> None:
    targets = [TARGETS[key] for key in site_keys]

    async with async_playwright() as p:
        browser, context = await launch_browser(
            p,
            headless=False,
            channel=browser_channel.strip() or None,
            storage_state_path=None,
        )
        try:
            page = await context.new_page()

            total = len(targets)
            for index, target in enumerate(targets, start=1):
                label = target["label"]
                url = target["url"]
                wait_until = cast(WaitUntilState, target.get("wait_until", "networkidle"))

                print(f"\n[{index}/{total}] Opening {label}: {url}")

                try:
                    await page.goto(url, wait_until=wait_until)
                except PlaywrightError as exc:  # pragma: no cover - interactive workflow
                    message = str(exc)
                    if "ERR_INVALID_AUTH_CREDENTIALS" in message:
                        print(
                            "\n>>> The gateway rejected the automatic request because it "
                            "requires manual credentials."
                        )
                        print(
                            "    Use the Edge window to complete the login (SSO/NTLM/MFA)."
                        )
                        print(
                            "    Once the site finishes loading, return here and continue."
                        )
                    else:
                        raise

                print(">>> Complete any interactive login in the Edge window.")
                input(f"Press ENTER here once '{label}' shows you are signed in... ")

                try:
                    await page.wait_for_load_state("networkidle", timeout=15_000)
                except PlaywrightTimeoutError:  # pragma: no cover - page may stay busy
                    pass

            await context.storage_state(path=str(storage_state_path))
            print(f"\nSaved {storage_state_path}")
        finally:
            await context.close()
            await browser.close()


async def main() -> None:
    args = parse_args()

    site_keys = dedupe_preserve_order(args.site or DEFAULT_SEQUENCE)
    storage_state_path = Path(args.storage_state).expanduser().resolve()
    storage_state_path.parent.mkdir(parents=True, exist_ok=True)

    await capture_logins(site_keys, storage_state_path, args.browser_channel)


if __name__ == "__main__":
    asyncio.run(main())
