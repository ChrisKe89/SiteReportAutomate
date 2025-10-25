"""Shared Playwright launch helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable


async def launch_browser(
    playwright: Any,
    *,
    headless: bool,
    channel: str | None,
    storage_state_path: Path | None,
    browser_args: Iterable[str] | None = None,
    context_kwargs: dict[str, Any] | None = None,
):
    """Launch Chromium using the shared headless/channel/storage settings."""
    launch_kwargs: dict[str, Any] = {"headless": headless}
    if browser_args:
        launch_kwargs["args"] = list(browser_args)
    if channel and channel.strip():
        browser_type = getattr(playwright, "chromium")
        launch_kwargs["channel"] = channel
    else:
        browser_type = playwright.chromium

    browser = await browser_type.launch(**launch_kwargs)

    resolved_context_kwargs: dict[str, Any] = {}
    if context_kwargs:
        resolved_context_kwargs.update(context_kwargs)
    if storage_state_path and storage_state_path.exists():
        resolved_context_kwargs.setdefault("storage_state", str(storage_state_path))

    context = await browser.new_context(**resolved_context_kwargs)
    return browser, context
