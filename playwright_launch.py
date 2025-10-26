"""Helper for launching Playwright Chromium instances with shared defaults."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union

from playwright.async_api import Browser, BrowserContext, Playwright

StorageState = Union[str, Path, None]


async def launch_browser(
    playwright: Playwright,
    *,
    headless: bool = True,
    channel: str | None = None,
    storage_state_path: StorageState = None,
    context_overrides: Optional[Dict[str, Any]] = None,
) -> Tuple[Browser, BrowserContext]:
    """Launch Chromium and create a context, optionally reusing storage state."""

    launch_kwargs: Dict[str, Any] = {"headless": headless}
    if channel:
        launch_kwargs["channel"] = channel

    browser = await playwright.chromium.launch(**launch_kwargs)

    context_kwargs: Dict[str, Any] = dict(context_overrides or {})
    if storage_state_path:
        state_path = Path(storage_state_path)
        if state_path.exists():
            context_kwargs["storage_state"] = str(state_path)

    context = await browser.new_context(**context_kwargs)
    return browser, context
