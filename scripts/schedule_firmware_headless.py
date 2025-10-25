"""Headless variant of the firmware scheduling automation script."""

from __future__ import annotations

import asyncio
import os

# Force the headless flag before importing the primary script to ensure it is honoured.
os.environ["FIRMWARE_HEADLESS"] = "true"

from schedule_firmware import run as _run  # type: ignore[attr-defined]


async def run() -> None:
    """Run the firmware scheduler with headless browser mode enforced."""

    await _run()


if __name__ == "__main__":
    asyncio.run(run())
