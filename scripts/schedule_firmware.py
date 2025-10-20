"""Automate firmware scheduling via Playwright.

This script reads device rows from an Excel workbook, searches the
SingleRequest firmware page, and schedules upgrades when eligible.
"""

from __future__ import annotations

import asyncio
import json
import os
import random
import re
import string
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from dotenv import load_dotenv  # type: ignore[import]
from openpyxl import Workbook, load_workbook  # type: ignore[import]
from playwright.async_api import (  # type: ignore[import]
    BrowserContext,
    Error as PlaywrightError,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
)

load_dotenv()

DEFAULT_URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net/firmware/SingleRequest.aspx"
INPUT_XLSX = Path(os.getenv("FIRMWARE_INPUT_XLSX", "downloads/VIC.xlsx"))
LOG_XLSX = Path(os.getenv("FIRMWARE_LOG_XLSX", "downloads/FirmwareLog.xlsx"))
BROWSER_CHANNEL = os.getenv("FIRMWARE_BROWSER_CHANNEL", "msedge")
ERRORS_JSON = Path(os.getenv("FIRMWARE_ERRORS_JSON", "errors.json")).expanduser()
HTTP_USERNAME = os.getenv("FIRMWARE_HTTP_USERNAME")
HTTP_PASSWORD = os.getenv("FIRMWARE_HTTP_PASSWORD")
AUTH_WARMUP_URL = os.getenv(
    "FIRMWARE_WARMUP_URL",
    "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx",
)
ALLOWLIST = os.getenv(
    "FIRMWARE_AUTH_ALLOWLIST", "*.fujixerox.net,*.xerox.com"
)
HEADLESS = os.getenv("FIRMWARE_HEADLESS", "false").lower() in {"1", "true", "yes"}
USER_DATA_DIR = Path(os.getenv("FIRMWARE_USER_DATA_DIR", "user-data"))
SCREENSHOT_DIR = Path(os.getenv("FIRMWARE_SCREENSHOT_DIR", "downloads/screenshots"))

OPCO_VALUE = "FXAU"  # FBAU label

# Map uppercase state abbreviations to timezone dropdown values
TIMEZONE_BY_STATE = {
    "NT": "+09:30",
    "SA": "+10:30",
    "ACT": "+11:00",
    "VIC": "+11:00",
    "NSW": "+11:00",
    "QLD": "+10:00",
    "TAS": "+11:00",
}


@dataclass
class DeviceRow:
    serial_number: str
    product_code: str
    state: str
    opco: str

    @classmethod
    def from_iterable(cls, row: Iterable[Any]) -> "DeviceRow | None":
        cells = ["" if cell is None else str(cell).strip() for cell in row]
        if len(cells) < 4:
            return None
        serial_number, product_code, opco, state = cells[:4]
        if not serial_number or not product_code:
            return None
        return cls(
            serial_number=serial_number,
            product_code=product_code,
            opco=opco or "",
            state=state.upper() if state else "",
        )


async def ensure_option_selected(
    page: Page, selector: str, value: str, *, timeout: float = 10_000
) -> None:
    await page.wait_for_selector(selector, state="visible", timeout=timeout)
    await page.select_option(selector, value=value)


def slugify_label(value: str) -> str:
    value = value.strip().lower()
    allowed = set(string.ascii_lowercase + string.digits + "-_")
    slug = ["-" if ch.isspace() else ch for ch in value]
    filtered = [ch for ch in slug if ch in allowed]
    result = "".join(filtered) or "step"
    while "--" in result:
        result = result.replace("--", "-")
    return result.strip("-") or "step"


class StepRecorder:
    def __init__(self, page: Page, row: DeviceRow):
        self.page = page
        self.row = row
        self.counter = 0
        self.base_dir = SCREENSHOT_DIR / slugify_label(row.serial_number or "row")
        self.base_dir.mkdir(parents=True, exist_ok=True)

    async def capture(self, label: str) -> None:
        self.counter += 1
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        filename = f"{self.counter:02d}_{timestamp}_{slugify_label(label)}.png"
        path = self.base_dir / filename
        try:
            await self.page.screenshot(path=str(path), full_page=True)
        except PlaywrightError:
            # Ignore screenshot failures to avoid blocking the workflow
            pass


async def fill_input(page: Page, selector: str, value: str) -> None:
    locator = page.locator(selector)
    await locator.wait_for(state="visible")
    await locator.click()
    await locator.fill(value)
    actual = (await locator.input_value()).strip()
    if actual != value:
        await locator.press("Control+A")
        await locator.type(value)
        actual = (await locator.input_value()).strip()
    if actual != value:
        await locator.evaluate(
            "(el, value) => {"
            "  el.value = value;"
            "  el.dispatchEvent(new Event('input', { bubbles: true }));"
            "  el.dispatchEvent(new Event('change', { bubbles: true }));"
            "}",
            value,
        )
        actual = (await locator.input_value()).strip()
    if actual != value:
        raise RuntimeError(f"Unable to populate input {selector!r} with value {value!r}")


def load_rows(path: Path) -> list[DeviceRow]:
    if not path.exists():
        raise FileNotFoundError(f"Input workbook not found: {path}")
    workbook = load_workbook(path, read_only=True)
    sheet = workbook.active
    if sheet is None:
        workbook.close()
        raise ValueError(f"No active worksheet found in workbook: {path}")
    rows: list[DeviceRow] = []
    # Skip header row by using min_row=2
    for raw in sheet.iter_rows(min_row=2, values_only=True):
        item = DeviceRow.from_iterable(raw)
        if item:
            rows.append(item)
    workbook.close()
    return rows


def pick_random_schedule_date() -> datetime:
    """Return a random date between tomorrow and six days from today (inclusive)."""
    start = datetime.now().date() + timedelta(days=1)
    end = datetime.now().date() + timedelta(days=6)
    span = (end - start).days
    offset = random.randint(0, span)
    return datetime.combine(start + timedelta(days=offset), datetime.min.time())


def pick_time_option(options: list[tuple[str, str]]) -> tuple[str, str] | None:
    """Pick a random time option in allowed windows from (value, label) pairs."""
    allowed: list[tuple[str, str]] = []
    for value, label in options:
        if not value or value.lower() == "select":
            continue
        match = re.search(r"(\d{1,2})(?::(\d{2}))?\s*([ap]m)?", label, re.IGNORECASE)
        if not match:
            match = re.search(r"(\d{1,2}):(\d{2})", value)
        if not match:
            continue
        hour = int(match.group(1))
        minute = (
            int(match.group(2))
            if match.lastindex and match.lastindex >= 2 and match.group(2)
            else 0
        )
        meridiem = (
            match.group(3).lower()
            if match.lastindex and match.lastindex >= 3 and match.group(3)
            else ""
        )
        if meridiem:
            if meridiem == "pm" and hour != 12:
                hour += 12
            if meridiem == "am" and hour == 12:
                hour = 0
        hour = hour % 24
        total_minutes = hour * 60 + minute
        if total_minutes < 8 * 60 or total_minutes >= 18 * 60:
            allowed.append((value, label))
    if not allowed:
        return None
    return random.choice(allowed)


def append_log(
    row: DeviceRow,
    status: str,
    message: str,
    *,
    scheduled_date: str = "",
    scheduled_time: str = "",
    timezone: str = "",
) -> None:
    LOG_XLSX.parent.mkdir(parents=True, exist_ok=True)
    if LOG_XLSX.exists():
        workbook = load_workbook(LOG_XLSX)
        sheet = workbook.active
        if sheet is None:
            workbook.close()
            raise ValueError("Log workbook has no active worksheet.")
    else:
        workbook = Workbook()
        sheet = workbook.active
        if sheet is None:
            workbook.close()
            raise ValueError("Unable to create log worksheet.")
        sheet.append(
            [
                "SerialNumber",
                "Product_Code",
                "OpcoID",
                "State",
                "Status",
                "Message",
                "ScheduledDate",
                "ScheduledTime",
                "TimeZone",
                "LoggedAt",
            ]
        )
    sheet.append(
        [
            row.serial_number,
            row.product_code,
            row.opco,
            row.state,
            status,
            message,
            scheduled_date,
            scheduled_time,
            timezone,
            datetime.now().isoformat(timespec="seconds"),
        ]
    )
    workbook.save(LOG_XLSX)
    workbook.close()


def record_error(row: DeviceRow, status: str, message: str) -> None:
    """Append an error entry to the JSON ledger for quick troubleshooting."""

    if ERRORS_JSON.parent != Path("."):
        ERRORS_JSON.parent.mkdir(parents=True, exist_ok=True)

    entries: list[dict[str, str]] = []
    if ERRORS_JSON.exists():
        try:
            existing = json.loads(ERRORS_JSON.read_text(encoding="utf-8"))
            if isinstance(existing, list):
                entries = [entry for entry in existing if isinstance(entry, dict)]
        except json.JSONDecodeError:
            entries = []

    entries.append(
        {
            "serial_number": row.serial_number,
            "product_code": row.product_code,
            "opco": row.opco,
            "state": row.state,
            "status": status,
            "message": message,
            "logged_at": datetime.now().isoformat(timespec="seconds"),
        }
    )

    ERRORS_JSON.write_text(json.dumps(entries, indent=2), encoding="utf-8")


async def select_time(page: Page) -> tuple[str, str]:
    selectors = [
        "select#MainContent_ddlTime",
        "select#MainContent_ddlTimeSlot",
        "select[id*='ddlTimeSlot']",
        "select[id*='ddlTime'][id*='Slot']",
        "select[id*='ddlTime']:not(#MainContent_ddlTimeZone)",
    ]
    for selector in selectors:
        dropdown = page.locator(selector)
        if await dropdown.count() == 0:
            continue
        options_locator = dropdown.locator("option")
        total = await options_locator.count()
        options: list[tuple[str, str]] = []
        for idx in range(total):
            option = options_locator.nth(idx)
            value = (await option.get_attribute("value")) or ""
            label = (await option.inner_text()).strip()
            options.append((value, label))
        choice = pick_time_option(options)
        if choice:
            await dropdown.select_option(choice[0])
            return choice
    raise RuntimeError("Could not determine a valid time option to select.")


async def set_schedule_date(page: Page, target: datetime) -> str:
    date_str = target.strftime("%d/%m/%Y")
    locator = page.locator("#MainContent_txtDateTime")
    await locator.wait_for(state="visible")
    await locator.evaluate(
        "(el, value) => { el.removeAttribute('readonly'); el.value = value; el.dispatchEvent(new Event('change', { bubbles: true })); }",
        date_str,
    )
    return date_str


async def select_timezone(page: Page, state: str) -> str:
    value = TIMEZONE_BY_STATE.get(state.upper())
    if not value:
        raise ValueError(f"No timezone mapping for state: {state}")
    selector = "#MainContent_ddlTimeZone"
    await page.wait_for_selector(selector, state="visible")
    await page.select_option(selector, value=value)
    return value


async def handle_row(page: Page, row: DeviceRow) -> None:
    stepper = StepRecorder(page, row)
    await stepper.capture("row-start")
    try:
        await ensure_option_selected(page, "#MainContent_ddlOpCoID", OPCO_VALUE)
        await stepper.capture("opco-selected")
        await fill_input(page, "#MainContent_ProductCode", row.product_code)
        await stepper.capture("product-code-filled")
        await fill_input(page, "#MainContent_SerialNumber", row.serial_number)
        await stepper.capture("serial-number-filled")
        await page.click("#MainContent_btnSearch")
        await stepper.capture("search-clicked")
        try:
            await page.wait_for_timeout(1000)
            await page.wait_for_load_state("networkidle")
        except PlaywrightTimeoutError:
            pass

        eligibility_table = page.locator("#MainContent_GridViewEligibility")
        if await eligibility_table.count() > 0 and await eligibility_table.is_visible():
            text = (await eligibility_table.inner_text()).lower()
            if "already upgraded" in text:
                await stepper.capture("already-upgraded")
                append_log(row, "AlreadyUpgraded", "Device already upgraded; skipped")
                return

        device_table = page.locator("#MainContent_GridViewDevice")
        if await device_table.count() == 0 or not await device_table.is_visible():
            message = "Device table not visible after search"
            await stepper.capture("device-table-missing")
            append_log(row, "NotEligible", message)
            record_error(row, "NotEligible", message)
            return

        schedule_date = pick_random_schedule_date()
        scheduled_date_str = await set_schedule_date(page, schedule_date)
        await stepper.capture("schedule-date-set")
        _, time_label = await select_time(page)
        await stepper.capture("time-selected")
        timezone_value = await select_timezone(page, row.state)
        await stepper.capture("timezone-selected")

        await page.click("#MainContent_submitButton")
        await stepper.capture("submit-clicked")
        message_locator = page.locator("#MainContent_lblMessage, #MainContent_lblStatus")
        status_text = ""
        try:
            await message_locator.wait_for(state="visible", timeout=10_000)
            status_text = (await message_locator.inner_text()).strip()
        except PlaywrightTimeoutError:
            status_text = "No confirmation message"

        append_log(
            row,
            "Scheduled",
            status_text or "Scheduled",
            scheduled_date=scheduled_date_str,
            scheduled_time=time_label,
            timezone=timezone_value,
        )
        await stepper.capture("row-complete")
    except Exception:
        await stepper.capture("error")
        raise


async def run() -> None:
    rows = load_rows(INPUT_XLSX)
    if not rows:
        print("No rows to process")
        return

    user_data_dir = USER_DATA_DIR.expanduser()
    user_data_dir.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as p:
        launch_kwargs: dict[str, Any] = {
            "user_data_dir": str(user_data_dir),
            "headless": HEADLESS,
            "channel": BROWSER_CHANNEL,
            "args": [
                f"--auth-server-allowlist={ALLOWLIST}",
                f"--auth-negotiate-delegate-allowlist={ALLOWLIST}",
                "--start-minimized",
            ],
        }
        if HTTP_USERNAME and HTTP_PASSWORD:
            launch_kwargs["http_credentials"] = {
                "username": HTTP_USERNAME,
                "password": HTTP_PASSWORD,
            }

        context: BrowserContext = await p.chromium.launch_persistent_context(
            **launch_kwargs
        )

        try:
            page: Page = await context.new_page()
            page.set_default_navigation_timeout(45_000)
            page.set_default_timeout(45_000)

            if AUTH_WARMUP_URL:
                try:
                    await page.goto(AUTH_WARMUP_URL, wait_until="domcontentloaded")
                except PlaywrightError as exc:
                    print(f"Warm-up navigation to {AUTH_WARMUP_URL} failed: {exc}")

            await page.goto(DEFAULT_URL, wait_until="domcontentloaded")
            for row in rows:
                try:
                    await handle_row(page, row)
                except Exception as exc:  # noqa: BLE001
                    message = str(exc)
                    append_log(row, "Failed", message)
                    record_error(row, "Failed", message)
        except PlaywrightError as exc:
            message = str(exc)
            if "ERR_INVALID_AUTH_CREDENTIALS" in message:
                print(
                    "Authentication failed when opening the firmware scheduling page. "
                    "Ensure Integrated Windows Authentication is available or set "
                    "FIRMWARE_HTTP_USERNAME/FIRMWARE_HTTP_PASSWORD in your .env file."
                )
            else:
                raise
        finally:
            try:
                await context.close()
            except PlaywrightError:
                pass


if __name__ == "__main__":
    asyncio.run(run())
