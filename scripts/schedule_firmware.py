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
PLAYWRIGHT_STEPS_LOG = Path(
    os.getenv("FIRMWARE_STEPS_LOG", "downloads/playwright_steps.jsonl")
)
BROWSER_CHANNEL = os.getenv("FIRMWARE_BROWSER_CHANNEL", "msedge")
ERRORS_JSON = Path(os.getenv("FIRMWARE_ERRORS_JSON", "errors.json")).expanduser()
HTTP_USERNAME = os.getenv("FIRMWARE_HTTP_USERNAME")
HTTP_PASSWORD = os.getenv("FIRMWARE_HTTP_PASSWORD")
AUTH_WARMUP_URL = os.getenv(
    "FIRMWARE_WARMUP_URL",
    "http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx",
)
ALLOWLIST = os.getenv("FIRMWARE_AUTH_ALLOWLIST", "*.fujixerox.net,*.xerox.com")
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


def _normalise_time_label(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip()).upper()


ALLOWED_TIME_OPTIONS: list[tuple[str, str]] = [
    ("00", "12 AM"),
    ("01", "01 AM"),
    ("02", "02 AM"),
    ("03", "03 AM"),
    ("04", "04 AM"),
    ("05", "05 AM"),
    ("06", "06 AM"),
    ("07", "07 AM"),
    ("19", "07 PM"),
    ("20", "08 PM"),
    ("21", "09 PM"),
    ("22", "10 PM"),
    ("23", "11 PM"),
]

EXPECTED_TIME_LABELS = [label for _, label in ALLOWED_TIME_OPTIONS]
EXPECTED_TIME_LABELS_NORMALISED = {
    _normalise_time_label(label) for label in EXPECTED_TIME_LABELS
}

ALLOWED_TIME_LABEL_BY_VALUE = {value: label for value, label in ALLOWED_TIME_OPTIONS}

PREFERRED_TIME_VALUE, PREFERRED_TIME_LABEL = ALLOWED_TIME_OPTIONS[0]


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
        self.log_path = PLAYWRIGHT_STEPS_LOG.expanduser()
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        self.log(
            "step-recorder-initialised",
            extra={
                "product_code": row.product_code,
                "state": row.state,
                "opco": row.opco,
            },
        )

    def log(self, action: str, *, extra: dict[str, Any] | None = None) -> None:
        payload: dict[str, Any] = {
            "timestamp": datetime.now().isoformat(timespec="milliseconds"),
            "serial_number": self.row.serial_number,
            "action": action,
            "step_counter": self.counter,
        }
        if extra:
            payload.update(extra)
        with self.log_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(payload) + "\n")

    async def capture(self, label: str) -> None:
        self.counter += 1
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        filename = f"{self.counter:02d}_{timestamp}_{slugify_label(label)}.png"
        path = self.base_dir / filename
        self.log("capture", extra={"label": label, "screenshot_path": str(path)})
        try:
            await self.page.screenshot(path=str(path), full_page=True)
        except PlaywrightError:
            # Ignore screenshot failures to avoid blocking the workflow
            pass


async def read_status_messages(page: Page) -> str:
    """Return the latest status message shown on the page."""

    selectors = [
        "#MainContent_MessageLabel li",
        "#MainContent_MessageLabel",
        "#MainContent_lblMessage",
        "#MainContent_lblStatus",
    ]
    messages: list[str] = []

    for selector in selectors:
        locator = page.locator(selector)
        count = await locator.count()
        if count == 0:
            continue

        for idx in range(count):
            element = locator.nth(idx)
            try:
                await element.wait_for(state="visible", timeout=5_000)
            except PlaywrightTimeoutError:
                pass

            try:
                text = (await element.inner_text()).strip()
            except PlaywrightTimeoutError:
                continue
            if not text:
                continue

            normalised = re.sub(r"\s+", " ", text)
            if normalised:
                messages.append(normalised)

    unique_messages: list[str] = []
    for message in messages:
        if message not in unique_messages:
            unique_messages.append(message)

    return " | ".join(unique_messages)


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
        raise RuntimeError(
            f"Unable to populate input {selector!r} with value {value!r}"
        )


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
    """Return a random date between three and six days from today (inclusive)."""
    start = datetime.now().date() + timedelta(days=3)
    end = datetime.now().date() + timedelta(days=6)
    span = (end - start).days
    offset = random.randint(0, span)
    return datetime.combine(start + timedelta(days=offset), datetime.min.time())


def pick_time_option(options: list[tuple[str, str]]) -> tuple[str, str] | None:
    """Return the first allowed option (preferring midnight) that exists."""

    cleaned_options: list[tuple[str, str]] = []
    for value, label in options:
        cleaned_value = value.strip()
        cleaned_label = label.strip()
        if not cleaned_value and not cleaned_label:
            continue
        cleaned_options.append((cleaned_value, cleaned_label))

    for allowed_value, allowed_label in ALLOWED_TIME_OPTIONS:
        allowed_normalised = _normalise_time_label(allowed_label)
        for value, label in cleaned_options:
            candidate_label = label or allowed_label
            candidate_normalised = _normalise_time_label(candidate_label)
            if candidate_normalised == allowed_normalised or (
                value and value == allowed_value
            ):
                return value or allowed_value, label or allowed_label

    return None


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


async def select_time(
    page: Page, *, stepper: StepRecorder | None = None
) -> tuple[str, str]:
    selectors = [
        "select#MainContent_ddlScheduleTime",
    ]
    for selector in selectors:
        dropdown_group = page.locator(selector)
        count = await dropdown_group.count()
        if count == 0:
            continue

        for idx in range(count):
            dropdown = dropdown_group.nth(idx)
            if not await dropdown.is_visible():
                continue
            if await dropdown.is_disabled():
                continue

            options_locator = dropdown.locator("option")
            total = await options_locator.count()
            options: list[tuple[str, str]] = []
            for option_idx in range(total):
                option = options_locator.nth(option_idx)
                value = ((await option.get_attribute("value")) or "").strip()
                label = (await option.inner_text()).strip()
                if not value and "select" in label.lower():
                    continue
                options.append((value, label))

            if stepper:
                observed_normalised = {
                    _normalise_time_label(label)
                    for _, label in options
                    if label.strip()
                }
                missing_expected = [
                    label
                    for label in EXPECTED_TIME_LABELS
                    if _normalise_time_label(label) not in observed_normalised
                ]
                unexpected_labels = [
                    label
                    for _, label in options
                    if label.strip()
                    and _normalise_time_label(label)
                    not in EXPECTED_TIME_LABELS_NORMALISED
                ]
                stepper.log(
                    "time-options-detected",
                    extra={
                        "selector": selector,
                        "option_count": len(options),
                        "options": options,
                        "missing_expected_labels": missing_expected,
                        "unexpected_labels": unexpected_labels,
                    },
                )
            choice = pick_time_option(options)
            if not choice:
                if stepper:
                    stepper.log(
                        "time-allowed-option-missing",
                        extra={
                            "selector": selector,
                            "option_count": len(options),
                            "allowed_time_options": ALLOWED_TIME_OPTIONS,
                        },
                    )
                continue

            value, label = choice
            cleaned_value = value.strip()
            cleaned_label = label.strip()
            canonical_label = cleaned_label or ALLOWED_TIME_LABEL_BY_VALUE.get(
                cleaned_value, ""
            )
            if not canonical_label:
                canonical_label = ALLOWED_TIME_LABEL_BY_VALUE.get(
                    PREFERRED_TIME_VALUE, PREFERRED_TIME_LABEL
                )
            target_label = canonical_label
            if stepper:
                stepper.log(
                    "time-option-selected",
                    extra={
                        "value": cleaned_value,
                        "label": cleaned_label or canonical_label,
                        "canonical_label": canonical_label,
                        "selector": selector,
                    },
                )

            try:
                await dropdown.select_option(label=target_label)
            except PlaywrightError:
                if cleaned_value:
                    await dropdown.select_option(value=cleaned_value)
                else:
                    raise

            # Confirm the dropdown reflects the selected option; fall back to JS if needed.
            selected_value = (await dropdown.input_value()).strip()
            if cleaned_value and selected_value != cleaned_value:
                await dropdown.evaluate(
                    "(el, value) => { el.value = value; el.dispatchEvent(new Event('change', { bubbles: true })); }",
                    cleaned_value,
                )
                selected_value = (await dropdown.input_value()).strip()

            selected_label_locator = dropdown.locator("option:checked")
            if await selected_label_locator.count() > 0:
                selected_label = (
                    await selected_label_locator.first.inner_text()
                ).strip()
            else:
                selected_label = target_label

            if selected_label.lower().replace(" ", "") != target_label.lower().replace(
                " ", ""
            ):
                await dropdown.evaluate(
                    "(el, payload) => {"
                    "  const options = Array.from(el.options);"
                    "  const match = options.find(o => o.text.trim().toLowerCase() === payload.label.trim().toLowerCase());"
                    "  if (match) {"
                    "    match.selected = true;"
                    "    el.dispatchEvent(new Event('change', { bubbles: true }));"
                    "  }"
                    "}",
                    {"label": target_label},
                )
                selected_value = (await dropdown.input_value()).strip()
                if await selected_label_locator.count() > 0:
                    selected_label = (
                        await selected_label_locator.first.inner_text()
                    ).strip()

            if not selected_value and cleaned_value:
                # As a last resort, set both value and label to ensure a submission-friendly state.
                await dropdown.evaluate(
                    "(el, payload) => {"
                    "  el.value = payload.value;"
                    "  const option = Array.from(el.options).find(o => o.value === payload.value || o.text === payload.label);"
                    "  if (option) option.selected = true;"
                    "  el.dispatchEvent(new Event('change', { bubbles: true }));"
                    "}",
                    {"value": cleaned_value, "label": canonical_label},
                )
                selected_value = (await dropdown.input_value()).strip()

            if stepper:
                stepper.log(
                    "time-option-confirmed",
                    extra={
                        "selected_value": selected_value or cleaned_value,
                        "selected_label": selected_label,
                        "selector": selector,
                    },
                )

            if selected_value or not cleaned_value:
                return selected_value or cleaned_value, selected_label
    raise RuntimeError("Could not determine a valid time option to select.")


async def set_schedule_date(page: Page, target: datetime) -> str:
    date_str = target.strftime("%d/%m/%Y")
    locator = page.locator("#MainContent_txtDateTime")
    await locator.wait_for(state="visible")
    await locator.click()

    # Attempt to use a jQuery UI style date picker if present.
    datepicker = page.locator("#ui-datepicker-div")
    if await datepicker.count() > 0:
        try:
            await datepicker.wait_for(state="visible")

            target_date = target.date()

            async def current_month_year() -> tuple[int, int]:
                month_text = (
                    await datepicker.locator(".ui-datepicker-month").first.inner_text()
                ).strip()
                year_text = (
                    await datepicker.locator(".ui-datepicker-year").first.inner_text()
                ).strip()
                month_number = datetime.strptime(month_text, "%B").month
                return month_number, int(year_text)

            # Navigate to the month containing the target date.
            for _ in range(12):
                month_number, year_number = await current_month_year()
                if (
                    year_number == target_date.year
                    and month_number == target_date.month
                ):
                    break
                await datepicker.locator(".ui-datepicker-next").click()
                await page.wait_for_timeout(200)

            day_locator = datepicker.locator(".ui-datepicker-calendar td a").filter(
                has_text=str(target_date.day)
            )
            if await day_locator.count() > 0:
                await day_locator.first.click()
                return date_str
        except Exception:  # noqa: BLE001
            pass

    # Fallback: directly set the value if date picker interaction fails.
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
    stepper.log("row-processing-started")
    try:
        stepper.log("selecting-opco", extra={"opco": OPCO_VALUE})
        await ensure_option_selected(page, "#MainContent_ddlOpCoID", OPCO_VALUE)
        await stepper.capture("opco-selected")
        stepper.log("filling-product-code", extra={"product_code": row.product_code})
        await fill_input(page, "#MainContent_ProductCode", row.product_code)
        await stepper.capture("product-code-filled")
        stepper.log("filling-serial-number", extra={"serial_number": row.serial_number})
        await fill_input(page, "#MainContent_SerialNumber", row.serial_number)
        await stepper.capture("serial-number-filled")
        stepper.log("clicking-search")
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
                stepper.log(
                    "row-processing-complete", extra={"status": "AlreadyUpgraded"}
                )
                return

        device_table = page.locator("#MainContent_GridViewDevice")
        if await device_table.count() == 0 or not await device_table.is_visible():
            message = await read_status_messages(page)
            if not message:
                message = "Device table not visible after search"
            await stepper.capture("device-table-missing")
            append_log(row, "NotEligible", message)
            record_error(row, "NotEligible", message)
            stepper.log(
                "row-processing-complete",
                extra={"status": "NotEligible", "message": message},
            )
            return

        schedule_date = pick_random_schedule_date()
        stepper.log(
            "setting-schedule-date",
            extra={"target_date": schedule_date.isoformat(timespec="seconds")},
        )
        scheduled_date_str = await set_schedule_date(page, schedule_date)
        await stepper.capture("schedule-date-set")
        time_value, time_label = await select_time(page, stepper=stepper)
        await stepper.capture("time-selected")
        stepper.log(
            "selecting-timezone",
            extra={
                "state": row.state,
                "timezone_mapping": TIMEZONE_BY_STATE.get(row.state, ""),
            },
        )
        timezone_value = await select_timezone(page, row.state)
        await stepper.capture("timezone-selected")

        stepper.log("clicking-submit")
        await page.click("#MainContent_submitButton")
        await stepper.capture("submit-clicked")
        try:
            await page.wait_for_timeout(500)
        except PlaywrightTimeoutError:
            pass

        status_text = await read_status_messages(page)
        if not status_text:
            status_text = "Scheduled"

        append_log(
            row,
            "Scheduled",
            status_text,
            scheduled_date=scheduled_date_str,
            scheduled_time=time_label or time_value,
            timezone=timezone_value,
        )
        stepper.log(
            "row-processing-complete",
            extra={
                "status": "Scheduled",
                "status_text": status_text,
                "scheduled_date": scheduled_date_str,
                "scheduled_time": time_label or time_value,
                "timezone": timezone_value,
            },
        )
        await stepper.capture("row-complete")
    except Exception:
        await stepper.capture("error")
        stepper.log("row-processing-error")
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
                    status_text = await read_status_messages(page)
                    message = status_text or str(exc)
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
