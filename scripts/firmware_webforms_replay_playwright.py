#!/usr/bin/env python3
"""
Headless Playwright replayer for SingleRequest.aspx (SEARCH + SCHEDULE) via DOM.

Why this version?
- WebForms + UpdatePanel only renders the scheduling controls after a real in-page
  "Search" postback. Fetching the delta doesn't update the DOM, so we now do the
  search by clicking the Search button in the page.

Flow:
  1) Navigate to the page (IWA/NTLM handled by the browser, optional storage_state).
  2) Fill OpCo/Product/Serial, click Search (DOM postback).
  3) Wait for result message and for schedule controls (if eligible).
  4) Fill Date/Time/Timezone, click Schedule (DOM postback or full reload).
  5) Read message and write CSV row.

Env:
  FIRMWARE_INPUT_XLSX=data/firmware_schedule.csv
  FIRMWARE_STORAGE_STATE=storage_state.json
  FIRMWARE_BROWSER_CHANNEL=msedge
  FIRMWARE_AUTH_ALLOWLIST=*.fujixerox.net,*.xerox.com
  FIRMWARE_OPCO=FXAU
  FIRMWARE_HEADLESS=true
  # Optional:
  FIRMWARE_TIME_VALUE=03
  FIRMWARE_DAYS_MIN=3
  FIRMWARE_DAYS_MAX=6
  FIRMWARE_DEBUG_TZ=0
"""

from __future__ import annotations

import asyncio
import contextlib
import csv
import os
import random
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from bs4 import BeautifulSoup
from playwright.async_api import async_playwright, Error as PWError  # type: ignore

# ---------- Constants ----------
BASE = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"
URL = f"{BASE}/firmware/SingleRequest.aspx"

DEFAULT_OPCO = os.getenv("FIRMWARE_OPCO", "FXAU")
INPUT_PATH = Path(os.getenv("FIRMWARE_INPUT_XLSX", "data/firmware_schedule.csv"))
STORAGE_STATE_PATH = Path(os.getenv("FIRMWARE_STORAGE_STATE", "storage_state.json"))
BROWSER_CHANNEL = os.getenv("FIRMWARE_BROWSER_CHANNEL", "msedge") or None
ALLOWLIST = os.getenv("FIRMWARE_AUTH_ALLOWLIST", "*.fujixerox.net,*.xerox.com")
HEADLESS = os.getenv("FIRMWARE_HEADLESS", "true").lower() in {"1", "true", "yes"}
DEBUG_TZ = os.getenv("FIRMWARE_DEBUG_TZ", "0").lower() in {"1", "true", "yes"}

PREFERRED_TIME_VALUE = (os.getenv("FIRMWARE_TIME_VALUE", "03") or "03").strip()
DAYS_MIN = int(os.getenv("FIRMWARE_DAYS_MIN", "3"))
DAYS_MAX = int(os.getenv("FIRMWARE_DAYS_MAX", "6"))

# AU state → timezone dropdown value
STATE_TZ = {
    "ACT": "+11:00",
    "NSW": "+11:00",
    "VIC": "+11:00",
    "TAS": "+11:00",
    "QLD": "+10:00",
    "SA": "+10:30",
    "NT": "+09:30",
}


# ---------- CSV input ----------
def normalize_row(raw: dict) -> dict:
    lower = {
        (k or "").strip().lower(): "" if v is None else str(v).strip()
        for k, v in raw.items()
    }

    def get(*names: str) -> str:
        for n in names:
            if n in lower and lower[n]:
                return lower[n]
        return ""

    return {
        "serial": get("serial", "serialnumber", "serial_number"),
        "product_code": get("product_code", "product", "productcode"),
        "state": get("state", "region").upper(),
        "opco": get("opco", "opcoid", "opco_id") or DEFAULT_OPCO,
    }


def read_rows(path: Path) -> Iterable[dict]:
    if not path.exists():
        raise FileNotFoundError(f"Input not found: {path}")

    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                yield normalize_row(row)
        return

    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        try:
            from openpyxl import load_workbook  # type: ignore
        except Exception as exc:
            raise RuntimeError("openpyxl is required for .xlsx files") from exc

        wb = load_workbook(path, read_only=True)
        try:
            ws = wb.active
            if ws is None:
                raise RuntimeError(f"No active worksheet in workbook: {path}")
            header_cells = next(ws.iter_rows(min_row=1, max_row=1))
            headers = [
                "" if c.value is None else str(c.value).strip() for c in header_cells
            ]
            for cells in ws.iter_rows(min_row=2, values_only=True):
                row = {
                    (headers[i] or f"col{i}"): ("" if v is None else str(v).strip())
                    for i, v in enumerate(cells)
                }
                yield normalize_row(row)
        finally:
            wb.close()
        return

    raise ValueError(f"Unsupported input type: {path.suffix}")


# ---------- Helpers ----------
def pick_schedule_date() -> str:
    start = datetime.now().date() + timedelta(days=DAYS_MIN)
    end = datetime.now().date() + timedelta(days=DAYS_MAX)
    span = (end - start).days
    offset = random.randint(0, max(0, span))
    return (start + timedelta(days=offset)).strftime("%Y-%m-%d")  # ISO


def timezone_for_state(state: str) -> str:
    return STATE_TZ.get(state.upper(), "+11:00")


async def debug_dump_timezone(page) -> None:
    if not DEBUG_TZ:
        return
    data = await page.evaluate("""
      () => {
        const sel = document.querySelector('#MainContent_ddlTimeZone');
        if (!sel) return {selected: null, options: []};
        return {
          selected: { value: sel.value, text: sel.options[sel.selectedIndex]?.text || '' },
          options: Array.from(sel.options).map(o => ({ value: o.value, text: o.text }))
        };
      }
    """)
    print("TZ selected:", data.get("selected"))
    for o in data.get("options", []):
        print("TZ option:", o)


def parse_status_from_page_html(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for sel in [
        "#MainContent_MessageLabel",
        "#MainContent_lblMessage",
        "#MainContent_lblStatus",
    ]:
        node = soup.select_one(sel)
        if node:
            return " ".join(node.get_text(" ", strip=True).split())
    return ""


async def read_status_from_dom(page) -> str:
    for sel in [
        "#MainContent_MessageLabel",
        "#MainContent_lblMessage",
        "#MainContent_lblStatus",
    ]:
        try:
            el = await page.query_selector(sel)
        except PWError:
            el = None
        if el:
            try:
                txt = (await el.inner_text()).strip()
            except PWError:
                txt = ""
            if txt:
                return " ".join(txt.split())
    return ""


async def _race_and_cancel(*aws, timeout: float | None = None):
    tasks = [asyncio.create_task(coro) for coro in aws]
    try:
        done, pending = await asyncio.wait(
            tasks, return_when=asyncio.FIRST_COMPLETED, timeout=timeout
        )
        for t in pending:
            with contextlib.suppress(Exception):
                t.cancel()
        if done:
            return await next(iter(done))
        return None
    finally:
        for t in tasks:
            if not t.done():
                with contextlib.suppress(Exception):
                    t.cancel()


# ---------- DOM actions ----------
async def fill_search_fields(page, opco: str, product_code: str, serial: str) -> None:
    # Ensure we're on the page
    await page.goto(URL, wait_until="domcontentloaded")

    # Set values and fire change/input so WebForms tracks state
    await page.evaluate(
        """
        ({opco, product, serial}) => {
          const setVal = (sel, val) => {
            const el = document.querySelector(sel);
            if (!el) return;
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
          };

          // IDs from earlier inspection/capture
          setVal('select#MainContent_ddlOpCoID', opco);
          setVal('#MainContent_ProductCode', product);
          setVal('#MainContent_SerialNumber', serial);
        }
        """,
        {"opco": opco, "product": product_code, "serial": serial},
    )


async def click_search(page) -> None:
    selectors = [
        'input[name="ctl00$MainContent$btnSearch"]',
        "#MainContent_btnSearch",
        '//input[contains(@value,"Search")]',
        '//button[contains(normalize-space(),"Search")]',
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel)
            if await loc.count() == 0:
                continue
            await loc.first.scroll_into_view_if_needed()
            await loc.first.click(timeout=3000)
            return
        except Exception:
            continue
    raise RuntimeError("Search button not found")


async def wait_after_search(page) -> str:
    # After clicking Search, the page may do UpdatePanel partial update (no nav)
    # or it may update various bits. We wait for either a message or the date box.
    try:
        await _race_and_cancel(
            page.wait_for_selector(
                "#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus",
                timeout=12000,
            ),
            page.wait_for_selector("#MainContent_txtDateTime", timeout=12000),
            timeout=12,
        )
    except Exception:
        pass

    msg = await read_status_from_dom(page)
    if not msg:
        try:
            html = await page.content()
            msg = parse_status_from_page_html(html) or msg
        except PWError:
            pass
    return msg


async def wait_for_schedule_controls(page):
    # Controls appear only if the device is eligible to schedule
    await page.wait_for_selector(
        "#MainContent_txtDateTime, #MainContent_ddlScheduleTime, #MainContent_ddlTimeZone",
        timeout=12000,
    )
    # If any is present but disabled, try to enable
    await page.evaluate("""
      () => {
        for (const sel of ['#MainContent_txtDateTime', '#MainContent_ddlScheduleTime', '#MainContent_ddlTimeZone']) {
          const el = document.querySelector(sel);
          if (el && el.disabled) el.disabled = false;
        }
      }
    """)


async def fill_schedule_fields(page, date_iso: str, time_val: str, tz_val: str) -> None:
    await page.evaluate(
        """
        ({date, time, tz}) => {
          const setVal = (sel, val) => {
            const el = document.querySelector(sel);
            if (!el) return;
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
          };
          setVal('#MainContent_txtDateTime', date);
          setVal('select#MainContent_ddlScheduleTime', time);
          setVal('select#MainContent_ddlTimeZone', tz);
        }
        """,
        {"date": date_iso, "time": time_val, "tz": tz_val},
    )


async def click_schedule(page) -> None:
    selectors = [
        'input[name="ctl00$MainContent$submitButton"]',
        "#MainContent_submitButton",
        'input[type="submit"][value="Schedule"]',
        'input[value="Schedule"]',
        "//input[contains(@value,'Schedule')]",
        "//button[contains(normalize-space(),'Schedule')]",
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel)
            if await loc.count() == 0:
                continue
            await loc.first.scroll_into_view_if_needed()
            # race: message OR reload
            nav = page.wait_for_load_state("domcontentloaded", timeout=15000)
            msg = page.wait_for_selector(
                "#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus",
                timeout=15000,
            )
            await loc.first.click(timeout=3000)
            with contextlib.suppress(Exception):
                await _race_and_cancel(nav, msg, timeout=15)
            return
        except Exception:
            continue

    # Manual full postback fallback
    await page.evaluate(
        """
        (eventTarget) => {
          var f = document.forms && document.forms[0];
          if (!f) return;
          var et = f.__EVENTTARGET || f.querySelector('input[name="__EVENTTARGET"]');
          if (!et) { et = document.createElement('input'); et.type='hidden'; et.name='__EVENTTARGET'; f.appendChild(et); }
          et.value = eventTarget;
          var ea = f.__EVENTARGUMENT || f.querySelector('input[name="__EVENTARGUMENT"]');
          if (!ea) { ea = document.createElement('input'); ea.type='hidden'; ea.name='__EVENTARGUMENT'; f.appendChild(ea); }
          ea.value = '';
          f.submit();
        }
        """,
        "ctl00$MainContent$submitButton",
    )
    with contextlib.suppress(Exception):
        await page.wait_for_load_state("domcontentloaded", timeout=15000)


async def wait_after_schedule(page) -> str:
    try:
        await _race_and_cancel(
            page.wait_for_selector(
                "#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus",
                timeout=15000,
            ),
            page.wait_for_load_state("domcontentloaded", timeout=15000),
            timeout=15,
        )
    except Exception:
        pass
    msg = await read_status_from_dom(page)
    if not msg:
        try:
            html = await page.content()
            msg = parse_status_from_page_html(html) or msg
        except PWError:
            pass
    return msg


# ---------- Main ----------
async def main() -> None:
    out_path = INPUT_PATH.with_name(INPUT_PATH.stem + "_out.csv")
    browser_args = [
        f"--auth-server-allowlist={ALLOWLIST}",
        f"--auth-negotiate-delegate-allowlist={ALLOWLIST}",
    ]

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=HEADLESS, channel=BROWSER_CHANNEL, args=browser_args
        )
        context_kwargs: Dict[str, Any] = {}
        if STORAGE_STATE_PATH.exists():
            context_kwargs["storage_state"] = str(STORAGE_STATE_PATH)
        context = await browser.new_context(**context_kwargs)
        page = await context.new_page()
        page.set_default_navigation_timeout(45_000)
        page.set_default_timeout(45_000)

        rows = list(read_rows(INPUT_PATH))
        if not rows:
            print(f"No rows found in {INPUT_PATH}")
            await context.close()
            await browser.close()
            return

        with out_path.open("w", newline="", encoding="utf-8") as fout:
            fieldnames = [
                "serial",
                "product_code",
                "state",
                "opco",
                "http_status_search",
                "status_text_search",
                "http_status_schedule",
                "status_text_schedule",
                "scheduled_date",
                "scheduled_time",
                "timezone_value",
            ]
            writer = csv.DictWriter(fout, fieldnames=fieldnames)
            writer.writeheader()

            for item in rows:
                serial = item["serial"]
                product = item["product_code"]
                if not serial or not product:
                    continue

                # 1) SEARCH via DOM (ensures UpdatePanel actually updates the page)
                await fill_search_fields(
                    page, item.get("opco") or DEFAULT_OPCO, product, serial
                )
                await click_search(page)
                status_s = await wait_after_search(page)
                code_s = 200
                print(f"[SEARCH] {serial}/{product} -> {status_s or '(no message)'}")

                # If the page didn’t render schedule controls (e.g., not eligible), record and continue
                has_controls = False
                try:
                    await page.wait_for_selector(
                        "#MainContent_txtDateTime, #MainContent_ddlScheduleTime, #MainContent_ddlTimeZone",
                        timeout=3000,
                    )
                    # at least one exists; check if all three exist:
                    txt = await page.query_selector("#MainContent_txtDateTime")
                    tim = await page.query_selector("#MainContent_ddlScheduleTime")
                    tzz = await page.query_selector("#MainContent_ddlTimeZone")
                    has_controls = bool(txt and tim and tzz)
                except Exception:
                    has_controls = False

                if not has_controls:
                    writer.writerow(
                        {
                            "serial": serial,
                            "product_code": product,
                            "state": item.get("state", ""),
                            "opco": item.get("opco", ""),
                            "http_status_search": code_s,
                            "status_text_search": status_s,
                            "http_status_schedule": 200,
                            "status_text_schedule": "",  # no schedule possible
                            "scheduled_date": "",
                            "scheduled_time": "",
                            "timezone_value": "",
                        }
                    )
                    continue

                # 2) SCHEDULE (controls exist)
                await wait_for_schedule_controls(page)

                date_iso = pick_schedule_date()
                time_val = PREFERRED_TIME_VALUE
                desired_tz_val = timezone_for_state(item.get("state", ""))

                # Try selecting by value; give a helpful label hint for +11:00
                label_hint = (
                    "Canberra, Melbourne, Sydney"
                    if desired_tz_val == "+11:00"
                    else None
                )
                # Select tz in DOM (ensures ViewState tracks it)
                await page.evaluate(
                    """
                    (val) => {
                      const sel = document.querySelector('select#MainContent_ddlTimeZone');
                      if (sel) { sel.value = val; sel.dispatchEvent(new Event('change', { bubbles: true })); }
                    }
                    """,
                    desired_tz_val,
                )
                # If not applied, try a smarter selection routine
                applied_val = (
                    await page.eval_on_selector(
                        "select#MainContent_ddlTimeZone", "el => el ? el.value : ''"
                    )
                    or ""
                )
                if not applied_val or applied_val != desired_tz_val:
                    # Try matching by label
                    await page.evaluate(
                        """
                        (needle) => {
                          const sel = document.querySelector('select#MainContent_ddlTimeZone');
                          if (!sel) return;
                          const m = Array.from(sel.options).find(o => (o.text||'').toUpperCase().includes((needle||'').toUpperCase()));
                          if (m) { sel.value = m.value; sel.dispatchEvent(new Event('change', { bubbles: true })); }
                        }
                        """,
                        label_hint or "",
                    )

                actual_tz_val = (
                    await page.eval_on_selector(
                        "select#MainContent_ddlTimeZone", "el => el ? el.value : ''"
                    )
                    or desired_tz_val
                )

                if DEBUG_TZ:
                    await debug_dump_timezone(page)

                await fill_schedule_fields(page, date_iso, time_val, actual_tz_val)
                await click_schedule(page)

                status_c = await wait_after_schedule(page)
                code_c = 200
                print(f"[SCHEDULE] {serial}/{product} -> {status_c or '(no message)'}")

                writer.writerow(
                    {
                        "serial": serial,
                        "product_code": product,
                        "state": item.get("state", ""),
                        "opco": item.get("opco", ""),
                        "http_status_search": code_s,
                        "status_text_search": status_s,
                        "http_status_schedule": code_c,
                        "status_text_schedule": status_c,
                        "scheduled_date": date_iso,
                        "scheduled_time": time_val,
                        "timezone_value": actual_tz_val,
                    }
                )

        # Clean close—prevents “Task exception was never retrieved” noise
        await page.close()
        await context.close()
        await browser.close()

    print(f"Done. Wrote: {out_path}")


if __name__ == "__main__":
    asyncio.run(main())
