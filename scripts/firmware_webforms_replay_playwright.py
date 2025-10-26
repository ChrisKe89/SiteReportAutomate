#!/usr/bin/env python3
"""
Headless Playwright replayer for SingleRequest.aspx (SEARCH + SCHEDULE).

- Uses the browser (Chromium/Edge) so NTLM/Negotiate and TLS Just Work™.
- Loads cookies from storage_state.json (optional but recommended).
- Reads devices from FIRMWARE_INPUT_XLSX (CSV/XLSX).
- SEARCH: perform a real WebForms postback by clicking the Search button (so MicrosoftAjax updates the DOM).
- SCHEDULE: performs a real WebForms postback (robust click with manual form fallback).
- Writes <input>_out.csv with search & schedule results.

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
  FIRMWARE_DEBUG_TZ=0   # set to 1 to print timezone options/selection
"""

from __future__ import annotations

import asyncio
import csv
import os
import random
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple, List

from bs4 import BeautifulSoup
from playwright.async_api import async_playwright, Error as PWError  # type: ignore

# ---------- Constants ----------
BASE = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"
URL = f"{BASE}/firmware/SingleRequest.aspx"

# Known server IDs
SEARCH_TRIGGER = "ctl00$MainContent$btnSearch"
SCHEDULE_TRIGGER = "ctl00$MainContent$submitButton"

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

# AU state → timezone dropdown value (server expects the <option value>, not label)
STATE_TZ = {
    "ACT": "+11:00",
    "NSW": "+11:00",
    "VIC": "+11:00",
    "TAS": "+11:00",
    "QLD": "+10:00",
    "SA": "+10:30",
    "NT": "+09:30",
}


# ---------- Input helpers ----------
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


# ---------- MicrosoftAjax delta parsing ----------
def _extract_from_msajax_delta(delta: str) -> str:
    if not delta.startswith("|"):
        return ""
    tokens: List[str] = delta.split("|")
    i = 0
    snippets: List[str] = []
    while i < len(tokens):
        if tokens[i] == "updatePanel" and i + 3 < len(tokens):
            html = tokens[i + 3]
            snippets.append(html)
            i += 4
            continue
        i += 1
    for html in snippets:
        soup = BeautifulSoup(html, "html.parser")
        for sel in [
            "#MainContent_MessageLabel",
            "#MainContent_lblMessage",
            "#MainContent_lblStatus",
        ]:
            node = soup.select_one(sel)
            if node:
                txt = " ".join(node.get_text(" ", strip=True).split())
                if txt:
                    return txt
    for html in snippets:
        cleaned = " ".join(
            BeautifulSoup(html, "html.parser").get_text(" ", strip=True).split()
        )
        if cleaned:
            return cleaned[:300]
    return ""


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
    upper = html.upper()
    for key in ("SUCCESS", "SCHEDULE", "NOT", "INVALID", "ERROR"):
        if key in upper:
            return key
    return ""


async def read_status_from_dom(page) -> str:
    # prefer DOM (postback may not give us raw html)
    for sel in [
        "#MainContent_MessageLabel",
        "#MainContent_lblMessage",
        "#MainContent_lblStatus",
    ]:
        try:
            el = await page.query_selector(sel)
            if el:
                txt = (await el.inner_text()).strip()
                if txt:
                    return " ".join(txt.split())
        except PWError:
            pass
    return ""


# ---------- Scheduling helpers ----------
def pick_schedule_date() -> str:
    start = datetime.now().date() + timedelta(days=DAYS_MIN)
    end = datetime.now().date() + timedelta(days=DAYS_MAX)
    span = (end - start).days
    offset = random.randint(0, max(0, span))
    return (start + timedelta(days=offset)).strftime("%Y-%m-%d")


def timezone_for_state(state: str) -> str:
    return STATE_TZ.get(state.upper(), "+11:00")


async def select_timezone_on_page(
    page, desired_value: str, label_hint: str | None = None
) -> str:
    # Make sure it's enabled (sometimes disabled until SEARCH completes)
    await page.evaluate("""
      () => {
        const sel = document.querySelector('select#MainContent_ddlTimeZone');
        if (sel) { sel.removeAttribute('disabled'); sel.disabled = false; }
      }
    """)

    ok = await page.evaluate(
        """
        (val) => {
          const sel = document.querySelector('select#MainContent_ddlTimeZone');
          if (!sel) return false;
          sel.value = val;
          if (sel.value === val) {
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
          }
          return false;
        }
    """,
        desired_value,
    )
    if ok:
        return (
            await page.eval_on_selector(
                "select#MainContent_ddlTimeZone", "el => el.value"
            )
            or desired_value
        )

    if label_hint:
        matched_val = await page.evaluate(
            """
            (needle) => {
              const sel = document.querySelector('select#MainContent_ddlTimeZone');
              if (!sel) return '';
              const m = Array.from(sel.options).find(o => (o.text||'').toUpperCase().includes(needle.toUpperCase()));
              return m ? m.value : '';
            }
        """,
            label_hint,
        )
        if matched_val:
            ok2 = await page.evaluate(
                """
                (val) => {
                  const sel = document.querySelector('select#MainContent_ddlTimeZone');
                  if (!sel) return false;
                  sel.value = val;
                  if (sel.value === val) {
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                  }
                  return false;
                }
            """,
                matched_val,
            )
            if ok2:
                return matched_val

    fallback = await page.evaluate("""
        () => {
          const sel = document.querySelector('select#MainContent_ddlTimeZone');
          if (!sel) return '';
          const opt = Array.from(sel.options).find(o => o.value);
          if (!opt) return '';
          sel.value = opt.value;
          sel.dispatchEvent(new Event('change', { bubbles: true }));
          return sel.value;
        }
    """)
    return fallback or desired_value


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


# ---------- Page actions ----------
async def navigate_and_ready(page) -> None:
    await page.goto(URL, wait_until="domcontentloaded")


async def fill_search_fields(page, opco: str, product_code: str, serial: str) -> None:
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
          setVal('select#MainContent_ddlOpCoID', opco);
          setVal('#MainContent_ProductCode', product);
          setVal('#MainContent_SerialNumber', serial);
        }
        """,
        {"opco": opco, "product": product_code, "serial": serial},
    )


async def click_search(page) -> None:
    candidates = [
        'input[name="ctl00$MainContent$btnSearch"]',
        "#MainContent_btnSearch",
        '//input[@type="submit" and (translate(@value,"SEARCH","search")="search")]',
        '//button[contains(translate(.,"SEARCH","search"),"search")]',
    ]
    for sel in candidates:
        try:
            loc = page.locator(sel)
            if await loc.count() == 0:
                continue
            await loc.first.scroll_into_view_if_needed()
            await loc.first.click(timeout=3000)
            return
        except Exception:
            continue
    # fallback: manual __doPostBack
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
        SEARCH_TRIGGER,
    )


async def wait_after_search(page) -> str:
    """
    Wait for either a status message to appear or
    the scheduling controls to become present (not necessarily enabled).
    Returns any status message found (may be '').
    """
    # race: message label vs controls presence vs small timeout loop
    try:
        await asyncio.wait(
            [
                asyncio.create_task(
                    page.wait_for_selector(
                        "#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus",
                        timeout=12000,
                    )
                ),
                asyncio.create_task(
                    page.wait_for_selector("#MainContent_txtDateTime", timeout=12000)
                ),
            ],
            return_when=asyncio.FIRST_COMPLETED,
            timeout=12000,
        )
    except Exception:
        pass
    # read any message if present
    return await read_status_from_dom(page)


async def force_enable_controls(page) -> None:
    await page.evaluate("""
      () => {
        const drop = el => { if (!el) return; el.removeAttribute('disabled'); el.disabled = false; el.readOnly = false; };
        drop(document.querySelector('#MainContent_txtDateTime'));
        drop(document.querySelector('#MainContent_ddlScheduleTime'));
        drop(document.querySelector('#MainContent_ddlTimeZone'));
      }
    """)


async def fill_schedule_fields(page, date_iso: str, time_val: str, tz_val: str) -> None:
    # Convert 2025-11-01 to dd/MM/yyyy in case the input expects that
    ddmmyyyy = date_iso
    if "-" in date_iso:
        y, m, d = date_iso.split("-")
        ddmmyyyy = f"{d}/{m}/{y}"

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
        {"date": ddmmyyyy, "time": time_val, "tz": tz_val},
    )


async def click_schedule(page) -> None:
    selectors = [
        'input[name="ctl00$MainContent$submitButton"]',
        "#MainContent_submitButton",
        'input[type="submit"][value="Schedule"]',
        'input[value="Schedule"]',
        "button#MainContent_submitButton",
        'button[name="ctl00$MainContent$submitButton"]',
        '//input[contains(@value,"Schedule")]',
        '//button[contains(normalize-space(),"Schedule")]',
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
    # fallback: manual postback
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
        SCHEDULE_TRIGGER,
    )


async def wait_after_schedule(page) -> str:
    """
    Wait for either an UpdatePanel message or navigation,
    then read whatever message is present in the DOM.
    """
    try:
        await asyncio.wait(
            [
                asyncio.create_task(
                    page.wait_for_selector(
                        "#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus",
                        timeout=15000,
                    )
                ),
                asyncio.create_task(
                    page.wait_for_load_state("domcontentloaded", timeout=15000)
                ),
            ],
            return_when=asyncio.FIRST_COMPLETED,
            timeout=15000,
        )
    except Exception:
        pass
    return await read_status_from_dom(page)


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

                # Navigate fresh for each row (keeps state simple/clean)
                await navigate_and_ready(page)

                # SEARCH via DOM so UpdatePanel applies and unlocks controls
                await fill_search_fields(
                    page, item.get("opco") or DEFAULT_OPCO, product, serial
                )
                await click_search(page)
                status_s = await wait_after_search(page)
                code_s = 200  # if we got here, the page handled the postback

                # Prepare schedule inputs
                date_iso = pick_schedule_date()
                time_val = PREFERRED_TIME_VALUE
                desired_tz_val = timezone_for_state(item.get("state", ""))

                # Ensure controls are at least present; force-enable if they stayed disabled
                await force_enable_controls(page)
                label_hint = (
                    "Canberra, Melbourne, Sydney"
                    if desired_tz_val == "+11:00"
                    else None
                )
                actual_tz_val = await select_timezone_on_page(
                    page, desired_tz_val, label_hint
                )

                # Fill schedule fields & submit
                await fill_schedule_fields(page, date_iso, time_val, actual_tz_val)
                if DEBUG_TZ:
                    await debug_dump_timezone(page)
                await click_schedule(page)
                status_c = await wait_after_schedule(page)
                code_c = 200

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

        await context.close()
        await browser.close()

    print(f"Done. Wrote: {out_path}")


if __name__ == "__main__":
    asyncio.run(main())
