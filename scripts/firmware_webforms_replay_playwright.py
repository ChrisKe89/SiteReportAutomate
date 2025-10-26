#!/usr/bin/env python3
"""
Headless Playwright replayer for SingleRequest.aspx (SEARCH + SCHEDULE).

- Uses the browser (Chromium/Edge) so NTLM/Negotiate and TLS Just Work™.
- Loads cookies from storage_state.json (optional but recommended).
- Reads devices from FIRMWARE_INPUT_XLSX (CSV/XLSX).
- SEARCH: posts the UpdatePanel payload via in-page fetch() (browser cookies/auth).
- SCHEDULE: performs a real WebForms postback by clicking the submit button in the DOM.
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

# UpdatePanel targets (from captured posts)
SEARCH_PANEL = "ctl00$MainContent$searchForm"
SEARCH_TRIGGER = "ctl00$MainContent$btnSearch"
SCHEDULE_PANEL = "ctl00$MainContent$pnlFWTimesssss"
SCHEDULE_TRIGGER = "ctl00$MainContent$submitButton"

# Headers for XHR; Origin/Referer/Cookies are handled by the browser.
XHR_HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "X-Requested-With": "XMLHttpRequest",
    "X-MicrosoftAjax": "Delta=true",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
}

DEFAULT_OPCO = os.getenv("FIRMWARE_OPCO", "FXAU")
INPUT_PATH = Path(os.getenv("FIRMWARE_INPUT_XLSX", "data/firmware_schedule.csv"))
STORAGE_STATE_PATH = Path(os.getenv("FIRMWARE_STORAGE_STATE", "storage_state.json"))
BROWSER_CHANNEL = os.getenv("FIRMWARE_BROWSER_CHANNEL", "msedge") or None
ALLOWLIST = os.getenv("FIRMWARE_AUTH_ALLOWLIST", "*.fujixerox.net,*.xerox.com")
HEADLESS = os.getenv("FIRMWARE_HEADLESS", "true").lower() in {"1", "true", "yes"}

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


def parse_status_from_html(html: str) -> str:
    if html.startswith("|"):
        return _extract_from_msajax_delta(html)
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
    """
    Selects timezone in the DOM so VIEWSTATE reflects the selection.
    Returns the actually selected option value.
    """
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


# ---------- WebForms via Playwright ----------
async def get_state(page) -> Dict[str, str]:
    # Navigating ensures IWA handshake + the form exists
    await page.goto(URL, wait_until="domcontentloaded")
    names = [
        "__VIEWSTATE",
        "__VIEWSTATEGENERATOR",
        "__EVENTVALIDATION",
        "__EVENTTARGET",
        "__EVENTARGUMENT",
        "__LASTFOCUS",
    ]
    values: Dict[str, str] = {}
    for name in names:
        try:
            val = await page.eval_on_selector(f'input[name="{name}"]', "el => el.value")
        except PWError:
            val = ""
        values[name] = val or ""
    return values


def urlencode_form(d: Dict[str, str]) -> str:
    from urllib.parse import urlencode

    return urlencode(d)


async def post_search(
    page, hidden: Dict[str, str], opco: str, product_code: str, serial: str
) -> Tuple[int, str]:
    form = {
        "ctl00$ScriptManager1": f"{SEARCH_PANEL}|{SEARCH_TRIGGER}",
        "__EVENTTARGET": hidden.get("__EVENTTARGET", ""),
        "__EVENTARGUMENT": hidden.get("__EVENTARGUMENT", ""),
        "__LASTFOCUS": hidden.get("__LASTFOCUS", ""),
        "__VIEWSTATE": hidden.get("__VIEWSTATE", ""),
        "__VIEWSTATEGENERATOR": hidden.get("__VIEWSTATEGENERATOR", ""),
        "__EVENTVALIDATION": hidden.get("__EVENTVALIDATION", ""),
        "ctl00$MainContent$ddlOpCoID": opco,
        "ctl00$MainContent$ProductCode": product_code,
        "ctl00$MainContent$SerialNumber": serial,
        "ctl00$ucAsync1$hdnTimeout": "30",
        "ctl00$ucAsync1$hdnSetTimeoutID": "4",
        "__ASYNCPOST": "true",
        SEARCH_TRIGGER: "Search",
    }
    payload = urlencode_form(form)
    result = await page.evaluate(
        """async ({url, headers, body}) => {
            const res = await fetch(url, { method: 'POST', headers, body, credentials: 'include' });
            const text = await res.text();
            return { status: res.status, text };
        }""",
        {"url": URL, "headers": XHR_HEADERS, "body": payload},
    )
    return int(result["status"]), str(result["text"])


async def dom_submit_schedule(
    page,
    opco: str,
    product_code: str,
    serial: str,
    date_iso: str,
    time_val: str,
    tz_val: str,
) -> tuple[int, str]:
    """
    Set fields in the DOM and trigger the real WebForms postback.
    Returns (status_code, parsed_status_text).
    """
    # Ensure form is present
    await page.goto(URL, wait_until="domcontentloaded")

    # Fill fields via DOM (set value + change) to satisfy WebForms state tracking
    await page.evaluate(
        """
        ({opco, product, serial, date, time, tz}) => {
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
          setVal('#MainContent_txtDateTime', date);
          setVal('select#MainContent_ddlScheduleTime', time);
          setVal('select#MainContent_ddlTimeZone', tz);
        }
    """,
        {
            "opco": opco,
            "product": product_code,
            "serial": serial,
            "date": date_iso,
            "time": time_val,
            "tz": tz_val,
        },
    )

    # Click the real submit button (fires __doPostBack / UpdatePanel)
    await page.click('input[name="ctl00$MainContent$submitButton"]')

    async def read_status() -> str:
        for sel in [
            "#MainContent_MessageLabel",
            "#MainContent_lblMessage",
            "#MainContent_lblStatus",
        ]:
            el = await page.query_selector(sel)
            if el:
                txt = (await el.inner_text()).strip()
                if txt:
                    return " ".join(txt.split())
        return ""

    # Wait for the UpdatePanel refresh and a visible message
    for _ in range(50):
        msg = await read_status()
        if msg:
            return 200, msg
        await page.wait_for_timeout(200)

    # Fallback: call succeeded but no visible message extracted
    return 200, ""


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

                # 1) Fresh hidden fields, then SEARCH (XHR delta parsing)
                hidden = await get_state(page)
                code_s, html_s = await post_search(
                    page, hidden, item.get("opco") or DEFAULT_OPCO, product, serial
                )
                status_s = parse_status_from_html(html_s)

                # 2) Prepare schedule inputs
                date_iso = pick_schedule_date()
                time_val = PREFERRED_TIME_VALUE
                desired_tz_val = timezone_for_state(item.get("state", ""))

                # Ensure timezone is actually selected in DOM (so VIEWSTATE agrees)
                await page.goto(URL, wait_until="domcontentloaded")
                label_hint = (
                    "Canberra, Melbourne, Sydney"
                    if desired_tz_val == "+11:00"
                    else None
                )
                actual_tz_val = await select_timezone_on_page(
                    page, desired_tz_val, label_hint
                )

                # 3) SCHEDULE via real postback
                code_c, status_c = await dom_submit_schedule(
                    page,
                    item.get("opco") or DEFAULT_OPCO,
                    product,
                    serial,
                    date_iso,
                    time_val,
                    actual_tz_val,
                )

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
