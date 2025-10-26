#!/usr/bin/env python3
"""
WebForms replayer for SingleRequest.aspx (SEARCH only).
- Reads devices from FIRMWARE_INPUT_XLSX (CSV or XLSX).
- Posts WebForms UpdatePanel payloads (no visible browser).
- Writes <input>_out.csv with HTTP/status columns.

Env you already have:
  FIRMWARE_INPUT_XLSX=data/firmware_schedule.csv   # or .xlsx
Optional envs:
  FIRMWARE_OPCO=FXAU
  FIRMWARE_DRY_RUN=false

Requires:
  pip install httpx beautifulsoup4
  (and if using .xlsx) pip install openpyxl
"""

from __future__ import annotations

import csv
import os
import time
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

import httpx
from bs4 import BeautifulSoup, Tag

# ----- Constants from your capture -----
BASE = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"
URL = f"{BASE}/firmware/SingleRequest.aspx"

# Search post had: ctl00$ScriptManager1=ctl00$MainContent$searchForm|ctl00$MainContent$btnSearch
SEARCH_PANEL = "ctl00$MainContent$searchForm"
SEARCH_TRIGGER = "ctl00$MainContent$btnSearch"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "X-Requested-With": "XMLHttpRequest",
    "X-MicrosoftAjax": "Delta=true",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Origin": BASE,
    "Referer": URL,
}

DEFAULT_OPCO = os.getenv("FIRMWARE_OPCO", "FXAU")
DRY_RUN = os.getenv("FIRMWARE_DRY_RUN", "false").lower() in {"1", "true", "yes"}

# ----- Input handling -----
INPUT_PATH = Path(os.getenv("FIRMWARE_INPUT_XLSX", "data/firmware_schedule.csv"))


def _as_str(value: Any) -> str:
    """Coerce soup/get() results to a plain string for type-checkers."""
    if value is None:
        return ""
    if isinstance(value, list):
        return ",".join(str(v) for v in value)
    return str(value)


def read_rows(path: Path) -> Iterable[dict]:
    """
    Yield dicts with keys: serial, product_code, state, opco
    Accepts CSV (header) or XLSX (first sheet, header row).
    """
    if not path.exists():
        raise FileNotFoundError(f"Input not found: {path}")

    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as f:
            r = csv.DictReader(f)
            for row in r:
                yield normalize_row(row)
        return

    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        # On-demand import to avoid mypy “assignment” issues.
        try:
            from openpyxl import load_workbook  # type: ignore
        except Exception as exc:  # pragma: no cover - import-time failure path
            raise RuntimeError(
                "openpyxl is required to read .xlsx files. Run: pip install openpyxl"
            ) from exc

        wb = load_workbook(path, read_only=True)
        try:
            ws = wb.active
            if ws is None:
                return
            # header row
            header_cells = next(ws.iter_rows(min_row=1, max_row=1))
            headers = [
                "" if c.value is None else str(c.value).strip() for c in header_cells
            ]
            for cells in ws.iter_rows(min_row=2, values_only=True):
                row = {
                    (headers[i] or f"col{i}").strip(): (
                        "" if v is None else str(v).strip()
                    )
                    for i, v in enumerate(cells)
                }
                yield normalize_row(row)
        finally:
            wb.close()
        return

    raise ValueError(f"Unsupported input type: {path.suffix}")


def normalize_row(raw: dict) -> dict:
    lower = {
        (k or "").strip().lower(): ("" if v is None else str(v).strip())
        for k, v in raw.items()
    }

    def get(*names: str) -> str:
        for n in names:
            if n in lower and lower[n]:
                return lower[n]
        return ""

    return {
        "serial": get("serial", "serialnumber", "serial_number"),
        "product_code": get("product_code", "product", "productcode", "product_code "),
        "state": get("state", "region"),
        "opco": get("opco", "opcoid", "opco_id") or DEFAULT_OPCO,
    }


# ----- WebForms helpers -----
def extract_hidden(html: str) -> Dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")

    def val(name: str) -> str:
        el = soup.find("input", {"name": name})
        if not el or not isinstance(el, Tag):
            return ""
        return _as_str(el.get("value"))

    fields: Dict[str, str] = {}
    for name in (
        "__VIEWSTATE",
        "__VIEWSTATEGENERATOR",
        "__EVENTVALIDATION",
        "__EVENTTARGET",
        "__EVENTARGUMENT",
        "__LASTFOCUS",
    ):
        fields[name] = val(name)
    return fields


def get_state(client: httpx.Client) -> Dict[str, str]:
    r = client.get(URL, headers={"User-Agent": HEADERS["User-Agent"]})
    r.raise_for_status()
    return extract_hidden(r.text)


def post_search(
    client: httpx.Client, opco: str, product_code: str, serial: str
) -> Tuple[int, str]:
    hidden = get_state(client)
    # mirrors your captured names
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
    if DRY_RUN:
        return 200, "DRY_RUN search"
    r = client.post(URL, data=form, headers=HEADERS, timeout=60.0)
    return r.status_code, r.text


def parse_status(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for sel in [
        "#MainContent_MessageLabel",
        "#MainContent_lblMessage",
        "#MainContent_lblStatus",
    ]:
        node = soup.select_one(sel)
        if node:
            return " ".join(node.get_text(" ", strip=True).split())
    # MicrosoftAjax partial updates often return pipe-delimited fragments
    if html.startswith("|"):  # MS AJAX delta format
        upper = html.upper()
        for key in ("SUCCESS", "SCHEDULE", "NOT", "INVALID", "ERROR"):
            if key in upper:
                return key
    return ""


def main() -> None:
    in_path = INPUT_PATH
    out_path = in_path.with_name(in_path.stem + "_out.csv")

    with (
        httpx.Client(follow_redirects=True, timeout=60.0) as client,
        open(out_path, "w", newline="", encoding="utf-8") as fout,
    ):
        rows = list(read_rows(in_path))
        if not rows:
            print(f"No rows found in {in_path}")
            return

        fieldnames = [
            "serial",
            "product_code",
            "state",
            "opco",
            "http_status_search",
            "status_text_search",
        ]
        writer = csv.DictWriter(fout, fieldnames=fieldnames)
        writer.writeheader()

        for item in rows:
            serial = item["serial"]
            product = item["product_code"]
            if not serial or not product:
                continue
            code, html = post_search(
                client, item.get("opco") or DEFAULT_OPCO, product, serial
            )
            status_text = parse_status(html)
            writer.writerow(
                {
                    "serial": serial,
                    "product_code": product,
                    "state": item.get("state", ""),
                    "opco": item.get("opco", ""),
                    "http_status_search": code,
                    "status_text_search": status_text,
                }
            )
            time.sleep(0.2)  # small politeness delay

    print(f"Done. Wrote: {out_path}")


if __name__ == "__main__":
    main()
