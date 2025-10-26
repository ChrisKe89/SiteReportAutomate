# save as scripts\webforms_replay.py
import csv
import re
from pathlib import Path

import httpx
from bs4 import BeautifulSoup
from bs4.element import Tag

BASE = "https://sgpaphq-epbbcs3.dc01.fujixerox.net"
URL = f"{BASE}/firmware/SingleRequest.aspx"

INPUT = Path("devices.csv")  # serial,product_code,state,opco (or adjust)
OUTPUT = Path("devices_out.csv")


def extract_hidden(html: str) -> dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")

    def val(name: str) -> str:
        element = soup.find("input", {"name": name})
        if isinstance(element, Tag):
            return element.attrs.get("value", "")
        return ""

    fields: dict[str, str] = {}
    for name in [
        "__VIEWSTATE",
        "__EVENTVALIDATION",
        "__VIEWSTATEGENERATOR",
        "__EVENTTARGET",
        "__EVENTARGUMENT",
    ]:
        v = val(name)
        if v:
            fields[name] = v
    return fields


def parse_result(html: str) -> dict:
    # TODO: tailor to the page — example grabs a message label if present
    soup = BeautifulSoup(html, "html.parser")
    msg = soup.select_one(
        "#MainContent_MessageLabel, #MainContent_lblMessage, #MainContent_lblStatus"
    )
    status = re.sub(r"\s+", " ", msg.get_text(strip=True)) if msg else ""
    return {"status_text": status}


def make_form(
    hidden: dict, serial: str, product_code: str, state: str, opco: str
) -> dict:
    # Replace these names with the exact ones from your HAR dump:
    form = {
        "__EVENTTARGET": hidden.get("__EVENTTARGET", ""),
        "__EVENTARGUMENT": hidden.get("__EVENTARGUMENT", ""),
        "__VIEWSTATE": hidden.get("__VIEWSTATE", ""),
        "__VIEWSTATEGENERATOR": hidden.get("__VIEWSTATEGENERATOR", ""),
        "__EVENTVALIDATION": hidden.get("__EVENTVALIDATION", ""),
        # INPUTS — adjust names to match your page
        "MainContent_txtSerial": serial,
        "MainContent_txtProductCode": product_code,
        "MainContent_txtOPCO": opco,  # if present
        "MainContent_txtState": state,  # if present
        # If the flow requires clicking a specific button, WebForms expects the button's
        # name in the form with a non-empty value:
        "MainContent_btnSearch": "Search",  # or MainContent_btnRequest / etc.
    }
    return form


def main():
    with httpx.Client(
        headers={"User-Agent": "Mozilla/5.0"}, timeout=30.0, follow_redirects=True
    ) as s:
        # 1) Warm-up GET to collect cookies + hidden fields
        r = s.get(URL)
        r.raise_for_status()
        hidden = extract_hidden(r.text)

        # Optional: if BASIC auth gateway is used before this page, you can warm it:
        # s.get("http://epgateway.sgp.xerox.com:8041/AlertManagement/businessrule.aspx", auth=("USERNAME","PASSWORD"))

        # 2) Iterate devices
        out_exists = OUTPUT.exists() and OUTPUT.stat().st_size > 0
        with OUTPUT.open("a", newline="", encoding="utf-8") as fout:
            writer = None
            with INPUT.open(newline="", encoding="utf-8-sig") as fin:
                reader = csv.DictReader(fin)
                for row in reader:
                    serial = (
                        row.get("serial") or row.get("SerialNumber") or ""
                    ).strip()
                    product_code = (
                        row.get("product_code") or row.get("Product_Code") or ""
                    ).strip()
                    state = (row.get("state") or "").strip().upper()
                    opco = (row.get("opco") or row.get("OpcoID") or "").strip()

                    if not serial or not product_code:
                        continue

                    # Fresh hidden fields per request is safest on WebForms
                    r0 = s.get(URL)
                    r0.raise_for_status()
                    hidden = extract_hidden(r0.text)

                    form = make_form(hidden, serial, product_code, state, opco)
                    r1 = s.post(URL, data=form)
                    result = parse_result(r1.text)

                    out_row = {
                        "serial": serial,
                        "product_code": product_code,
                        "state": state,
                        "opco": opco,
                        "http_status": r1.status_code,
                        **result,
                    }

                    if writer is None:
                        writer = csv.DictWriter(fout, fieldnames=list(out_row.keys()))
                        if not out_exists:
                            writer.writeheader()
                            out_exists = True
                    writer.writerow(out_row)


if __name__ == "__main__":
    main()
