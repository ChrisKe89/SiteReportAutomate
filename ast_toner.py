import asyncio
from openpyxl import load_workbook, Workbook
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup

# Input and output file paths
INPUT_XLSX = "input.xlsx"  # Your cleaned spreadsheet
OUTPUT_XLSX = "output.xlsx"  # Where results will be saved

# Selectors for the input fields and search button
SELECTORS = {
    "product_family": "#MainContent_txtProductFamily",
    "product_code": "#MainContent_txtProductCode",
    "serial_number": "#MainContent_txtSerialNumber",
    "search_btn": "#MainContent_btnSearch",
    "results_table": "#MainContent_gvResults",
}

PAGE_URL = "https://sgpaphq-epbbcs3.dc01.fujixerox.net/rdhc/PartStatuses.aspx"


def parse_html_table(html):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.select_one("table")
    if table is None:
        raise ValueError("No <table> found in the HTML!")
    headers = [th.get_text(strip=True) for th in table.select("tr th")]
    rows = []
    for tr in table.select("tr")[1:]:
        cells = [td.get_text(strip=True) for td in tr.select("td")]
        if cells:
            rows.append(cells)
    return headers, rows


async def main():
    wb = load_workbook(INPUT_XLSX)
    ws = wb.active

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Results"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(storage_state="storage_state.json")
        page = await context.new_page()
        await page.goto(PAGE_URL)

        for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            serial_number = str(row[0].value).strip() if row[0].value else ""
            product_code = str(row[1].value).strip() if row[1].value else ""
            product_family = str(row[6].value).strip() if row[6].value else ""

            await page.fill(SELECTORS["product_family"], product_family)
            await page.fill(SELECTORS["product_code"], product_code)
            await page.fill(SELECTORS["serial_number"], serial_number)
            await page.click(SELECTORS["search_btn"])
            await page.wait_for_timeout(3000)

            table_html = await page.inner_html(SELECTORS["results_table"])
            headers, data_rows = parse_html_table(table_html)

            # Write headers if first row
            if idx == 2:
                out_ws.append(
                    ["Input Row", "Product Family", "Product Code", "Serial Number"]
                    + headers
                )

            # Write each result row with input values
            for data_row in data_rows:
                out_ws.append(
                    [idx, product_family, product_code, serial_number] + data_row
                )

        out_wb.save(OUTPUT_XLSX)
        await browser.close()


# To run in a notebook or interactive environment, use:
# asyncio.get_event_loop().run_until_complete(main())
# Otherwise, use:
if __name__ == "__main__":
    asyncio.run(main())
