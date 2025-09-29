# EPGW-BusinessRules.py
# Skeleton for: login monitoring, row-by-row rule creation, completion ledgering

import argparse
import asyncio
import json
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, cast

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from playwright.async_api import (
    async_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

# ============
# Config area — fill these once you know the site/flow
# ============
BASE_URL = "https://YOUR-SITE.example.com"  # TODO
RULES_URL = f"{BASE_URL}/AddBusinessRule.aspx"  # TODO
ALLOWLIST = "*.yourdomain.example.com"  # for Windows IWA (Kerberos/NTLM)

# Selectors: update to your page
SEL_USERNAME = "#username"  # TODO (if form auth)
SEL_PASSWORD = "#password"  # TODO
SEL_LOGIN_BTN = "button[type=submit]"  # TODO
SEL_LOGIN_SENTINEL = "#mainNav"  # element that proves you’re logged in (TODO)

# Rule form selectors (example placeholders)
SEL_NEW_RULE_BTN = "#btnNewRule"  # TODO
SEL_RULE_NAME = "#ruleName"  # TODO
SEL_RULE_TYPE = "#ruleType"  # TODO
SEL_RULE_SAVE = "#btnSaveRule"  # TODO
SEL_TOAST_SUCCESS = ".toast-success"  # TODO (some confirmation element/text)

# ============
# Defaults/Paths
# ============
REPO_DIR = Path(__file__).parent.resolve()
USER_DATA_DIR = REPO_DIR / "user-data"  # persistent browser profile
USER_DATA_DIR.mkdir(exist_ok=True)
LOG_DIR = REPO_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
AUTH_STATUS_PATH = REPO_DIR / "auth_status.json"

NAV_TIMEOUT_MS = 45000
POST_ACTION_PAUSE_MS = 800  # short cushion between UI steps


# ============
# Data model for a Rule (adjust to match your spreadsheet columns)
# ============
@dataclass
class RuleRow:
    RuleName: str
    RuleType: str
    Extra1: Optional[str] = None
    Extra2: Optional[str] = None


# ============
# Excel helpers
# ============
def _get_ws_for_read(path: Path, sheet_name: Optional[str]) -> Worksheet:
    wb = load_workbook(path, data_only=True)
    ws_like = wb[sheet_name] if sheet_name else wb.active
    assert ws_like is not None, "Workbook has no active worksheet"
    ws = cast(Worksheet, ws_like)
    return ws


def _get_ws_for_write(wb: Workbook, sheet_name: str) -> Worksheet:
    ws_like = (
        wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
    )
    assert ws_like is not None, "Unable to resolve worksheet"
    ws = cast(Worksheet, ws_like)
    return ws


def read_rules_from_xlsx(
    path: Path, sheet_name: Optional[str] = None
) -> List[Dict[str, Any]]:
    ws = _get_ws_for_read(path, sheet_name)
    # Header row (coerce to str for stable typing/keys)
    header_cells = next(ws.iter_rows(min_row=1, max_row=1, values_only=False))
    headers: List[str] = [
        str(c.value) if c.value is not None else "" for c in header_cells
    ]

    rows: List[Dict[str, Any]] = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v in (None, "", 0) for v in r):
            continue
        row_dict: Dict[str, Any] = {}
        for i, h in enumerate(headers):
            row_dict[h] = r[i] if i < len(r) else None
        rows.append(row_dict)
    return rows


def append_completed_row(
    completed_path: Path,
    headers: List[str],
    row: Dict[str, Any],
    processed_at: datetime,
    status: str,
    sheet_name: str = "Completed",
) -> None:
    """
    Appends the original row + two extra columns: ProcessedAt (ISO) and Status.
    Creates workbook if it doesn't exist.
    """
    if completed_path.exists():
        wb = load_workbook(completed_path)
        ws = _get_ws_for_write(wb, sheet_name)
        # Ensure header row exists
        if ws.max_row == 0:
            ws.append(headers + ["ProcessedAt", "Status"])
    else:
        wb = Workbook()
        ws_any = wb.active
        assert ws_any is not None
        ws = cast(Worksheet, ws_any)
        ws.title = sheet_name
        ws.append(headers + ["ProcessedAt", "Status"])

    # Row values as a list in header order
    values: List[Any] = [row.get(h, "") for h in headers]
    values += [processed_at.isoformat(timespec="seconds"), status]
    ws.append(values)
    wb.save(completed_path)


def remove_processed_rows_inplace(
    path: Path,
    kept_rows: List[Dict[str, Any]],
    headers: List[str],
    sheet_name: Optional[str] = None,
) -> None:
    """Rewrite input file with only the kept (unprocessed or failed) rows."""
    wb = load_workbook(path)
    ws_like = wb[sheet_name] if sheet_name else wb.active
    assert ws_like is not None
    ws = cast(Worksheet, ws_like)

    # Clear all except header (delete from row 2 to the end)
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # Append kept rows as lists (not dicts)
    for r in kept_rows:
        ws.append([r.get(h, "") for h in headers])

    wb.save(path)


# ============
# Auth/Logging helpers
# ============
def setup_logging(verbosity: int) -> logging.Logger:
    level = logging.INFO if verbosity == 0 else logging.DEBUG
    logger = logging.getLogger("EPGW-BusinessRules")
    logger.setLevel(level)
    fh = logging.FileHandler(
        LOG_DIR / f"run-{datetime.now().strftime('%Y%m%d-%H%M%S')}.log",
        encoding="utf-8",
    )
    fh.setLevel(level)
    ch = logging.StreamHandler()
    ch.setLevel(level)
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)
    # Replace handlers on rerun (avoid duplicates in notebooks/VS Code)
    logger.handlers = [fh, ch]
    return logger


async def record_auth_status(page, logged_in: bool, note: str = "") -> None:
    """Write a small auth status JSON + screenshot for audit."""
    info = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "url": page.url,
        "logged_in": logged_in,
        "note": note,
    }
    AUTH_STATUS_PATH.write_text(json.dumps(info, indent=2), encoding="utf-8")
    # Screenshot for traceability
    shot = REPO_DIR / f"auth-{datetime.now().strftime('%Y%m%d-%H%M%S')}.png"
    try:
        await page.screenshot(path=str(shot), full_page=True)
    except Exception:
        pass


# ============
# Playwright core
# ============
async def get_browser_context(pw, headless: bool):
    """
    Persistent context (Edge channel if available) with IWA flags.
    """
    context = await pw.chromium.launch_persistent_context(
        user_data_dir=str(USER_DATA_DIR),
        headless=headless,
        channel="msedge",  # comment if Edge channel not installed
        args=[
            f"--auth-server-allowlist={ALLOWLIST}",
            f"--auth-negotiate-delegate-allowlist={ALLOWLIST}",
            "--start-minimized",
        ],
        accept_downloads=True,
    )
    return context


async def ensure_logged_in(page, logger: logging.Logger) -> bool:
    """
    Navigate to RULES_URL and confirm a known 'logged-in' sentinel exists.
    If your site uses classic form auth, add a one-time login flow here.
    """
    await page.goto(RULES_URL, wait_until="networkidle")
    try:
        await page.wait_for_selector(SEL_LOGIN_SENTINEL, timeout=8000)
        logger.info("Login OK (sentinel visible).")
        await record_auth_status(page, True, "Sentinel found")
        return True
    except PlaywrightTimeoutError:
        logger.warning("Sentinel not found; likely not logged in.")
        await record_auth_status(page, False, "Sentinel not found")
        return False


# ============
# Rule creation
# ============
async def create_rule_from_row(
    page, row: Dict[str, Any], logger: logging.Logger
) -> str:
    """
    Create a single rule from a spreadsheet row.
    Return "SUCCESS" or error text.
    """
    name = str(row.get("RuleName", "") or "").strip()
    rtype = str(row.get("RuleType", "") or "").strip()
    if not name or not rtype:
        return "ERROR: Missing RuleName or RuleType"

    try:
        # Start "new rule" flow
        await page.click(SEL_NEW_RULE_BTN)
        await page.wait_for_timeout(POST_ACTION_PAUSE_MS)

        # Fill fields
        await page.fill(SEL_RULE_NAME, name)
        # If RuleType is a dropdown:
        # await page.select_option(SEL_RULE_TYPE, rtype)
        # or a typed field:
        await page.fill(SEL_RULE_TYPE, rtype)

        # TODO: add any other fields here using row["Extra1"], row["Extra2"], etc.

        # Save
        async with page.expect_response(lambda r: r.ok, timeout=10000):
            await page.click(SEL_RULE_SAVE)

        # Confirm success (toast/text/sentinel)
        try:
            await page.wait_for_selector(SEL_TOAST_SUCCESS, timeout=8000)
        except PlaywrightTimeoutError:
            # If no toast, tolerate a lighter check (page settled)
            pass

        logger.info(f"Rule created: {name}")
        return "SUCCESS"

    except Exception as e:
        logger.exception(f"Failed creating rule: {name}")
        return f"ERROR: {e}"


# ============
# Orchestration
# ============
async def run_job(
    input_xlsx: Path,
    completed_xlsx: Path,
    sheet_name: Optional[str],
    headless: bool,
    max_rows: Optional[int],
    mutate_input: bool,
    logger: logging.Logger,
) -> None:
    rows = read_rules_from_xlsx(input_xlsx, sheet_name=sheet_name)
    if not rows:
        logger.warning("No rows found in input; nothing to do.")
        return

    headers = list(rows[0].keys())

    async with async_playwright() as p:
        context = await get_browser_context(p, headless=headless)
        page = await context.new_page()
        page.set_default_navigation_timeout(NAV_TIMEOUT_MS)
        page.set_default_timeout(NAV_TIMEOUT_MS)

        try:
            if not await ensure_logged_in(page, logger):
                logger.error("Login not established. Aborting.")
                return

            processed_count = 0
            kept_unprocessed: List[Dict[str, Any]] = []

            for row in rows:
                if max_rows is not None and processed_count >= max_rows:
                    kept_unprocessed.append(row)
                    continue

                status = await create_rule_from_row(page, row, logger)
                when = datetime.now()

                # Append to Completed ledger with stamp + status
                append_completed_row(completed_xlsx, headers, row, when, status)

                if status == "SUCCESS":
                    processed_count += 1
                else:
                    # keep it for retry if mutating input later
                    kept_unprocessed.append(row)

            logger.info(f"Processed OK: {processed_count} / {len(rows)}")

            # Optionally mutate the input: remove successfully processed rows
            if mutate_input and processed_count > 0:
                remove_processed_rows_inplace(
                    input_xlsx, kept_unprocessed, headers, sheet_name=sheet_name
                )
                logger.info("Input workbook updated (removed successful rows).")

        finally:
            await context.close()


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="EPGW Business Rules — row-by-row rule creator with login monitor and completion ledger."
    )
    ap.add_argument(
        "--input", required=True, help="Path to input XLSX (rows to create)."
    )
    ap.add_argument(
        "--completed", required=True, help="Path to completed XLSX (append-only)."
    )
    ap.add_argument("--sheet", default=None, help="Worksheet name (default: active).")
    ap.add_argument(
        "--headless",
        action="store_true",
        help="Run headless (keep off while testing/IWA).",
    )
    ap.add_argument(
        "--max-rows",
        type=int,
        default=None,
        help="Limit rows this run (useful for testing).",
    )
    ap.add_argument(
        "--mutate-input",
        action="store_true",
        help="Remove successful rows from input workbook.",
    )
    ap.add_argument(
        "-v", "--verbose", action="count", default=0, help="Increase log verbosity."
    )
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    logger = setup_logging(args.verbose)

    input_xlsx = Path(args.input).resolve()
    completed_xlsx = Path(args.completed).resolve()
    if not input_xlsx.exists():
        raise FileNotFoundError(f"Input not found: {input_xlsx}")

    asyncio.run(
        run_job(
            input_xlsx=input_xlsx,
            completed_xlsx=completed_xlsx,
            sheet_name=args.sheet,
            headless=args.headless,
            max_rows=args.max_rows,
            mutate_input=args.mutate_input,
            logger=logger,
        )
    )


if __name__ == "__main__":
    main()
