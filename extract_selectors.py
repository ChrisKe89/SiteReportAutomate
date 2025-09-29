# extract_selectors.py
# Usage:
#   py extract_selectors.py path\to\page.html  > selectors.csv
#   (or) type page.html | py extract_selectors.py -  > selectors.csv

import sys
import csv
import re
from pathlib import Path
from typing import List, Optional, Iterable

from bs4 import BeautifulSoup
from bs4.element import Tag

INTERESTING = {"input", "select", "textarea", "button", "a"}


def text_snip(s: str, limit: int = 60) -> str:
    s = re.sub(r"\s+", " ", (s or "")).strip()
    return s[:limit]


def safe_classes(el: Tag) -> List[str]:
    """Return element class list as List[str] (never None)."""
    raw = el.get("class")
    if isinstance(raw, list):
        # BeautifulSoup can give List[str]
        return [c for c in raw if isinstance(c, str)]
    if isinstance(raw, str):
        # Rare, but normalize
        return [raw]
    return []


def previous_element_siblings(el: Tag) -> Iterable[Tag]:
    """Yield previous siblings that are Tag (skip strings/comments)."""
    for sib in el.previous_siblings:
        if isinstance(sib, Tag):
            yield sib


def nth_of_type(el: Tag) -> int:
    n = 1
    for sib in previous_element_siblings(el):
        if sib.name == el.name:
            n += 1
    return n


def css_piece(el: Tag) -> str:
    classes = ".".join(safe_classes(el))
    base = el.name + (("." + classes) if classes else "")
    # Prefer name attr for form controls
    if el.has_attr("name"):
        return f'{base}[name="{el.get("name")}"]'
    # Fall back to nth-of-type for stability
    idx = nth_of_type(el)
    return f"{base}:nth-of-type({idx})"


def css_selector(el: Tag) -> str:
    # If it has an ID, that's the best selector
    if el.has_attr("id"):
        return f"#{el.get('id')}"

    # Build a short path up to 3 ancestors
    parts: List[str] = [css_piece(el)]
    p = el.parent
    hops = 0
    while isinstance(p, Tag) and p.name not in {"html"} and hops < 3:
        if p.has_attr("id"):
            parts.append(f"#{p.get('id')}")
            break
        parts.append(css_piece(p))
        p = p.parent
        hops += 1
    return " > ".join(reversed(parts))


def xpath(el: Tag) -> str:
    parts: List[str] = []
    cur: Optional[Tag] = el
    while isinstance(cur, Tag):
        idx = nth_of_type(cur)
        parts.append(f"{cur.name}[{idx}]")
        parent = cur.parent
        cur = parent if isinstance(parent, Tag) else None
    return "/" + "/".join(reversed(parts))


def should_include(el: Tag) -> bool:
    if el.has_attr("id"):
        return True
    return el.name in INTERESTING


def load_html(src: str) -> str:
    if src == "-" or not src:
        return sys.stdin.read()
    p = Path(src)
    return p.read_text(encoding="utf-8", errors="ignore")


def main() -> None:
    src = sys.argv[1] if len(sys.argv) > 1 else "-"
    html = load_html(src)
    soup = BeautifulSoup(html, "html.parser")

    writer = csv.writer(sys.stdout, lineterminator="\n")
    writer.writerow(
        [
            "order",
            "tag",
            "id",
            "name",
            "type",
            "classes",
            "text",
            "css_selector",
            "xpath",
        ]
    )

    order = 0
    # DOM order
    for node in soup.find_all(True):  # returns List[Tag]
        el: Tag = node  # help type checkers
        if not should_include(el):
            continue
        order += 1

        tag = el.name
        id_val = el.get("id") or ""
        name_val = el.get("name") or ""
        typ = el.get("type") or ""
        classes_list = safe_classes(el)
        classes = " ".join(classes_list)

        txt = ""
        if tag in {"button", "a", "option", "label"}:
            txt = text_snip(el.get_text())

        selector = css_selector(el)
        xpth = xpath(el)

        writer.writerow(
            [order, tag, id_val, name_val, typ, classes, txt, selector, xpth]
        )


if __name__ == "__main__":
    main()
