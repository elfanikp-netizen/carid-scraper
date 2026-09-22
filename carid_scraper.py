#!/usr/bin/env python3
"""
carid_scraper.py
Look up part numbers (Partslink) from an Excel file on carid.com and write
Oldest Year / Newest Year / Brand / Model / Type / Interchange Number / OEM Number
to a new Excel file.

HOW CLOUDFLARE IS HANDLED
  The script opens a normal, visible Google Chrome window and controls it through
  Chrome's DevTools port. If Cloudflare shows a "verify you are human" screen you
  click it yourself; the script waits, then carries on. The Chrome profile is kept
  in ./chrome_profile so the clearance cookie is reused on later runs.
  Nothing here tries to auto-solve or spoof the check.

USAGE
  python carid_scraper.py --input parts.xlsx --output results.xlsx
  python carid_scraper.py --input parts.xlsx --limit 3 --debug     # test run
  python carid_scraper.py --parse-file debug/XYZ_product.html      # test parser offline
"""

import argparse
import json
import os
import platform
import random
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path
from urllib.parse import quote_plus

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# --------------------------------------------------------------------------- #
# SETTINGS  (edit here if the site layout differs from what the script expects)
# --------------------------------------------------------------------------- #
BASE = "https://www.carid.com"
CDP_PORT = 9222
PROFILE_DIR = Path("chrome_profile").resolve()
DEBUG_DIR = Path("debug")

# CARiD uses a specific visible input with id="search-field".
# Prefer that exact field and only fall back to generic inputs if it is absent.
SEARCH_INPUT_SELECTORS = [
    'input[id="search-field"]',
    'input#search-field',
    'input[name="search-field"]',
    'input[type="search"]',
    'input[name="q"]',
    'input[name="search"]',
    'input[name="keywords"]',
    'input[placeholder*="Search" i]',
    'input[id*="search" i]',
]

# Elements that usually wrap the fitment ("This part fits ...") information
FITMENT_CONTAINER_SELECTORS = (
    '[id*="fit" i], [class*="fit" i], [id*="vehicle" i], [class*="vehicle" i], '
    '[id*="compat" i], [class*="compat" i], [class*="application" i]'
)

OUTPUT_COLUMNS = [
    "Partslink Number", "Oldest Year", "Newest Year", "Brand", "Model", "Type",
    "Interchange Number", "Interchange 1", "Interchange 2", "Interchange 3", "Interchange 4", "Interchange 5",
    "OEM Number", "OEM 1", "OEM 2", "OEM 3", "OEM 4", "OEM 5",
    "Number Values", "Multiple values", "Part Brand", "Product URL", "Status",
]

KNOWN_MAKES = [
    "Acura", "Alfa Romeo", "American Motors", "Aston Martin", "Audi", "Austin Healey",
    "Avanti", "Bentley", "BMW", "BrightDrop", "Bugatti", "Buick", "Cadillac", "Chevy",
    "Chevrolet", "Chrysler", "Citroen", "Dacia", "Daewoo", "Dodge", "Eagle", "Ferrari",
    "Fiat", "Ford", "Freightliner", "Genesis", "GEO", "GMC", "Honda", "Hummer", "Hyundai",
    "Ineos", "Infiniti", "International", "Isuzu", "Jaguar", "Jeep", "Karma", "Kia",
    "Laforza", "Lamborghini", "Land Rover", "Lexus", "Lincoln", "Lordstown", "Lotus",
    "Lucid", "Mahindra", "Maserati", "Maybach", "Mazda", "McLaren", "Mercedes-Benz",
    "Mercedes", "Mercury", "MG", "Mini", "Mitsubishi", "Morgan", "Mullen", "Nikola",
    "Nissan", "Oldsmobile", "Opel", "Pagani", "Peugeot", "Plymouth", "Polestar", "Pontiac",
    "Porsche", "Ram", "Renault", "Rivian", "Rolls-Royce", "Rolls Royce", "Saab", "Saturn",
    "Scion", "Seat", "Skoda", "Slate", "Smart Car", "Smart", "Subaru", "Suzuki", "Tesla",
    "Toyota", "Triumph", "VinFast", "Volkswagen", "Volvo", "Workhorse",
]
_MAKE_RE = "|".join(re.escape(m) for m in sorted(KNOWN_MAKES, key=len, reverse=True))
_YEAR = r"((?:19|20)\d{2})"
_RANGE = _YEAR + r"(?:\s*(?:-|–|—|to|through)\s*" + _YEAR + r")?"
# "2015-2019 Ford F-150"
FIT_RE_A = re.compile(r"\b" + _RANGE + r"\s+(" + _MAKE_RE + r")\b[\s:,-]*(.*)", re.I)
# "Ford F-150 2015-2019"
FIT_RE_B = re.compile(r"\b(" + _MAKE_RE + r")\b\s+(.{1,60}?)\s+" + _RANGE + r"\b", re.I)


# --------------------------------------------------------------------------- #
# HTML PARSING  (works on saved HTML too, so you can test it offline)
# --------------------------------------------------------------------------- #
def _clean(s):
    return re.sub(r"\s+", " ", s or "").strip()


def _jsonld_product(soup):
    """Return the first schema.org Product object found, or {}."""
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or tag.get_text() or "")
        except Exception:
            continue
        stack = data if isinstance(data, list) else [data]
        while stack:
            item = stack.pop()
            if isinstance(item, list):
                stack.extend(item)
            elif isinstance(item, dict):
                t = item.get("@type")
                types = t if isinstance(t, list) else [t]
                if "Product" in types:
                    return item
                stack.extend(v for v in item.values() if isinstance(v, (dict, list)))
    return {}


def is_product_html(html):
    soup = BeautifulSoup(html, "lxml")
    if _jsonld_product(soup):
        return True
    og = soup.find("meta", property="og:type")
    return bool(og and "product" in (og.get("content") or "").lower())


def _extract_pairs(soup):
    """Collect 'Label: value' style pairs from tables, definition lists and text lines."""
    pairs = {}

    def put(k, v):
        k = _clean(k).strip(":# ").lower()
        v = _clean(v)
        if k and v and len(k) <= 40 and k not in pairs:
            pairs[k] = v

    for tr in soup.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        if len(cells) == 2:
            put(cells[0].get_text(" "), cells[1].get_text(" "))
    for dt in soup.find_all("dt"):
        dd = dt.find_next_sibling("dd")
        if dd:
            put(dt.get_text(" "), dd.get_text(" "))
    for el in soup.find_all(["li", "p", "div", "span"]):
        if el.find(["li", "p", "div", "table"]):
            continue
        m = re.match(r"^\s*([A-Za-z][A-Za-z /#.&-]{1,38}?)\s*:\s*(.+)$", el.get_text(" ", strip=True))
        if m:
            put(m.group(1), m.group(2))
    return pairs


def _numbers_from(raw):
    """Keep tokens that contain a digit (part numbers), preserve original formatting/dashes."""
    tokens = [t for t in re.split(r"[,;/|\s]+", raw or "") if re.search(r"\d", t)]
    seen, out = set(), []
    for t in tokens:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return "; ".join(out)


def _normalize_part_value(value):
    return re.sub(r"[^A-Za-z0-9]", "", (value or "")).upper()


def _sanitize_oem_value(value):
    if value is None:
        return ""
    value = str(value).strip()
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"[^A-Za-z0-9]", "", value)
    return value.upper()


def _strip_spaces_from_oem_text(value):
    if value is None:
        return ""
    raw = str(value).strip()
    if not raw:
        return ""

    pieces = []
    for chunk in re.split(r"[,;/|\n]+", raw):
        chunk = chunk.strip()
        if not chunk:
            continue
        cleaned = re.sub(r"\s+", "", chunk)
        if cleaned:
            pieces.append(cleaned)

    merged = []
    run = []
    for piece in pieces:
        is_short_fragment = bool(re.fullmatch(r"(?i)[A-Za-z0-9]{1,6}", piece)) and bool(re.search(r"\d", piece))
        if is_short_fragment:
            run.append(piece)
            continue
        if run:
            merged.append("".join(run))
            run = []
        merged.append(piece)
    if run:
        merged.append("".join(run))

    return "; ".join(merged)


def _is_valid_interchange(value):
    if value is None:
        return False
    s = str(value).strip()
    if len(s) < 8:
        return False
    return bool(re.fullmatch(r"(?i)[A-Za-z0-9]{2,4}-[A-Za-z0-9]{5,7}", s))


def _split_number_candidates(raw, part_number=""):
    """Return all interchange and OEM candidates, preserving each value's original formatting style."""
    if not raw:
        return []
    parts = [t.strip() for t in re.split(r"[,;/|\n]+", str(raw)) if t and t.strip()]
    if part_number:
        pn = _normalize_part_value(part_number)
        parts = [t for t in parts if _normalize_part_value(t) != pn]

    out = []
    seen = set()
    for p in parts:
        v = str(p).strip()
        if not re.search(r"\d", v):
            continue
        if _is_valid_interchange(v):
            key = _sanitize_oem_value(v)
            display = v
        else:
            key = _sanitize_oem_value(v)
            display = key
        if not key:
            continue
        if key not in seen:
            seen.add(key)
            out.append(display)
    return out


def _choose_oem_fields(raw, part_number=""):
    """Given a raw OE-like field, remove the exact partslink value and pick:
    - interchange = longest valid hyphenated number (xxx-xxxxx or xxx-xxxxxx)
    - OEM = the remaining value that is not hyphenated and not the partslink number,
      with spaces/symbols removed"""
    if not raw:
        return "", ""
    parts = [t.strip() for t in re.split(r"[,;/|\n]+", str(raw)) if t and t.strip()]
    if part_number:
        pn = _normalize_part_value(part_number)
        parts = [t for t in parts if _normalize_part_value(t) != pn]
    if not parts:
        return "", ""

    interchange_candidates = []
    for t in parts:
        v = t.strip()
        if _is_valid_interchange(v):
            interchange_candidates.append(v)
    interchange = ""
    if interchange_candidates:
        interchange = max(interchange_candidates, key=lambda s: len(s.split("-")[-1]))

    oem_candidates = []
    for t in parts:
        v = t.strip()
        if _is_valid_interchange(v):
            continue
        if _normalize_part_value(v) == _normalize_part_value(part_number):
            continue
        oem_candidates.append(v)

    oem = _sanitize_oem_value(oem_candidates[-1]) if oem_candidates else ""
    return interchange, oem


def _collect_unique_values(raw_value, part_number="", max_values=5, keep_interchanges=False):
    """Return up to max_values unique values from a raw text blob, preserving order."""
    values = []
    seen = set()
    for candidate in _split_number_candidates(raw_value, part_number):
        if _is_valid_interchange(candidate) and not keep_interchanges:
            continue
        if not _is_valid_interchange(candidate) and keep_interchanges:
            continue
        key = _sanitize_oem_value(candidate)
        if candidate and key and key not in seen:
            seen.add(key)
            values.append(candidate)
        if len(values) >= max_values:
            break
    return values


def _regex_numbers(text, *labels):
    for label in labels:
        for pattern in (
            rf"(?is)\b{re.escape(label)}\b(?:\s+(?:number|numbers|no\.?)?)?\s*[:\-]?\s*"
            rf"([A-Z0-9][A-Za-z0-9\-.,;/ ]{{0,200}})",
            rf"(?is)\b{re.escape(label)}\b.*?[:\-]\s*([A-Z0-9][A-Za-z0-9\-.,;/ ]{{0,200}})",
        ):
            m = re.search(pattern, text)
            if m:
                value = m.group(1)
                if value and re.search(r"\d", value):
                    return _numbers_from(value)
    return ""


def _pair_lookup(pairs, *needles):
    for key, val in pairs.items():
        if any(n in key for n in needles):
            return val
    return ""


def _fitment(soup):
    """Return (oldest, newest, brands, models, used_fallback)."""
    lines, used_fallback = [], False
    containers = soup.select(FITMENT_CONTAINER_SELECTORS)
    for c in containers:
        rows = c.find_all("tr")
        if rows:
            for r in rows:
                lines.append(_clean(" ".join(x.get_text(" ") for x in r.find_all(["th", "td"]))))
        lines.extend(_clean(x) for x in c.get_text("\n").split("\n"))
    entries = _match_fitment(lines)
    if not entries:
        used_fallback = True
        scope = soup.find("main") or soup.body or soup
        entries = _match_fitment(_clean(x) for x in scope.get_text("\n").split("\n"))

    years, brands, models = [], [], []
    for y1, y2, make, model in entries:
        years.extend([int(y1), int(y2 or y1)])
        if make and make not in brands:
            brands.append(make)
        if model and model not in models:
            models.append(model)
    if not years:
        return "", "", "", "", used_fallback
    return min(years), max(years), "; ".join(brands), "; ".join(models), used_fallback


def _match_fitment(lines):
    out = []
    for line in lines:
        if not line or len(line) > 200:
            continue
        m = FIT_RE_A.search(line)
        if m:
            y1, y2, make, rest = m.group(1), m.group(2), m.group(3), m.group(4)
            model = re.split(r"\s{2,}|\s[|•]\s", rest)[0].strip(" ,;:-")[:60]
            out.append((y1, y2, make.title() if make.islower() else make, model))
            continue
        m = FIT_RE_B.search(line)
        if m:
            make, model, y1, y2 = m.group(1), m.group(2), m.group(3), m.group(4)
            out.append((y1, y2, make, model.strip(" ,;:-")))
    return out


def _schema_product_numbers(ld):
    numbers = []
    is_similar = ld.get("isSimilarTo") or []
    if isinstance(is_similar, dict):
        is_similar = [is_similar]
    for item in is_similar:
        if isinstance(item, dict):
            mpn = item.get("mpn")
            if mpn:
                numbers.append(str(mpn))
    if isinstance(ld.get("mpn"), str):
        numbers.append(ld["mpn"])
    return "; ".join(numbers)


def parse_product(html, url="", part_number=""):
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n")
    pairs = _extract_pairs(soup)
    ld = _jsonld_product(soup)

    oe_pair_values = [
        val for key, val in pairs.items()
        if any(n in key for n in ("oe numbers", "oe number", "oem", "oem number", "oem numbers", "original equipment"))
    ]
    oe_raw = (
        " ; ".join(v for v in oe_pair_values if v)
        or _pair_lookup(pairs, "oe numbers", "oe number", "oem", "oem number", "oem numbers", "original equipment")
        or _regex_numbers(text, "oe numbers", "oe number", "oem number", "oem numbers", "original equipment")
        or _schema_product_numbers(ld)
    )
    oe_raw = _strip_spaces_from_oem_text(oe_raw)
    interchange_pair_values = [
        val for key, val in pairs.items()
        if any(n in key for n in ("interchange", "cross reference"))
    ]
    interchange_raw = (
        " ; ".join(v for v in interchange_pair_values if v)
        or _pair_lookup(pairs, "interchange", "cross reference")
        or _regex_numbers(text, "interchange number", "interchange numbers", "interchange")
    )
    if part_number:
        log_action("Part", f"Raw OE / interchange text for {part_number}: oe_raw={oe_raw!r}, interchange_raw={interchange_raw!r}")
    number_values = []
    seen_numbers = set()
    for raw_value in (oe_raw, interchange_raw):
        for candidate in _split_number_candidates(raw_value, part_number):
            key = _sanitize_oem_value(candidate)
            if candidate and key and key not in seen_numbers:
                seen_numbers.add(key)
                number_values.append(candidate)

    oe_values = [
        _sanitize_oem_value(v)
        for v in _collect_unique_values(oe_raw, part_number, max_values=5, keep_interchanges=False)
    ]
    interchange_values = [
        v.strip()
        for v in _collect_unique_values(interchange_raw, part_number, max_values=5, keep_interchanges=True)
    ]

    fallback_seen_oem = set(oe_values)
    fallback_seen_interchange = set(interchange_values)
    for candidate in number_values:
        value = str(candidate).strip()
        if not value:
            continue
        if _is_valid_interchange(value):
            clean = value.strip()
            if clean and clean not in fallback_seen_interchange:
                fallback_seen_interchange.add(clean)
                interchange_values.append(clean)
        else:
            clean = _sanitize_oem_value(value)
            if clean and clean not in fallback_seen_oem:
                fallback_seen_oem.add(clean)
                oe_values.append(clean)

    interchange, oem = _choose_oem_fields(oe_raw, part_number)

    if not oem and oe_raw:
        candidate = _numbers_from(oe_raw)
        if candidate:
            oem = _sanitize_oem_value(candidate)
    if interchange and not _is_valid_interchange(interchange):
        interchange = ""
    if not interchange and oe_raw:
        direct = (
            _numbers_from(_pair_lookup(pairs, "interchange", "cross reference"))
            or _regex_numbers(text, "interchange number", "interchange numbers", "interchange")
        )
        if _is_valid_interchange(direct):
            interchange = direct

    if not interchange:
        interchange = ""

    if not oem and oe_raw:
        for candidate in re.split(r"[,;/|\n]+", str(oe_raw)):
            cleaned = _sanitize_oem_value(candidate)
            if cleaned and cleaned != _normalize_part_value(part_number):
                oem = cleaned

    if not oe_values and oem:
        oe_values = [oem]
    if not interchange_values and interchange:
        interchange_values = [interchange]

    part_type = _pair_lookup(pairs, "part type", "product type", "type")
    if not part_type:
        part_type = _clean(ld.get("category", "")) if isinstance(ld.get("category", ""), str) else ""
    if not part_type:
        crumbs = [_clean(a.get_text()) for a in soup.select('[class*="breadcrumb" i] a')]
        part_type = crumbs[-1] if crumbs else ""

    brand_ld = ld.get("brand")
    if isinstance(brand_ld, dict):
        brand_ld = brand_ld.get("name")
    part_brand = _clean(brand_ld or _pair_lookup(pairs, "brand", "manufacturer"))

    oldest, newest, brands, models, fallback = _fitment(soup)

    row = {
        "Oldest Year": oldest,
        "Newest Year": newest,
        "Brand": brands,
        "Model": models,
        "Type": part_type,
        "Interchange Number": interchange_values[0] if interchange_values else interchange,
        "Interchange 1": interchange_values[0] if len(interchange_values) > 0 else "",
        "Interchange 2": interchange_values[1] if len(interchange_values) > 1 else "",
        "Interchange 3": interchange_values[2] if len(interchange_values) > 2 else "",
        "Interchange 4": interchange_values[3] if len(interchange_values) > 3 else "",
        "Interchange 5": interchange_values[4] if len(interchange_values) > 4 else "",
        "OEM Number": oe_values[0] if oe_values else _sanitize_oem_value(oem),
        "OEM 1": oe_values[0] if len(oe_values) > 0 else "",
        "OEM 2": oe_values[1] if len(oe_values) > 1 else "",
        "OEM 3": oe_values[2] if len(oe_values) > 2 else "",
        "OEM 4": oe_values[3] if len(oe_values) > 3 else "",
        "OEM 5": oe_values[4] if len(oe_values) > 4 else "",
        "Number Values": "; ".join(number_values),
        "Multiple values": len(interchange_values) > 1 or len(oe_values) > 1,
        "Part Brand": part_brand,
        "Product URL": url,
    }
    if part_number:
        log_action("Part", f"Parsed number values for {part_number}: number_values={number_values}, interchange={interchange_values or ([interchange] if interchange else [])}, oem={oe_values or ([oem] if oem else [])}")
    key_fields = ["Oldest Year", "Brand", "Model", "Interchange Number", "OEM Number"]
    if not any(row[k] for k in key_fields):
        row["Status"] = "parse_empty"
    else:
        row["Status"] = "ok" + ("-fitment-fallback" if fallback and oldest != "" else "")
    return row


# --------------------------------------------------------------------------- #
# BROWSER / CLOUDFLARE
# --------------------------------------------------------------------------- #
def find_chrome(custom=None):
    if custom:
        return custom
    system = platform.system()
    cands = []
    if system == "Windows":
        for env in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
            base = os.environ.get(env)
            if base:
                cands.append(Path(base) / "Google/Chrome/Application/chrome.exe")
    elif system == "Darwin":
        cands.append(Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"))
    else:
        for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser"):
            p = shutil.which(name)
            if p:
                cands.append(Path(p))
    for c in cands:
        if c.exists():
            return str(c)
    return None


def cdp_alive(port):
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=1)
        return True
    except Exception:
        return False


def _kill_chrome_processes():
    """Best-effort cleanup of stale Chrome/Chromium processes left behind by earlier runs."""
    commands = []
    if platform.system() == "Windows":
        commands = [
            ["taskkill", "/F", "/IM", "chrome.exe"],
            ["taskkill", "/F", "/IM", "chromium.exe"],
            ["taskkill", "/F", "/IM", "msedge.exe"],
        ]
    else:
        commands = [
            ["pkill", "-f", "chrome.*--remote-debugging-port=9222"],
            ["pkill", "-f", "chromium.*--remote-debugging-port=9222"],
        ]
    for cmd in commands:
        try:
            subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
        except Exception:
            pass


def reset_chrome(chrome_path=None):
    """Hard reset a stale DevTools Chrome session before continuing."""
    log_action("Chrome", "Resetting stale browser session")
    _kill_chrome_processes()
    time.sleep(2)
    if chrome_path is None:
        chrome_path = find_chrome()
    return start_chrome(chrome_path)


def start_chrome(chrome_path):
    if cdp_alive(CDP_PORT):
        log_action("Chrome", f"Re-using existing session on port {CDP_PORT}")
        return None
    if not chrome_path:
        sys.exit("Could not find Google Chrome. Install it or pass --chrome-path \"C:\\...\\chrome.exe\"")
    log_action("Chrome", f"Launching browser at {chrome_path}")
    try:
        PROFILE_DIR.mkdir(exist_ok=True)
    except PermissionError as e:
        sys.exit(f"Permission denied creating Chrome profile folder: {PROFILE_DIR}. Close any Chrome windows using it and rerun.")
    except OSError as e:
        sys.exit(f"Could not create Chrome profile folder: {PROFILE_DIR} ({e})")
    proc = subprocess.Popen([
        chrome_path,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={PROFILE_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        BASE,
    ])
    for _ in range(40):
        if cdp_alive(CDP_PORT):
            log_action("Chrome", "DevTools connection ready")
            return proc
        time.sleep(0.5)
    proc.terminate()
    sys.exit("Chrome started but the DevTools port never opened. Close all Chrome windows and retry.")


def is_challenge(page):
    try:
        title = (page.title() or "").lower()
        if any(t in title for t in ("just a moment", "attention required", "security check", "verify you are human")):
            return True
        body = page.inner_text("body", timeout=3000).lower()[:1500]
        return any(t in body for t in (
            "verify you are human", "checking your browser", "review the security of your connection",
            "enable javascript and cookies to continue", "performing security verification",
        ))
    except Exception:
        return False


def wait_for_challenge(page, timeout=900):
    if not is_challenge(page):
        return
    log_action("Cloudflare", "Verification page detected; finish the check in the Chrome window")
    start = time.time()
    while time.time() - start < timeout:
        time.sleep(2)
        if not is_challenge(page):
            log_action("Cloudflare", "Cleared; continuing")
            time.sleep(1.5)
            return
    sys.exit("Timed out waiting for the Cloudflare check to be completed.")


def goto(page, url):
    log_action("Navigate", url)
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        log_action("Warning", f"Navigation issue ({type(e).__name__}); checking for challenge")
    wait_for_challenge(page)


def human_pause(lo, hi):
    time.sleep(random.uniform(lo, hi))


def log_action(action, details=""):
    stamp = time.strftime("%H:%M:%S")
    message = f"[{stamp}] {action}"
    if details:
        message += f" - {details}"
    print(message, flush=True)


def open_search_panel(page):
    """CARiD starts with a collapsed header search widget; click it so the real #search-field appears."""
    log_action("Search", "Opening the site search panel")
    selectors = [
        '.header-search-label',
        '.js-search-input-for-preact-render .header-search-label',
        'button[aria-label*="Open search" i]',
        '.search-form .search-label',
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel)
            for idx in range(min(loc.count(), 10)):
                candidate = loc.nth(idx)
                try:
                    if candidate.is_visible():
                        candidate.click(force=True)
                        page.wait_for_timeout(500)
                        log_action("Search", "Search panel opened")
                        return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


# --------------------------------------------------------------------------- #
# SEARCH
# --------------------------------------------------------------------------- #
def apply_search_value(box, value):
    try:
        box.evaluate("(el) => { el.readOnly = false; el.removeAttribute('readonly'); el.disabled = false; el.focus(); }")
    except Exception:
        pass
    try:
        box.fill(value)
    except Exception:
        pass
    try:
        box.evaluate("(el) => { el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); }")
    except Exception:
        pass


def set_search_value(page, box, value):
    if not value:
        return False
    log_action("Search", f"Entering part number: {value}")
    try:
        box.wait_for(state="visible", timeout=20000)
        box.scroll_into_view_if_needed()
        box.click(force=True)
        box.focus()
    except Exception:
        pass

    handled = page.evaluate(
        """
        (partNumber) => {
            const selectors = [
                '#search-field',
                'input[id="search-field"]',
                'input#search-field',
                'input[name="search-field"]',
                'input[type="search"]',
                'input[name="q"]',
                'input[name="search"]',
                'input[id*="search" i]',
            ];
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                if (!el) continue;
                if (el.hidden || el.getAttribute('aria-hidden') === 'true') continue;
                el.scrollIntoView({ block: 'center', inline: 'center' });
                el.focus();
                el.click();
                el.value = '';
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.value = partNumber;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
            }
            return false;
        }
        """,
        value,
    )
    if handled:
        return True

    try:
        box.fill("")
        box.type(value, delay=70)
        return True
    except Exception:
        return False


def run_search(page, pn, args):
    log_action("Search", f"Starting lookup for {pn}")
    if args.search_url:
        goto(page, args.search_url.format(q=quote_plus(pn)))
    else:
        goto(page, BASE)
        open_search_panel(page)

        box = find_search_box(page)
        if box is None:
            raise RuntimeError("search field not found after page load")
        log_action("Search", f"Found search box for {pn}")

        if not set_search_value(page, box, pn):
            raise RuntimeError("search field found but could not receive the part number")

        if not wait_for_search_results(page):
            log_action("Search", "No live result list yet; waiting a bit longer")
            page.wait_for_timeout(1500)
        if click_first_search_result(page, pn):
            log_action("Search", "Clicked first matching product result")
            return

        try:
            box.press("Enter")
        except Exception:
            try:
                page.keyboard.press("Enter")
            except Exception:
                pass
        try:
            page.wait_for_load_state("domcontentloaded", timeout=30000)
        except Exception:
            pass
        wait_for_challenge(page)
    time.sleep(random.uniform(2.0, 3.5))
    wait_for_challenge(page)


def normalize_href(href):
    if not href:
        return ""
    href = str(href).strip().split("#", 1)[0]
    if not href:
        return ""
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        return BASE + href
    return href


def _pn_matches(pn, text="", href=""):
    if pn is None:
        return True
    needle = re.sub(r"[^a-z0-9]", "", str(pn).lower())
    haystack = re.sub(r"[^a-z0-9]", "", (str(text or "") + " " + str(href or "")).lower())
    return needle in haystack


def _preferred_result_links(items, pn=""):
    exact = []
    fallback = []
    for href, text in items:
        href = normalize_href(href)
        if not href or not href.startswith(BASE) or "/search/" in href:
            continue
        if ".html" not in href and "/product/" not in href and "/parts/" not in href and "/accessory/" not in href and "/radiator/" not in href:
            continue
        if not text.strip():
            continue
        if pn and _pn_matches(pn, text, href):
            exact.append(href)
        else:
            fallback.append(href)
    return exact or fallback


def collect_product_links(page, limit, pn=""):
    js = """els => els
        .filter(a => !a.closest('header, nav, footer'))
        .map(a => ({href: a.href,
                    text: (a.textContent || '').trim(),
                    wrap: (a.closest('[class*=prod i], [class*=item i], [class*=result i], [class*=card i]') || {}).className || ''}))"""
    try:
        items = page.eval_on_selector_all("a[href]", js)
    except Exception:
        return []
    links = []
    candidates = []
    for it in items:
        href = normalize_href(it["href"])
        text = it.get("text", "")
        if not href.startswith(BASE):
            continue
        if "/search/" in href:
            continue
        if not (".html" in href or "/product/" in href or "/parts/" in href or "/accessory/" in href or "/radiator/" in href):
            continue
        candidates.append((href, text))
    links = _preferred_result_links(candidates, pn)
    if links:
        return links[:limit]

    # Fallback: if the page is just a dropdown suggestions list, take the first visible product result.
    for sel in ("a[href]", "li a[href]", "[role='option'] a[href]", "div a[href]"):
        try:
            locs = page.locator(sel).all()
        except Exception:
            continue
        candidates = []
        for loc in locs[:30]:
            try:
                href = normalize_href(loc.get_attribute("href") or "")
                text = loc.inner_text()
            except Exception:
                continue
            if not href or not href.startswith(BASE) or "/search/" in href:
                continue
            if ".html" not in href and "/product/" not in href and "/parts/" not in href and "/accessory/" not in href and "/radiator/" not in href:
                continue
            if not text.strip():
                continue
            candidates.append((href.split("#")[0], text))
        links = _preferred_result_links(candidates, pn)
        if links:
            return links[:limit]
    return []


def wait_for_search_results(page, timeout_ms=20000):
    """Wait until the live CARiD result list is visible."""
    deadline = time.time() + (timeout_ms / 1000.0)
    selectors = [
        '.departments-grid',
        '.departments-grid .item-departments',
        '.departments-grid .item-departments a[href]',
        '.item-departments',
        '.item-departments a[href]',
        '[class*="departments-grid" i]',
        '[class*="item-departments" i]',
        '[class*="item-departments" i] a[href]',
    ]
    while time.time() < deadline:
        for sel in selectors:
            try:
                loc = page.locator(sel)
                if loc.count() and loc.first.is_visible():
                    return True
            except Exception:
                continue
        time.sleep(0.25)
    return False


def click_first_search_result(page, pn):
    """Click the first visible product result in the live search results panel."""
    log_action("Search", "Checking result list for a product match")
    if not wait_for_search_results(page):
        log_action("Search", "No visible results found in the live list")
        return False

    item_links = page.locator('.departments-grid .item-departments a[href]')
    try:
        count = item_links.count()
    except Exception:
        count = 0
    if count:
        for i in range(min(count, 20)):
            loc = item_links.nth(i)
            try:
                if not loc.is_visible():
                    continue
                href = normalize_href(loc.get_attribute("href") or "")
                if not href or not href.startswith(BASE):
                    continue
                before_url = page.url
                loc.click(force=True)
                page.wait_for_timeout(1000)
                if page.url == before_url:
                    goto(page, href)
                wait_for_challenge(page)
                return True
            except Exception:
                continue

    item_links = page.locator('.item-departments a[href]')
    try:
        count = item_links.count()
    except Exception:
        count = 0
    if count:
        for i in range(min(count, 20)):
            loc = item_links.nth(i)
            try:
                if not loc.is_visible():
                    continue
                href = normalize_href(loc.get_attribute("href") or "")
                if not href or not href.startswith(BASE):
                    continue
                before_url = page.url
                loc.click(force=True)
                page.wait_for_timeout(1000)
                if page.url == before_url:
                    goto(page, href)
                wait_for_challenge(page)
                return True
            except Exception:
                continue

    selectors = [
        ".departments-grid .item-departments [href]",
        ".item-departments [href]",
        "[role='option'] a[href]",
        "[class*='search-results' i] a[href]",
        "[id*='search-results' i] a[href]",
        "[class*='autocomplete' i] a[href]",
        "[class*='predictive-search' i] a[href]",
        "[role='listbox'] a[href]",
        "li a[href]",
        "a[href]",
    ]
    seen = set()
    for sel in selectors:
        try:
            items = page.locator(sel).all()[:50]
        except Exception:
            continue
        candidates = []
        for loc in items:
            try:
                if not loc.is_visible():
                    continue
                href = normalize_href(loc.get_attribute("href") or "")
                text = (loc.inner_text() or "").strip()
            except Exception:
                continue
            if not href or not href.startswith(BASE):
                continue
            if "/search/" in href:
                continue
            if not (text or href):
                continue
            if ".html" not in href and "/product/" not in href and "/parts/" not in href and "/accessory/" not in href and "/radiator/" not in href:
                continue
            href_norm = href.split("#")[0]
            if href_norm in seen:
                continue
            seen.add(href_norm)
            candidates.append((href_norm, text))
        preferred = _preferred_result_links(candidates, pn)
        if not preferred:
            continue
        href = preferred[0]
        for loc in items:
            try:
                if normalize_href(loc.get_attribute("href") or "").split("#")[0] == href:
                    before_url = page.url
                    loc.click(force=True)
                    page.wait_for_timeout(1000)
                    if page.url == before_url:
                        goto(page, href)
                    wait_for_challenge(page)
                    return True
            except Exception:
                continue
    return False


def find_search_box(page):
    selectors = [
        '#search-field',
        'input#search-field',
        'input[name="search-field"]',
        '.search-form #search-field',
        'input#native-search-input',
        'input.search-input',
        'input[type="search"]',
        'input[name="q"]',
        'input[name="search"]',
        'input[placeholder*="Search by Make Model Year" i]',
        'input[placeholder*="Search" i]',
        'input[aria-label*="Search" i]',
        'input[id*="search" i]',
        'input[type="text"]',
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel)
            for idx in range(min(loc.count(), 20)):
                candidate = loc.nth(idx)
                if not candidate.count():
                    continue
                if candidate.get_attribute("type") == "hidden":
                    continue
                if candidate.get_attribute("aria-hidden") == "true":
                    continue
                try:
                    if candidate.is_visible():
                        return candidate
                except Exception:
                    continue
        except Exception:
            continue

    # The search widget is sometimes collapsed in the header until the label/button is clicked.
    opened = open_search_panel(page)
    if opened:
        for sel in selectors:
            try:
                loc = page.locator(sel)
                for idx in range(min(loc.count(), 20)):
                    candidate = loc.nth(idx)
                    if candidate.count() and candidate.is_visible():
                        return candidate
            except Exception:
                continue
    try:
        box = page.get_by_role("searchbox")
        for idx in range(min(box.count(), 20)):
            candidate = box.nth(idx)
            if candidate.is_visible():
                return candidate
    except Exception:
        pass
    return None


def dump_debug(page, name, tag):
    DEBUG_DIR.mkdir(exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", name)
    try:
        (DEBUG_DIR / f"{safe}_{tag}.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(DEBUG_DIR / f"{safe}_{tag}.png"), full_page=False)
    except Exception as e:
        print(f"[debug] could not save debug files: {e}")


def _page_matches_part_number(html, pn):
    if pn is None:
        return True
    normal = re.sub(r"[^a-z0-9]", "", str(pn).lower())
    if not normal:
        return True
    body = re.sub(r"[^a-z0-9]", "", html.lower())
    return normal in body


def process_part(page, pn, args):
    """Return a list of result rows (dicts) for one part number."""
    log_action("Part", f"Processing {pn}")
    for attempt in range(2):
        run_search(page, pn, args)
        if args.debug:
            dump_debug(page, pn, "search")

        html = page.content()
        if is_product_html(html):
            log_action("Part", f"Detected product page for {pn}")
            if _page_matches_part_number(html, pn):
                urls = [page.url]
                break
            # stale/incorrect product page: reset and retry once with a fresh page
            if attempt == 0:
                log_action("Part", "Product page mismatched; retrying with a fresh search")
                goto(page, BASE)
                continue
            return [{"Partslink Number": pn, "Status": "not_found"}]

        urls = collect_product_links(page, args.max_products, pn)
        if urls and any(_page_matches_part_number(page.content(), pn) for _ in [0]):
            log_action("Part", f"Found {len(urls)} matching product URL(s) for {pn}")
            break
        if not urls:
            log_action("Part", f"No product results found for {pn}")
            if not args.debug:
                dump_debug(page, pn, "search_noresult")
            return [{"Partslink Number": pn, "Status": "not_found"}]
        if attempt == 0:
            log_action("Part", "Search results were stale; refreshing the page")
            goto(page, BASE)
            continue
        break

    rows = []
    for i, url in enumerate(urls):
        if url != page.url:
            log_action("Part", f"Opening product page {i + 1}: {url}")
            human_pause(1.5, 3.0)
            goto(page, url)
            time.sleep(1.5)
        html = page.content()
        if not _page_matches_part_number(html, pn):
            log_action("Part", f"Skipped product page for {pn}; part number not found on page")
            continue
        log_action("Part", f"Parsing product page {i + 1} for {pn}")
        row = parse_product(html, page.url, part_number=pn)
        row["Partslink Number"] = pn
        if args.debug or row["Status"] == "parse_empty":
            dump_debug(page, pn, f"product{i + 1}")
        rows.append(row)
    if not rows:
        return [{"Partslink Number": pn, "Status": "not_found"}]
    return rows


# --------------------------------------------------------------------------- #
# EXCEL I/O
# --------------------------------------------------------------------------- #
def read_part_numbers(path, column):
    df = pd.read_excel(path, dtype=str)
    if column:
        if column not in df.columns:
            sys.exit(f"Column '{column}' not found. Columns in file: {list(df.columns)}")
        series = df[column]
    else:
        pick = next((c for c in df.columns if re.search(r"part\s*-?\s*link|partslink|part\s*(no|num)", str(c), re.I)), None)
        series = df[pick] if pick else df.iloc[:, 0]
        print(f"[input] Using column: {pick or df.columns[0]!r}")
    seen, out = set(), []
    for v in series.dropna():
        v = str(v).strip()
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def save_results(rows, out_path):
    df = pd.DataFrame(rows)
    for c in OUTPUT_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    df = df[OUTPUT_COLUMNS]
    tmp = str(out_path) + ".tmp.xlsx"
    try:
        df.to_excel(tmp, index=False)
        os.replace(tmp, out_path)
    except PermissionError as e:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        raise PermissionError(
            f"Permission denied writing to '{out_path}'. Close Excel or any app that has the file open, then rerun."
        ) from e
    except OSError as e:
        raise OSError(f"Could not write '{out_path}': {e}") from e


def normalize_resume_row(row):
    """Accept older files using 'Partslink' while the code expects 'Partslink Number'."""
    if not isinstance(row, dict):
        return {}
    out = {}
    for key, val in row.items():
        if key is None:
            continue
        name = str(key).strip()
        normalized = {
            "partslink": "Partslink Number",
            "partslink number": "Partslink Number",
            "part number": "Partslink Number",
        }.get(name.lower(), name)
        out[normalized] = val
    if "Partslink Number" not in out:
        for old_name in ("Partslink", "Part Number"):
            if old_name in out:
                out["Partslink Number"] = out[old_name]
                break
    return out


def load_resume_rows(out_path):
    try:
        prev = pd.read_excel(out_path, dtype=str).fillna("")
    except Exception:
        return []
    return [normalize_resume_row(r) for r in prev.to_dict("records")]


# --------------------------------------------------------------------------- #
# MAIN
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="carid.com part lookup -> Excel")
    ap.add_argument("--input", help="Excel file with the part numbers")
    ap.add_argument("--output", default="results.xlsx", help="Excel file to write (default results.xlsx)")
    ap.add_argument("--column", help="Column name that holds the part numbers (auto-detected if omitted)")
    ap.add_argument("--limit", type=int, help="Only process the first N part numbers (for testing)")
    ap.add_argument("--max-products", type=int, default=1, help="Products to read per search (default 1)")
    ap.add_argument("--delay-min", type=float, default=4.0, help="Min seconds between part numbers")
    ap.add_argument("--delay-max", type=float, default=9.0, help="Max seconds between part numbers")
    ap.add_argument("--search-url", help="Optional search URL template, e.g. https://www.carid.com/search/?q={q}")
    ap.add_argument("--chrome-path", help="Path to chrome.exe / Chrome binary if auto-detect fails")
    ap.add_argument("--no-resume", action="store_true", help="Ignore an existing output file and start over")
    ap.add_argument("--debug", action="store_true", help="Save HTML + screenshot of every page into ./debug")
    ap.add_argument("--parse-file", help="Parse a saved product HTML file and print the result (no browser)")
    args = ap.parse_args()

    if args.parse_file:
        html = Path(args.parse_file).read_text(encoding="utf-8", errors="ignore")
        print(json.dumps(parse_product(html, args.parse_file), indent=2, ensure_ascii=False, default=str))
        return
    if not args.input:
        ap.error("--input is required")

    parts = read_part_numbers(args.input, args.column)
    if args.limit:
        parts = parts[: args.limit]
    print(f"[input] {len(parts)} part numbers")

    out_path = Path(args.output)
    rows = []
    if out_path.exists() and not args.no_resume:
        rows = load_resume_rows(out_path)
        done = {
            str(r.get("Partslink Number", "")).strip()
            for r in rows
            if str(r.get("Partslink Number", "")).strip()
            and not str(r.get("Status", "")).startswith("error")
        }
        rows = [r for r in rows if str(r.get("Partslink Number", "")).strip() in done]
        parts = [p for p in parts if p not in done]
        if done:
            print(f"[resume] {len(done)} already done, {len(parts)} left")
        else:
            print("[resume] existing output file does not contain a valid part-number column; starting fresh")
    if not parts:
        print("Nothing to do.")
        return

    chrome_proc = start_chrome(find_chrome(args.chrome_path))
    try:
        with sync_playwright() as pw:
            browser = None
            context = None
            while True:
                try:
                    browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
                    context = browser.contexts[0] if browser.contexts else browser.new_context()
                    page = context.pages[0] if context.pages else context.new_page()
                    page.set_default_timeout(30000)
                    break
                except Exception as e:
                    msg = str(e)
                    if "Target.createTarget" in msg or "Failed to open a new tab" in msg or "Target closed" in msg:
                        print("[browser] Stale Chrome/CDP session detected; restarting Chrome...")
                        chrome_proc = reset_chrome(find_chrome(args.chrome_path))
                        time.sleep(2)
                        continue
                    raise

            goto(page, BASE)
            print("[ready] Browser is on carid.com. Starting...\n")

            for n, pn in enumerate(parts, 1):
                log_action("Loop", f"Starting part {n}/{len(parts)}: {pn}")
                if page is not None:
                    try:
                        page.close()
                    except Exception:
                        pass
                try:
                    page = context.new_page()
                except Exception as e:
                    msg = str(e)
                    if "Target.createTarget" in msg or "Failed to open a new tab" in msg or "Target closed" in msg:
                        print("[browser] Browser context became stale; restarting Chrome before continuing...")
                        chrome_proc = reset_chrome(find_chrome(args.chrome_path))
                        browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
                        context = browser.contexts[0] if browser.contexts else browser.new_context()
                        page = context.new_page()
                    else:
                        raise
                page.set_default_timeout(30000)
                print(f"[{n}/{len(parts)}] {pn} ... ", end="", flush=True)
                try:
                    result = process_part(page, pn, args)
                except KeyboardInterrupt:
                    raise
                except Exception as e:
                    log_action("Error", f"{type(e).__name__}: {e}")
                    result = [{"Partslink Number": pn, "Status": f"error: {type(e).__name__}"}]
                    if args.debug:
                        dump_debug(page, pn, "error")
                else:
                    statuses = ", ".join(sorted({r.get('Status', '') for r in result}))
                    log_action("Result", statuses)
                    print(statuses)
                rows.extend(result)
                save_results(rows, out_path)
                if n < len(parts):
                    human_pause(args.delay_min, args.delay_max)
    except KeyboardInterrupt:
        print("\nStopped by user. Progress is saved; run the same command again to resume.")
    finally:
        save_results(rows, out_path)
        if chrome_proc:
            chrome_proc.terminate()
    print(f"\nDone. Results: {out_path.resolve()}")


if __name__ == "__main__":
    main()
