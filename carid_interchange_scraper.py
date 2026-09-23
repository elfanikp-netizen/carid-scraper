#!/usr/bin/env python3
"""
carid_interchange_scraper.py

Same browser flow as carid_scraper.py, but it treats the input column as an
interchange/cross-reference value instead of a Partslink value.

With this version, the script searches the interchange value, then writes the
matched product data into the same output columns used by carid_scraper.py.

Usage:
  python carid_interchange_scraper.py --input input.xlsx --output results.xlsx
  python carid_interchange_scraper.py --input input.xlsx --column Interchange --limit 3 --debug
"""

import argparse
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup

from carid_scraper import (
    OUTPUT_COLUMNS,
    collect_product_links,
    dump_debug,
    find_chrome,
    goto,
    human_pause,
    is_product_html,
    load_resume_rows,
    log_action,
    parse_product,
    run_search,
    save_results,
    start_chrome,
    sync_playwright,
)


def read_interchange_values(path, column):
    df = pd.read_excel(path, dtype=str)
    if column:
        if column not in df.columns:
            sys.exit(f"Column '{column}' not found. Columns in file: {list(df.columns)}")
        series = df[column]
    else:
        pick = next(
            (
                c
                for c in df.columns
                if re.search(r"interchange|cross[- ]?ref|cross[- ]?reference|crossref", str(c), re.I)
            ),
            None,
        )
        if pick is None:
            pick = next(
                (
                    c
                    for c in df.columns
                    if re.search(r"oem|oe|oe number|oem number|part.*link|partslink", str(c), re.I)
                ),
                None,
            )
        series = df[pick] if pick else df.iloc[:, 0]
        print(f"[input] Using column: {pick or df.columns[0]!r}")

    seen, out = set(), []
    for v in series.dropna():
        v = str(v).strip()
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def extract_partslink_number_from_html(html):
    """Return a product partslink in the format: 9 chars total, first 2 alphabetic, must contain a digit."""
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text(" ", strip=True)
    for match in re.finditer(r"\b([A-Za-z]{2}[A-Za-z0-9]{7})\b", text):
        candidate = match.group(1).upper()
        if len(candidate) != 9:
            continue
        if not candidate[:2].isalpha():
            continue
        if not any(ch.isdigit() for ch in candidate):
            continue
        return candidate

    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = tag.string or tag.get_text() or ""
            if not data:
                continue
            for pattern in (r'"mpn"\s*:\s*"([^"\\]+)"', r'"sku"\s*:\s*"([^"\\]+)"'):
                m = re.search(pattern, data, re.I)
                if m:
                    candidate = re.sub(r"[^A-Za-z0-9]", "", m.group(1)).upper()
                    if len(candidate) == 9 and candidate[:2].isalpha():
                        return candidate
        except Exception:
            continue

    return ""


def strip_partslink_from_oem_fields(row, partslink_value):
    if not partslink_value:
        return row
    partslink_upper = str(partslink_value).upper()
    for key in ["OEM Number", "OEM 1", "OEM 2", "OEM 3", "OEM 4", "OEM 5"]:
        if key not in row:
            continue
        values = []
        for v in str(row.get(key, "")).split("; "):
            if not v:
                continue
            if v.upper() == partslink_upper:
                continue
            values.append(v)
        row[key] = "; ".join(values)
    return row


def process_interchange_part(page, value, args):
    """Process an interchange/search value without requiring the product page to contain that exact value."""
    log_action("Part", f"Processing interchange {value}")
    print(f"[live] search value: {value}")
    for attempt in range(2):
        run_search(page, value, args)
        if args.debug:
            dump_debug(page, value, "search")

        html = page.content()
        if is_product_html(html):
            log_action("Part", f"Detected product page for interchange {value}")
            urls = [page.url]
            break

        urls = collect_product_links(page, args.max_products, "")
        if urls:
            log_action("Part", f"Found {len(urls)} candidate product URL(s) for interchange {value}")
            break
        if not urls:
            log_action("Part", f"No product results found for {value}")
            if not args.debug:
                dump_debug(page, value, "search_noresult")
            return [{"Partslink Number": "", "Status": "not_found"}]
        if attempt == 0:
            log_action("Part", "Search results were stale; refreshing the page")
            goto(page, "https://www.carid.com")
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
        if not is_product_html(html):
            log_action("Part", f"Skipped product page for {value}; page was not recognized as a product page")
            continue
        log_action("Part", f"Parsing product page {i + 1} for {value}")
        print(f"[live] parsing product page {i + 1}: {page.url}")
        row = parse_product(html, page.url)
        partslink_value = extract_partslink_number_from_html(html)
        print(f"[live] extracted partslink: {partslink_value or '<none>'}")
        if partslink_value:
            row["Partslink Number"] = partslink_value
        else:
            row["Partslink Number"] = row.get("Partslink Number", "")
        if partslink_value:
            cleaned_values = []
            for v in str(row.get("Number Values", "")).split("; "):
                if not v:
                    continue
                if v.upper() == partslink_value:
                    continue
                cleaned_values.append(v)
            row["Number Values"] = "; ".join(cleaned_values)
            row = strip_partslink_from_oem_fields(row, partslink_value)
        row["Interchange Number"] = row.get("Interchange Number") or value
        print(f"[live] parsed row: {row}")
        if args.debug or row["Status"] == "parse_empty":
            dump_debug(page, value, f"product{i + 1}")
        rows.append(row)
    if not rows:
        return [{"Partslink Number": "", "Status": "not_found"}]
    return rows


def is_cdp_stale_error(message):
    msg = str(message).lower()
    return any(token in msg for token in (
        "econnreset",
        "connection reset",
        "target.createtarget",
        "failed to open a new tab",
        "target closed",
        "websocket",
        "unexpected eof",
        "connect_over_cdp",
    ))


def main():
    ap = argparse.ArgumentParser(description="carid.com interchange lookup -> Excel")
    ap.add_argument("--input", help="Excel file with the interchange values")
    ap.add_argument("--output", default="results.xlsx", help="Excel file to write (default results.xlsx)")
    ap.add_argument("--column", help="Column name that holds the interchange values (auto-detected if omitted)")
    ap.add_argument("--limit", type=int, help="Only process the first N interchange values (for testing)")
    ap.add_argument("--max-products", type=int, default=1, help="Products to read per search (default 1)")
    ap.add_argument("--delay-min", type=float, default=4.0, help="Min seconds between searches")
    ap.add_argument("--delay-max", type=float, default=9.0, help="Max seconds between searches")
    ap.add_argument("--search-url", help="Optional search URL template, e.g. https://www.carid.com/search/?q={q}")
    ap.add_argument("--chrome-path", help="Path to chrome.exe / Chrome binary if auto-detect fails")
    ap.add_argument("--no-resume", action="store_true", help="Ignore an existing output file and start over")
    ap.add_argument("--debug", action="store_true", help="Save HTML + screenshot of every page into ./debug")
    ap.add_argument("--parse-file", help="Parse a saved product HTML file and print the result (no browser)")
    args = ap.parse_args()

    if args.parse_file:
        from carid_scraper import parse_product

        html = Path(args.parse_file).read_text(encoding="utf-8", errors="ignore")
        print(parse_product(html, args.parse_file))
        return

    if not args.input:
        ap.error("--input is required")

    values = read_interchange_values(args.input, args.column)
    if args.limit:
        values = values[: args.limit]
    print(f"[input] {len(values)} interchange values")

    out_path = Path(args.output)
    rows = []
    if out_path.exists() and not args.no_resume:
        rows = load_resume_rows(out_path)
        done = {
            str(r.get("Partslink Number", "")).strip()
            for r in rows
            if str(r.get("Partslink Number", "")).strip() and not str(r.get("Status", "")).startswith("error")
        }
        rows = [r for r in rows if str(r.get("Partslink Number", "")).strip() in done]
        values = [v for v in values if v not in done]
        if done:
            print(f"[resume] {len(done)} already done, {len(values)} left")
        else:
            print("[resume] existing output file does not contain a valid part-number column; starting fresh")
    if not values:
        print("Nothing to do.")
        return

    chrome_proc = start_chrome(find_chrome(args.chrome_path))
    try:
        with sync_playwright() as pw:
            browser = None
            context = None
            while True:
                try:
                    browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
                    context = browser.contexts[0] if browser.contexts else browser.new_context()
                    page = context.pages[0] if context.pages else context.new_page()
                    page.set_default_timeout(30000)
                    break
                except Exception as e:
                    msg = str(e)
                    if is_cdp_stale_error(msg):
                        print("[browser] Stale Chrome/CDP session detected; restarting Chrome...")
                        try:
                            if browser is not None:
                                browser.close()
                        except Exception:
                            pass
                        try:
                            if chrome_proc is not None:
                                chrome_proc.terminate()
                        except Exception:
                            pass
                        chrome_proc = start_chrome(find_chrome(args.chrome_path))
                        time.sleep(2)
                        continue
                    raise

            goto(page, "https://www.carid.com")
            print("[ready] Browser is on carid.com. Starting...\n")

            for n, value in enumerate(values, 1):
                log_action("Loop", f"Starting interchange {n}/{len(values)}: {value}")
                if page is not None:
                    try:
                        page.close()
                    except Exception:
                        pass
                try:
                    page = context.new_page()
                except Exception as e:
                    msg = str(e)
                    if is_cdp_stale_error(msg):
                        print("[browser] Browser context became stale; restarting Chrome before continuing...")
                        try:
                            if browser is not None:
                                browser.close()
                        except Exception:
                            pass
                        try:
                            if chrome_proc is not None:
                                chrome_proc.terminate()
                        except Exception:
                            pass
                        chrome_proc = start_chrome(find_chrome(args.chrome_path))
                        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
                        context = browser.contexts[0] if browser.contexts else browser.new_context()
                        page = context.new_page()
                    else:
                        raise
                page.set_default_timeout(30000)
                print(f"[{n}/{len(values)}] {value} ... ", end="", flush=True)
                try:
                    result = process_interchange_part(page, value, args)
                except KeyboardInterrupt:
                    raise
                except Exception as e:
                    log_action("Error", f"{type(e).__name__}: {e}")
                    result = [{"Partslink Number": value, "Status": f"error: {type(e).__name__}"}]
                    if args.debug:
                        dump_debug(page, value, "error")
                else:
                    statuses = ", ".join(sorted({r.get("Status", "") for r in result}))
                    log_action("Result", statuses)
                    print(statuses)

                rows.extend(result)
                save_results(rows, out_path)
                if n < len(values):
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
