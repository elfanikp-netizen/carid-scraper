# carid_scraper

Reads part numbers (Partslink) from an Excel file, searches each one on carid.com,
and writes an Excel file with:

Oldest Year, Newest Year, Brand, Model, Type, Interchange Number, OEM Number
(plus Part Brand, Product URL and Status columns).

## How Cloudflare is handled

The script opens a normal, visible Google Chrome window and controls it through
Chrome's DevTools port. If Cloudflare shows "Verify you are human", **you click it
once** in that window; the script detects that the page cleared and continues.
The Chrome profile is stored in `./chrome_profile`, so the clearance is usually
remembered on later runs. The script does not auto-solve or spoof the check.

## 1. Requirements

- Python 3.9+  (https://www.python.org/downloads/ - tick "Add Python to PATH" on Windows)
- Google Chrome installed

## 2. Install (once)

```
cd carid_scraper
python -m venv venv
venv\Scripts\activate          (Windows)
source venv/bin/activate       (Mac / Linux)
pip install -r requirements.txt
```

No `playwright install` step is needed because your own Chrome is used.

## 3. Prepare the Excel input

One column with the part numbers, header in the first row, e.g.

| Partslink |
|-----------|
| FO1000123 |
| GM1200456 |

The column is auto-detected (header containing "partslink", "part link" or "part number");
otherwise the first column is used. Use `--column "Your Header"` to force one.
Numbers are read as text, so leading zeros are kept.

## 4. Test on 3 parts first

```
python carid_scraper.py --input input.xlsx --output output.xlsx --limit 3 --debug --no-resume
```

1. A Chrome window opens on carid.com.
2. If a Cloudflare check appears, complete it in that window. Wait; the terminal says "Cleared".
3. The script types each number in the site search box and reads the first product.
4. Open `test.xlsx` and check the values. `--debug` also saves the HTML and a screenshot of
   every page into `./debug` so you can see exactly what the script saw.

## 5. Full run

```
python carid_scraper.py --input input.xlsx --output output.xlsx 
```

- Progress is saved after every part number. Stop with Ctrl+C any time; run the same
  command again and it resumes where it stopped (rows with `error` are retried).
- Use `--no-resume` to start over.
- Keep the Chrome window open and do not click around in it while it runs.
- Default pause is 4-9 s between parts. Please keep it polite (`--delay-min 6 --delay-max 12`
  for big lists). Check CARiD's Terms of Use before scraping at volume.

## Options

| Option | Meaning |
|---|---|
| `--column NAME` | header of the part-number column |
| `--max-products N` | read up to N products per search (one output row each), default 1 |
| `--limit N` | only first N part numbers |
| `--search-url URL` | skip the search box, e.g. `https://www.carid.com/search/?q={q}` (only if you know the real pattern) |
| `--chrome-path PATH` | if Chrome is not auto-detected |
| `--debug` | save HTML + screenshot of each page to `./debug` |
| `--parse-file FILE` | parse a saved product page offline and print the result |

## Status column

- `ok` - fields found
- `ok-fitment-fallback` - fitment was read from the whole page, not a fitment block; double-check it
- `parse_empty` - product page opened but nothing matched (HTML saved to `./debug`)
- `not_found` - search returned no product
- `error: ...` - something failed; it is retried on the next run

## If some columns come out empty

The parser was written without being able to open a CARiD product page, so the labels and
layout are best guesses. To fix it:

1. Run with `--debug` on 2-3 numbers.
2. Open `debug/<number>_product1.html` in a browser / editor and find where Interchange,
   OEM and fitment appear.
3. Either send me that HTML file (or the exact text around those fields) and I will adjust
   the parser, or edit the settings at the top of `carid_scraper.py`
   (`SEARCH_INPUT_SELECTORS`, `FITMENT_CONTAINER_SELECTORS`, the label names in `parse_product`).
4. Test without the browser: `python carid_scraper.py --parse-file debug/<number>_product1.html`

## Troubleshooting

- **"Could not find Google Chrome"** - add `--chrome-path "C:\Program Files\Google\Chrome\Application\chrome.exe"`.
- **DevTools port never opened** - close every Chrome window (check Task Manager), retry.
- **Cloudflare loops or keeps re-asking** - finish the check manually in the window, browse to a
  couple of pages by hand, then restart the script (the profile keeps the cookie). Slower delays help.
- **"search box not found"** - the site layout changed or a popup is covering the page; close the
  popup in Chrome, or update `SEARCH_INPUT_SELECTORS`.
- **Brand column** - filled with the vehicle make(s) from the fitment list (Ford, Toyota ...).
  The part manufacturer is in `Part Brand`.
