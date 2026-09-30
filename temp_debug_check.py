from pathlib import Path

from carid_scraper import parse_product
from carid_interchange_scraper import extract_partslink_number_from_html

for path in sorted(Path('debug').glob('*01034*')):
    html = path.read_text(encoding='utf-8', errors='ignore')
    parts = extract_partslink_number_from_html(html)
    row = parse_product(html, str(path), part_number=parts)
    print('FILE=', path.name)
    print('PARTSLINK=', parts)
    print('OEM Number=', row.get('OEM Number'))
    print('OEM 1=', row.get('OEM 1'))
    print('OEM 2=', row.get('OEM 2'))
    print('OEM 3=', row.get('OEM 3'))
    print('OEM 4=', row.get('OEM 4'))
    print('OEM 5=', row.get('OEM 5'))
    print('---')
