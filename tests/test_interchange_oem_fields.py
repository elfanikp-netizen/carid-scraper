import unittest
from types import SimpleNamespace

import carid_interchange_scraper as module


class FakePage:
    url = "https://example.com/product"

    def content(self):
        return "<html></html>"


class TestInterchangeOEMFields(unittest.TestCase):
    def test_process_interchange_part_does_not_keep_partslink_in_oem_fields(self):
        module.run_search = lambda page, value, args: None
        module.collect_product_links = lambda page, max_products, q: []
        module.goto = lambda page, url: None
        module.human_pause = lambda *args, **kwargs: None
        module.is_product_html = lambda html: True
        module.parse_product = lambda html, url, part_number="": {
            "Status": "ok",
            "OEM Number": "AB-C123",
            "OEM 1": "AB C 123",
            "OEM 2": "ABC999",
            "OEM 3": "AB/C123",
            "OEM 4": "",
            "OEM 5": "",
            "Partslink Number": part_number,
        }
        module.extract_partslink_number_from_html = lambda html: "ABC123"

        result = module.process_interchange_part(
            FakePage(),
            "123-4567",
            SimpleNamespace(debug=False, max_products=1),
        )

        self.assertEqual(result[0]["OEM Number"], "")
        self.assertEqual(result[0]["OEM 1"], "")
        self.assertEqual(result[0]["OEM 2"], "ABC999")
        self.assertEqual(result[0]["OEM 3"], "")

    def test_strip_partslink_from_oem_fields_handles_interchange_style_numbers(self):
        row = {
            "OEM Number": "11-01034A",
            "OEM 1": "11 01034A",
            "OEM 2": "OTHER-999",
            "OEM 3": "1101034A",
            "OEM 4": "",
            "OEM 5": "",
        }

        cleaned = module.strip_partslink_from_oem_fields(row, "1101034A")

        self.assertEqual(cleaned["OEM Number"], "")
        self.assertEqual(cleaned["OEM 1"], "")
        self.assertEqual(cleaned["OEM 2"], "OTHER-999")
        self.assertEqual(cleaned["OEM 3"], "")


if __name__ == "__main__":
    unittest.main()
