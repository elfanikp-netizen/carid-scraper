import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import MagicMock, patch

import pandas as pd

import carid_scraper


class CaridBrowserIsolationTests(unittest.TestCase):
    def test_partslink_input_keeps_duplicate_rows(self) -> None:
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            pd.DataFrame({"Partslink Number": ["GM1240366V", "GM1240366V"]}).to_excel(
                input_path, index=False
            )

            values = carid_scraper.read_part_numbers(input_path, "Partslink Number")

        self.assertEqual(values, ["GM1240366V", "GM1240366V"])

    def test_existing_product_page_search_box_is_reused(self) -> None:
        page = MagicMock()
        page.url = "https://www.carid.com/some-product.html"
        search_input = MagicMock()
        search_input.is_visible.return_value = True
        page.locator.return_value.count.return_value = 1
        page.locator.return_value.nth.return_value = search_input

        self.assertTrue(carid_scraper.open_search_panel(page))

        self.assertEqual(page.locator.call_args.args[0], "#search-field")

    def test_visible_search_box_does_not_toggle_panel(self) -> None:
        page = MagicMock()
        page.url = "https://www.carid.com/"
        search_input = MagicMock()
        search_input.is_visible.return_value = True
        page.locator.return_value.count.return_value = 1
        page.locator.return_value.nth.return_value = search_input

        self.assertTrue(carid_scraper.open_search_panel(page))

        self.assertEqual(page.locator.call_args.args[0], "#search-field")

    @patch("carid_scraper.click_first_search_result", return_value=True)
    @patch("carid_scraper.wait_for_search_results", return_value=True)
    @patch("carid_scraper.set_search_value", return_value=True)
    @patch("carid_scraper.find_search_box", return_value=MagicMock())
    @patch("carid_scraper.open_search_panel")
    @patch("carid_scraper.goto")
    def test_next_lookup_from_product_page_does_not_navigate_home(
        self, goto, open_panel, _find_box, _set_value, _wait_results, _click_result
    ) -> None:
        page = MagicMock()
        page.url = "https://www.carid.com/some-product.html"
        args = MagicMock(search_url=None)

        carid_scraper.run_search(page, "PN1234567", args)

        goto.assert_not_called()
        open_panel.assert_called_once_with(page)

    def test_runner_port_overrides_standalone_port_selection(self) -> None:
        with patch.dict(os.environ, {"CARID_CDP_PORT": "35671"}):
            self.assertEqual(carid_scraper.get_cdp_port(), 35671)

    def test_standalone_port_is_ephemeral_and_stable_for_the_process(self) -> None:
        with patch.dict(os.environ, {}, clear=True), patch.object(carid_scraper, "_LOCAL_CDP_PORT", None):
            first = carid_scraper.get_cdp_port()
            second = carid_scraper.get_cdp_port()

        self.assertGreater(first, 0)
        self.assertEqual(first, second)

    @patch("carid_scraper.platform.system", return_value="Windows")
    @patch("carid_scraper.subprocess.run")
    def test_cleanup_targets_only_owned_chrome_process_tree(self, run_process, _system) -> None:
        process = MagicMock()
        process.pid = 43210
        process.poll.return_value = None

        carid_scraper.stop_chrome(process)

        run_process.assert_called_once()
        self.assertEqual(
            run_process.call_args.args[0],
            ["taskkill", "/F", "/T", "/PID", "43210"],
        )
        process.wait.assert_called_once_with(timeout=5)


if __name__ == "__main__":
    unittest.main()