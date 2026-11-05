import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from carid_interchange_scraper import completed_interchange_values, read_interchange_values


class CaridInterchangeResumeTests(unittest.TestCase):
    def test_interchange_input_keeps_duplicate_rows(self) -> None:
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            pd.DataFrame({"Interchange Number": ["101-10026", "101-10026"]}).to_excel(
                input_path, index=False
            )

            values = read_interchange_values(input_path, "Interchange Number")

        self.assertEqual(values, ["101-10026", "101-10026"])

    def test_resume_uses_input_interchange_not_returned_partslink(self) -> None:
        saved = [
            {"Interchange Number": "ABC123", "Partslink Number": "XY9876543", "Status": "OK"},
            {"Interchange Number": "NOTFOUND", "Partslink Number": "", "Status": "not_found"},
            {"Interchange Number": "RETRY", "Partslink Number": "", "Status": "error: Timeout"},
        ]

        self.assertEqual(
            completed_interchange_values(saved),
            {"abc123", "notfound"},
        )


if __name__ == "__main__":
    unittest.main()