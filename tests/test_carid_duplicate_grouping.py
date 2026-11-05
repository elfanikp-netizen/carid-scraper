import unittest

from carid_scraper import (
    expand_results_by_inputs,
    group_query_values,
    result_groups_from_rows,
)


class CaridDuplicateGroupingTests(unittest.TestCase):
    def test_partslink_result_set_is_replayed_for_each_input_occurrence(self):
        parts = ["AB1234567", "AB1234567", "CD1234567"]
        groups = group_query_values(parts)
        results = result_groups_from_rows([
            {"Partslink Number": "AB1234567", "Status": "OK", "Product URL": "first"},
            {"Partslink Number": "AB1234567", "Status": "OK", "Product URL": "second"},
        ], "Partslink Number")

        self.assertEqual(groups["AB1234567"], [0, 1])
        expanded = expand_results_by_inputs(parts, results)
        self.assertEqual([row["Product URL"] for row in expanded], ["first", "second", "first", "second"])

    def test_interchange_groups_by_selected_query_and_retries_errors(self):
        values = ["101-10026", "101-10026"]
        groups = group_query_values(values)
        saved = result_groups_from_rows([
            {"Interchange Number": "101-10026", "Status": "error: Timeout"},
        ], "Interchange Number")

        self.assertEqual(groups["10110026"], [0, 1])
        self.assertEqual(saved, {})


if __name__ == "__main__":
    unittest.main()