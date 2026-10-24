from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from check_data import compare_totals


class CheckTotals(unittest.TestCase):
    def test_repeated_entries_are_summed_before_comparison(self):
        master = pd.DataFrame({"entry": ["A", "A", "B"], "weight": [2.5, 3.5, 0.0]})
        summary = pd.DataFrame({"entry": ["B", "A"], "total": [0.0, 6.0]})
        result = compare_totals(master, summary, "entry", "weight", "total")
        self.assertTrue(result["matches"].all())

    def test_missing_measurements_are_not_zero(self):
        master = pd.DataFrame({"entry": ["A", "B"], "weight": [float("nan"), 0.0]})
        summary = pd.DataFrame({"entry": ["A", "B"], "total": [0.0, 0.0]})
        result = compare_totals(master, summary, "entry", "weight", "total").set_index("entry")
        self.assertFalse(result.loc["A", "matches"])
        self.assertTrue(result.loc["B", "matches"])

    def test_missing_summary_entry_is_exposed(self):
        master = pd.DataFrame({"entry": ["A", "B"], "weight": [3.0, 5.0]})
        summary = pd.DataFrame({"entry": ["A"], "total": [3.0]})
        result = compare_totals(master, summary, "entry", "weight", "total").set_index("entry")
        self.assertFalse(result.loc["B", "matches"])


if __name__ == "__main__":
    unittest.main()
