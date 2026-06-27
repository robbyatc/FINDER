import io
import unittest

import pandas as pd
from openpyxl import load_workbook

from reconciliation import (
    build_excel_report,
    reconcile_dat_vs_stream,
    validate_required_columns,
)


class ReconciliationTests(unittest.TestCase):
    def setUp(self):
        self.dep = pd.DataFrame(
            {
                "'callsign'": ["'LNI201'", "'LNI201'", "'CTV911'"],
                "'adep'": ["'WIMM'", "'WIMM'", "'WIMM'"],
                "'ades'": ["'WIII'", "'WIII'", "'WSSS'"],
                "'eobd'": ["'260523'", "'260523'", "'260523'"],
                "'atd'": ["'2026-05-23 22:19:00'", "'2026-05-23 22:19:00'", "'23:01'"],
                "'departureRunway'": ["'23'", "'23'", "'23'"],
                "'departureGate'": ["'30'", "'30'", "'27'"],
                "'register'": ["'PKLHO'", "'PKLHO'", "'PKGTG'"],
            }
        )
        self.arr = pd.DataFrame(
            {
                "'callsign'": ["'LNI200'", "'WON1265'"],
                "'adep'": ["'WIII'", "'WIMB'"],
                "'ades'": ["'WIMM'", "'WIMM'"],
                "'eobd'": ["'260523'", "'260523'"],
                "'ata'": ["'2026-05-24 00:09:00'", "'2026-05-24 00:28:00'"],
                "'arrivalRunway'": ["'23'", "'23'"],
                "'arrivalGate'": ["'30'", "'11'"],
                "'register'": ["'PKLHO'", "'PKWHT'"],
            }
        )
        self.stream = pd.DataFrame(
            {
                "DATE OF FLIGHT": ["2026-05-23", "2026-05-23", "2026-05-23"],
                "FLIGHT NUMBER": ["LNI201", "LNI200", "EXTRA1"],
                "AERODROME": ["WIMM", "WIII", "WIMM"],
                "TO FROM": ["WIII", "WIMM", "WIII"],
                "AC REGISTER": ["PKLHO", "PKLHO", "PKXXX"],
                "D/A/L/O": ["D", "A", "D"],
                "RWY": ["23", "23", "05"],
                "PARKING": ["30", "30", "10"],
                "ATD": ["2026-05-23 22:19", "", "09:00"],
                "ATA": ["", "2026-05-24 00:09", ""],
            }
        )

    def test_reconciliation_counts_and_fields(self):
        result = reconcile_dat_vs_stream(self.dep, self.arr, self.stream)
        summary = result["summary"]
        self.assertEqual(summary["total_dat_dep"], 3)
        self.assertEqual(summary["total_dat_arr"], 2)
        self.assertEqual(summary["total_dat_combined"], 5)
        self.assertEqual(summary["duplicate_dat"], 1)
        self.assertEqual(summary["matched"], 2)
        self.assertEqual(summary["missing_in_stream"], 2)
        self.assertEqual(summary["extra_in_stream"], 1)
        self.assertAlmostEqual(summary["accuracy_percentage"], 50.0)

        matched_dep = result["matched"].loc[
            result["matched"]["FLIGHT NUMBER"] == "LNI201"
        ].iloc[0]
        self.assertEqual(matched_dep["DATE OF FLIGHT"], "2026-05-23")
        self.assertEqual(matched_dep["DEPARTURE GATE"], "30")
        self.assertEqual(matched_dep["DEPARTURE RUNWAY"], "23")
        self.assertEqual(matched_dep["ATD"], "22:19")

    def test_excel_report_has_required_sheets(self):
        result = reconcile_dat_vs_stream(self.dep, self.arr, self.stream)
        workbook = load_workbook(io.BytesIO(build_excel_report(result)), read_only=True)
        self.assertEqual(
            workbook.sheetnames,
            [
                "Summary",
                "Missing in Stream",
                "Matched",
                "Extra in Stream",
                "Duplicate DAT",
            ],
        )
        self.assertEqual(workbook["Summary"]["A1"].value, "FINDER — DAT vs STREAM Reconciliation Report")

    def test_validation_reports_missing_stream_movement(self):
        mapping = {"flight": "f", "eobd": "d", "adep": "o", "ades": "x"}
        issues = validate_required_columns(mapping, mapping, mapping)
        self.assertEqual(len(issues), 1)
        self.assertIn("D/A/L/O", issues[0])


if __name__ == "__main__":
    unittest.main()
