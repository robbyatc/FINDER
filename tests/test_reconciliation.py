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
                "Need Review",
                "Extra in Stream",
                "Duplicate DAT",
                "Audit Detail",
            ],
        )
        self.assertEqual(workbook["Summary"]["A1"].value, "FINDER — DAT vs STREAM Reconciliation Report")
        missing_headers = [
            cell.value for cell in workbook["Missing in Stream"][4]
        ]
        self.assertEqual(
            missing_headers,
            [
                "DATE OF FLIGHT",
                "FLIGHT NUMBER",
                "AERODROME",
                "TO FROM",
                "AC REGISTER",
                "ATD",
                "ATA",
                "D/A/L/O",
                "ARRIVAL GATE",
                "DEPARTURE GATE",
                "DEPARTURE RUNWAY",
                "ARRIVAL RUNWAY",
                "MATCH_REASON",
                "STREAM STATUS FLIGHT",
                "DAT MOVEMENT DATETIME",
                "STREAM MOVEMENT DATETIME",
                "TIME DIFFERENCE MINUTES",
            ],
        )

    def test_validation_reports_missing_stream_movement(self):
        mapping = {"flight": "f", "eobd": "d", "adep": "o", "ades": "x"}
        issues = validate_required_columns(mapping, mapping, mapping)
        self.assertEqual(len(issues), 1)
        self.assertIn("D/A/L/O", issues[0])

    def test_sia990_selects_complete_arrival_record(self):
        dep = pd.DataFrame(columns=["callsign", "adep", "ades", "eobd"])
        arr = pd.DataFrame(
            {
                "callsign": ["SIA990", "SIA990"],
                "adep": ["WSSS", "WSSS"],
                "ades": ["WIMM", "WIMM"],
                "eobd": ["260529", "260529"],
                "atd": ["", "23:44"],
                "ata": ["", "00:48"],
                "register": ["", "9VMBT"],
                "arrivalGate": ["", "13"],
                "arrivalRunway": ["", "23"],
                "timeStamp": ["2026-05-29 23:00", "2026-05-30 01:00"],
                "messageNum": ["1", "2"],
            }
        )
        stream = pd.DataFrame(
            columns=[
                "DATE OF FLIGHT",
                "FLIGHT NUMBER",
                "AERODROME",
                "TO FROM",
                "D/A/L/O",
            ]
        )

        result = reconcile_dat_vs_stream(dep, arr, stream)
        selected = result["missing"].iloc[0]
        self.assertEqual(selected["ATD"], "23:44")
        self.assertEqual(selected["ATA"], "00:48")
        self.assertEqual(selected["AC REGISTER"], "9VMBT")
        self.assertEqual(selected["ARRIVAL GATE"], "13")
        self.assertEqual(selected["ARRIVAL RUNWAY"], "23")
        self.assertEqual(selected["DAT MOVEMENT DATETIME"], "2026-05-30 00:48")
        self.assertEqual(len(result["duplicates"]), 1)
        duplicate = result["duplicates"].iloc[0]
        self.assertFalse(bool(duplicate["SELECTED_RECORD_FLAG"]))
        self.assertEqual(int(duplicate["COMPLETENESS_SCORE"]), 0)
        self.assertFalse(bool(duplicate["HAS_MOVEMENT_TIME"]))

    def test_2ambo_complete_record_is_missing_for_other_and_wrong_date(self):
        dep = pd.DataFrame(columns=["callsign", "adep", "ades", "eobd"])
        arr = pd.DataFrame(
            {
                "callsign": ["2AMBO", "2AMBO"],
                "adep": ["ZGGG", "ZGGG"],
                "ades": ["WIMM", "WIMM"],
                "eobd": ["260620", "260620"],
                "atd": ["", "2026-06-20 10:36"],
                "ata": ["", "2026-06-20 14:08"],
                "register": ["", "2AMBO"],
                "arrivalGate": ["", "01"],
                "arrivalRunway": ["", "23"],
                "timeStamp": ["2026-06-20 11:00", "2026-06-20 15:00"],
                "messageNum": ["10", "11"],
            }
        )
        stream = pd.DataFrame(
            {
                "DATE OF FLIGHT": ["2026-06-20"],
                "FLIGHT NUMBER": ["2AMBO"],
                "AERODROME": ["ZGGG"],
                "TO FROM": ["WIMM"],
                "D/A/L/O": ["A"],
                "ATA": ["2026-06-19 14:08"],
                "STATUS FLIGHT": ["OTHER"],
            }
        )

        result = reconcile_dat_vs_stream(dep, arr, stream)
        self.assertEqual(result["summary"]["matched"], 0)
        self.assertEqual(result["summary"]["missing_in_stream"], 1)
        missing = result["missing"].iloc[0]
        self.assertEqual(missing["ATD"], "10:36")
        self.assertEqual(missing["ATA"], "14:08")
        self.assertEqual(missing["AC REGISTER"], "2AMBO")
        self.assertEqual(missing["ARRIVAL GATE"], "01")
        self.assertEqual(missing["ARRIVAL RUNWAY"], "23")
        self.assertIn("STREAM STATUS OTHER", missing["MATCH_REASON"])
        self.assertIn("STREAM INVALID ATA DATE", missing["MATCH_REASON"])
        self.assertEqual(len(result["duplicates"]), 1)

    def test_general_duplicate_prefers_movement_time_then_completeness(self):
        dep = pd.DataFrame(
            {
                "callsign": ["ABC123", "ABC123", "ABC123"],
                "adep": ["WIMM"] * 3,
                "ades": ["WIII"] * 3,
                "eobd": ["260620"] * 3,
                "atd": ["", "08:00", "08:00"],
                "register": ["PKAAA", "", "PKBEST"],
                "departureGate": ["10", "", "12"],
                "departureRunway": ["23", "", "23"],
                "timeStamp": [
                    "2026-06-20 09:00",
                    "2026-06-20 08:00",
                    "2026-06-20 07:00",
                ],
                "messageNum": ["30", "20", "10"],
            }
        )
        arr = pd.DataFrame(columns=["callsign", "adep", "ades", "eobd"])
        stream = pd.DataFrame(
            columns=[
                "DATE OF FLIGHT",
                "FLIGHT NUMBER",
                "AERODROME",
                "TO FROM",
                "D/A/L/O",
            ]
        )
        result = reconcile_dat_vs_stream(dep, arr, stream)
        selected = result["dat_unique"].iloc[0]
        self.assertEqual(selected["AC REGISTER"], "PKBEST")
        self.assertEqual(selected["DEPARTURE GATE"], "12")
        self.assertEqual(len(result["duplicates"]), 2)

    def test_time_tolerance_and_date_validation(self):
        dep = pd.DataFrame(
            {
                "callsign": ["TOL30", "TOL31", "TOL121", "DATEBAD"],
                "adep": ["WIMM"] * 4,
                "ades": ["WIII"] * 4,
                "eobd": ["260620"] * 4,
                "atd": ["10:00", "11:00", "12:00", "13:00"],
            }
        )
        arr = pd.DataFrame(columns=["callsign", "adep", "ades", "eobd"])
        stream = pd.DataFrame(
            {
                "DATE OF FLIGHT": ["2026-06-20"] * 4,
                "FLIGHT NUMBER": ["TOL30", "TOL31", "TOL121", "DATEBAD"],
                "AERODROME": ["WIMM"] * 4,
                "TO FROM": ["WIII"] * 4,
                "D/A/L/O": ["D"] * 4,
                "ATD": [
                    "2026-06-20 10:30",
                    "2026-06-20 11:31",
                    "2026-06-20 14:01",
                    "2026-06-19 13:00",
                ],
                "STATUS FLIGHT": ["ACTIVE"] * 4,
            }
        )
        result = reconcile_dat_vs_stream(
            dep, arr, stream, time_tolerance_minutes=30
        )
        self.assertEqual(set(result["matched"]["FLIGHT NUMBER"]), {"TOL30"})
        self.assertEqual(set(result["need_review"]["FLIGHT NUMBER"]), {"TOL31"})
        self.assertEqual(
            set(result["missing"]["FLIGHT NUMBER"]), {"TOL121", "DATEBAD"}
        )
        reasons = result["missing"].set_index("FLIGHT NUMBER")["MATCH_REASON"]
        self.assertEqual(reasons["TOL121"], "STREAM TIME MISMATCH")
        self.assertIn("STREAM INVALID ATD DATE", reasons["DATEBAD"])

    def test_non_billable_flight_is_separated_from_billing_review(self):
        dep = pd.DataFrame(
            {
                "callsign": ["TEST1", "LNI777"],
                "adep": ["WIMM", "WIMM"],
                "ades": ["WIII", "WIII"],
                "eobd": ["260620", "260620"],
                "atd": ["08:00", "09:00"],
            }
        )
        arr = pd.DataFrame(columns=["callsign", "adep", "ades", "eobd"])
        stream = pd.DataFrame(
            columns=[
                "DATE OF FLIGHT",
                "FLIGHT NUMBER",
                "AERODROME",
                "TO FROM",
                "D/A/L/O",
            ]
        )
        result = reconcile_dat_vs_stream(dep, arr, stream)
        self.assertEqual(set(result["missing_billing"]["FLIGHT NUMBER"]), {"LNI777"})
        self.assertEqual(
            set(result["missing_non_billable"]["FLIGHT NUMBER"]), {"TEST1"}
        )


if __name__ == "__main__":
    unittest.main()
