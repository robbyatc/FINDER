import io
import unittest
import zipfile

import pandas as pd

from comparison import (
    actual_file_format,
    compare_datasets,
    detect_mapping,
    display_value,
    excel_sheet_names,
    normalize_value,
    read_actual_excel,
    read_source_csv,
    results_to_zip,
)


class ComparisonTests(unittest.TestCase):
    def setUp(self):
        self.source = pd.DataFrame(
            {
                "Flight Number": ["GA 100", "ID200", "ID200", "QZ300"],
                "From": ["CGK", "CGK", "CGK", "DPS"],
                "To": ["DPS", "SUB", "SUB", "CGK"],
                "Date of Flight": ["27/06/2026"] * 4,
                "EOBT": ["08:30", "09:00", "09:00", "10:00"],
                "ATD": ["08:41", "09:12", "09:15", "10:10"],
                "Register": ["PK-GAA", "PK-IDA", "PK-IDB", "PK-QZA"],
            }
        )
        self.actual = pd.DataFrame(
            {
                "Callsign": ["GA100", "ID 200", "ID200", "JT400"],
                "ADEP": ["WIII", "CGK", "CGK", "SUB"],
                "ADES": ["DPS", "SUB", "SUB", "CGK"],
                "EOBD": ["2026-06-27"] * 4,
                "EOBT": [830, 900, 900, 1100],
                "ATD": ["08:43", "09:12", "09:15", "11:05"],
                "Registration": ["PKGAA", "PK-IDA", "PK-IDB", "PK-JTA"],
            }
        )
        self.source_map = detect_mapping(self.source.columns)
        self.actual_map = detect_mapping(self.actual.columns)

    def test_alias_detection(self):
        self.assertEqual(self.source_map["flight"], "Flight Number")
        self.assertEqual(self.actual_map["flight"], "Callsign")
        self.assertEqual(self.source_map["adep"], "From")
        self.assertEqual(self.actual_map["ades"], "ADES")

    def test_normalization(self):
        self.assertEqual(normalize_value("GA 100", "flight"), "GA100")
        self.assertEqual(normalize_value("27/06/2026", "eobd"), "2026-06-27")
        self.assertEqual(normalize_value("27062026", "eobd"), "2026-06-27")
        self.assertEqual(normalize_value("20260627", "eobd"), "2026-06-27")
        self.assertEqual(normalize_value(0.5, "eobt"), "12:00")
        self.assertEqual(normalize_value(830, "eobt"), "08:30")
        self.assertEqual(normalize_value("2026-06-27 08:30", "atd"), "08:30")
        self.assertEqual(normalize_value("-", "ata"), "")
        self.assertEqual(display_value("2026-06-27 -", "ata"), "")

    def test_html_report_disguised_as_xls(self):
        html = b"""
        <html><body>
        <table><tr><td>Report title</td></tr></table>
        <table><thead><tr>
          <th>DATE OF FLIGHT</th><th>FLIGHT NUMBER</th><th>AERODROME</th>
          <th>TO FROM</th><th>AC REGISTER</th><th>RWY</th><th>ATD</th>
        </tr></thead><tbody>
          <tr><td>2026-06-27</td><td>CTV910</td><td>WIII</td>
          <td>WIMM</td><td>PKGTG</td><td>23</td><td>2026-06-27 08:30</td></tr>
        </tbody></table>
        </body></html>
        """
        self.assertEqual(actual_file_format(html), "html")
        self.assertEqual(excel_sheet_names(html), ["Tabel Data"])
        frame = read_actual_excel(html, "Tabel Data")
        self.assertEqual(frame.shape, (1, 7))
        self.assertEqual(frame.loc[0, "FLIGHT NUMBER"], "CTV910")
        mapping = detect_mapping(frame.columns)
        self.assertEqual(mapping["register"], "AC REGISTER")
        self.assertEqual(mapping["runway"], "RWY")

    def test_comparison_with_duplicates_and_differences(self):
        result = compare_datasets(
            self.source,
            self.actual,
            self.source_map,
            self.actual_map,
            ["flight", "eobd", "ades"],
        )
        summary = result["summary"]
        self.assertEqual(summary["matched_total"], 3)
        self.assertEqual(summary["source_only_total"], 1)
        self.assertEqual(summary["actual_only_total"], 1)
        self.assertEqual(summary["differences_total"], 1)
        self.assertIn("ATD", result["differences"].iloc[0]["Field Berbeda"])

    def test_incomplete_keys_do_not_match_by_default(self):
        source = pd.DataFrame({"Flight Number": [""], "Date of Flight": ["27/06/2026"]})
        actual = pd.DataFrame({"Callsign": [""], "EOBD": ["2026-06-27"]})
        source_map = detect_mapping(source.columns)
        actual_map = detect_mapping(actual.columns)
        result = compare_datasets(
            source, actual, source_map, actual_map, ["flight", "eobd"]
        )
        self.assertEqual(result["summary"]["matched_total"], 0)
        self.assertEqual(result["summary"]["source_only_total"], 1)
        self.assertEqual(result["summary"]["actual_only_total"], 1)

    def test_csv_separator_and_zip_export(self):
        frame, encoding = read_source_csv(b"Flight Number;From;To\nGA100;CGK;DPS\n")
        self.assertEqual(encoding, "utf-8-sig")
        self.assertEqual(frame.shape, (1, 3))

        result = compare_datasets(
            self.source,
            self.actual,
            self.source_map,
            self.actual_map,
            ["flight", "eobd", "ades"],
        )
        with zipfile.ZipFile(io.BytesIO(results_to_zip(result))) as archive:
            self.assertEqual(len(archive.namelist()), 5)
            self.assertIn("03_semua_tidak_cocok.csv", archive.namelist())


if __name__ == "__main__":
    unittest.main()
