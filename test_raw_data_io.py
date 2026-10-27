import os
import tempfile
import unittest
from unittest import mock

import pandas as pd

from customer_filter import discover_customers
from raw_data_io import (
    find_matching_raw_data_file,
    find_raw_data_file_by_prefix,
    read_raw_data,
)


class RawDataIoTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.frame = pd.DataFrame(
            {
                "point_time": ["2026/10/27 08:00", "2026/10/27 09:00"],
                "point_val": [1.25, 1.5],
                "Customer": ["A", "B"],
            }
        )

    def test_reads_csv_and_xlsx_with_same_columns(self):
        csv_path = os.path.join(self.temp_dir.name, "G_C.csv")
        xlsx_path = os.path.join(self.temp_dir.name, "G_D.xlsx")
        self.frame.to_csv(csv_path, index=False)
        self.frame.to_excel(xlsx_path, index=False)

        pd.testing.assert_frame_equal(read_raw_data(csv_path), self.frame)
        pd.testing.assert_frame_equal(read_raw_data(xlsx_path), self.frame)

    def test_excel_uses_first_worksheet_and_honors_nrows(self):
        path = os.path.join(self.temp_dir.name, "G_C.xlsx")
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            self.frame.to_excel(writer, sheet_name="RawData", index=False)
            pd.DataFrame({"wrong": [1]}).to_excel(writer, sheet_name="Other", index=False)

        result = read_raw_data(path, nrows=1)
        self.assertEqual(list(result.columns), list(self.frame.columns))
        self.assertEqual(len(result), 1)

    def test_finds_supported_excel_extensions(self):
        for filename in ("Group_Chart.xls", "Group_Other.xlsm", "Group_Third.xlsb"):
            open(os.path.join(self.temp_dir.name, filename), "wb").close()

        matched = find_matching_raw_data_file(self.temp_dir.name, "Group", "Chart")
        prefixed = find_raw_data_file_by_prefix(self.temp_dir.name, "Group_Other")
        self.assertTrue(matched.endswith("Group_Chart.xls"))
        self.assertTrue(prefixed.endswith("Group_Other.xlsm"))

    def test_legacy_xls_selects_xlrd_engine(self):
        path = os.path.join(self.temp_dir.name, "Group_Chart.xls")
        with mock.patch("raw_data_io.pd.read_excel", return_value=self.frame) as reader:
            result = read_raw_data(path, nrows=5, float_precision="round_trip")

        self.assertIs(result, self.frame)
        reader.assert_called_once_with(path, nrows=5, sheet_name=0, engine="xlrd")

    def test_customer_discovery_includes_excel_raw_data(self):
        path = os.path.join(self.temp_dir.name, "Group_Chart.xlsx")
        self.frame.to_excel(path, index=False)

        customers, errors = discover_customers(self.temp_dir.name)

        self.assertEqual(customers, ["A", "B"])
        self.assertEqual(errors, [])


if __name__ == "__main__":
    unittest.main()
