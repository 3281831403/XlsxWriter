###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("optimize16.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test add_table() with totals and formulas in constant_memory mode."""

        workbook = Workbook(self.got_filename, {"constant_memory": True})

        worksheet = workbook.add_worksheet()
        xformat = workbook.add_format({"num_format": 2})

        worksheet.set_column("B:K", 10.288)

        worksheet.add_table(
            "B3:K6",
            {
                "total_row": 1,
                "columns": [
                    {"header": "Column1", "total_string": "Total"},
                    {"header": "Column2"},
                    {"header": "Column3", "total_function": "average"},
                    {"header": "Column4", "total_function": "count"},
                    {"header": "Column5", "total_function": "count_nums"},
                    {"header": "Column6", "total_function": "max"},
                    {"header": "Column7", "total_function": "min"},
                    {"header": "Column8", "total_function": "sum"},
                    {"header": "Column9", "total_function": "stdDev"},
                    {
                        "header": "Column10",
                        "total_function": "var",
                        "formula": "SUM(Table1[[#This Row],[Column1]:[Column3]])",
                        "format": xformat,
                    },
                ],
            },
        )

        # Write data rows, leaving the formula column (K) to the table.
        data = [1, 2, 3, None, None, 6, 7, 8, 9]
        worksheet.write_row("B4", data)
        worksheet.write_row("B5", data)

        workbook.close()

        self.assertExcelEqual()
