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
        self.set_filename("optimize17.xlsx")

    def test_create_file(self):
        """Test autofit() in constant_memory mode."""

        workbook = Workbook(self.got_filename, {"constant_memory": True})

        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, "Hello")
        worksheet.write(0, 1, "World")
        worksheet.write(0, 2, 123)
        worksheet.write(0, 3, 1234567)
        worksheet.write(1, 0, "Foo bar")
        worksheet.write(2, 0, "Foo bar bar")

        worksheet.autofit()

        workbook.close()

        self.assertExcelEqual()
