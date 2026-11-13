###############################################################################
#
# Tests for add_table() and autofit() in 'constant_memory' mode.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import os
import unittest

from xlsxwriter.exceptions import OverlappingRange, RowAlreadyWrittenError
from xlsxwriter.workbook import Workbook


class TestConstantMemoryTable(unittest.TestCase):
    """
    Test add_table() in constant_memory mode.
    """

    def setUp(self):
        self.test_dir = "xlsxwriter/test/comparison/output/"
        os.makedirs(self.test_dir, exist_ok=True)
        self.filename = self.test_dir + "py_constant_memory_table.xlsx"

    def tearDown(self):
        if os.path.exists(self.filename):
            os.remove(self.filename)

    def _get_workbook(self):
        workbook = Workbook(self.filename, {"constant_memory": True})
        return workbook, workbook.add_worksheet()

    def test_add_table_returns_zero(self):
        """Test that add_table() works in constant_memory mode."""
        workbook, worksheet = self._get_workbook()

        self.assertEqual(worksheet.add_table(0, 0, 100, 3), 0)

        workbook.close()

    def test_add_table_after_flushed_row_raises(self):
        """Test that add_table() raises if the first row is already written."""
        workbook, worksheet = self._get_workbook()

        for row in range(10):
            worksheet.write_row(row, 0, [row])

        # The table's first row has already been flushed to the temp file.
        with self.assertRaises(RowAlreadyWrittenError):
            worksheet.add_table(5, 0, 100, 2)

        with self.assertRaises(RowAlreadyWrittenError):
            worksheet.add_table(0, 0, 100, 2)

        workbook.close()

    def test_add_table_on_current_row_is_allowed(self):
        """Test that add_table() works on the current, unflushed, row."""
        workbook, worksheet = self._get_workbook()

        for row in range(10):
            worksheet.write_row(row, 0, [row])

        # The current row hasn't been flushed yet so a table can start there.
        self.assertEqual(worksheet.add_table(10, 0, 100, 2), 0)
        self.assertEqual(worksheet.add_table(10, 4, 100, 6), 0)

        workbook.close()

    def test_add_table_overlap_table(self):
        """Test overlapping table ranges in constant_memory mode."""
        workbook, worksheet = self._get_workbook()

        worksheet.add_table(0, 0, 50, 2)

        with self.assertRaises(OverlappingRange):
            worksheet.add_table(40, 0, 90, 2)

        # A non-overlapping table is allowed.
        self.assertEqual(worksheet.add_table(0, 4, 50, 6), 0)

        workbook.close()

    def test_add_table_overlap_merge(self):
        """Test a table overlapping a merge range in constant_memory mode."""
        workbook, worksheet = self._get_workbook()

        worksheet.merge_range(60, 0, 70, 1, "merged")

        with self.assertRaises(OverlappingRange):
            worksheet.add_table(65, 0, 80, 1)

        workbook.close()

    def test_merge_range_overlap_table(self):
        """Test a merge range overlapping a table in constant_memory mode."""
        workbook, worksheet = self._get_workbook()

        worksheet.add_table(0, 0, 50, 2)

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range(45, 1, 55, 2, "overlap")

        workbook.close()

    def test_merge_range_overlap_merge(self):
        """Test overlapping merge ranges in constant_memory mode."""
        workbook, worksheet = self._get_workbook()

        worksheet.merge_range(60, 0, 70, 1, "merged")

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range(62, 0, 65, 1, "overlap")

        workbook.close()

    def test_memory_usage_is_bounded(self):
        """Test that per-cell structures don't grow in constant_memory mode."""
        workbook, worksheet = self._get_workbook()

        worksheet.add_table(
            0,
            0,
            10001,
            3,
            {
                "total_row": 1,
                "columns": [
                    {"header": "A", "total_string": "Total"},
                    {"header": "B", "total_function": "sum"},
                    {"header": "C"},
                    {"header": "D", "formula": "=[@B]*2"},
                ],
            },
        )

        for row in range(1, 10001):
            worksheet.write_row(row, 0, [f"item-{row}", row, row * 0.5])

        # The per-cell overlap maps used in standard mode aren't populated.
        self.assertEqual(worksheet.table_cells, {})
        self.assertEqual(worksheet.merged_cells, {})

        # Only the current row of cell data is kept in memory.
        self.assertLessEqual(len(worksheet.table), 1)

        # The table metadata is bounded by the number of tables/columns and
        # doesn't grow with the number of rows.
        self.assertEqual(len(worksheet.tables), 1)
        self.assertEqual(len(worksheet.table_ranges), 1)
        self.assertEqual(len(worksheet.pending_table_totals), 1)
        self.assertEqual(len(worksheet.pending_table_formulas), 1)
        self.assertLessEqual(len(worksheet.autofit_col_widths), 4)

        workbook.close()


class TestConstantMemoryAutofit(unittest.TestCase):
    """
    Test autofit() in constant_memory mode.
    """

    def setUp(self):
        self.test_dir = "xlsxwriter/test/comparison/output/"
        os.makedirs(self.test_dir, exist_ok=True)
        self.filename = self.test_dir + "py_constant_memory_autofit.xlsx"

    def tearDown(self):
        if os.path.exists(self.filename):
            os.remove(self.filename)

    def test_autofit(self):
        """Test that autofit() sets column widths in constant_memory mode."""
        workbook = Workbook(self.filename, {"constant_memory": True})
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, "a")
        worksheet.write_string(1, 0, "aaa")
        worksheet.write_string(2, 0, "aaaaa")

        worksheet.autofit()

        # The width is set from the longest string (7 pixels per 'a'
        # character), plus 7 pixels of padding.
        self.assertEqual(worksheet.col_info[0].width, 5 * 7 + 7)
        self.assertTrue(worksheet.col_info[0].autofit)

        workbook.close()

    def test_autofit_no_data(self):
        """Test that autofit() on an empty worksheet doesn't raise."""
        workbook = Workbook(self.filename, {"constant_memory": True})
        worksheet = workbook.add_worksheet()

        worksheet.autofit()

        self.assertEqual(worksheet.col_info, {})

        workbook.close()

    def test_autofit_unflushed_row(self):
        """Test that autofit() accounts for the current unflushed row."""
        workbook = Workbook(self.filename, {"constant_memory": True})
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, "a")
        worksheet.write_string(1, 0, "aaa")
        # The widest string is in the current, unflushed, row.
        worksheet.write_string(2, 0, "aaaaaaaaaa")

        worksheet.autofit()

        self.assertEqual(worksheet.col_info[0].width, 10 * 7 + 7)

        workbook.close()
