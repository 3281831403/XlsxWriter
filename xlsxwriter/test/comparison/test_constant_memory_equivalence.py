###############################################################################
#
# Tests that a worksheet generated in 'constant_memory' mode is equivalent to
# one generated in standard mode when using add_table() and autofit().
#
# The same generation script is run in both modes and the resulting files are
# compared cell by cell, as well as on the table definition (table1.xml) and
# the column widths.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import os
import re
import unittest
import xml.etree.ElementTree as ET
from datetime import datetime
from zipfile import ZipFile

from xlsxwriter.workbook import Workbook

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _shared_strings(zip_file):
    # Get the list of shared strings from a workbook, if present.
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []

    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    strings = []

    for si in root.findall(f"{NS}si"):
        # Concatenate all text nodes to handle rich strings.
        strings.append("".join(node.text or "" for node in si.iter(f"{NS}t")))

    return strings


def _get_cells(filename):
    # Extract a dict of cell data from a worksheet XML file, resolving both
    # shared strings and the in-line strings used in constant_memory mode.
    zip_file = ZipFile(filename)
    strings = _shared_strings(zip_file)
    root = ET.fromstring(zip_file.read("xl/worksheets/sheet1.xml"))

    cells = {}
    for cell in root.iter(f"{NS}c"):
        ref = cell.get("r")
        cell_type = cell.get("t", "n")
        style = cell.get("s", "0")

        formula = cell.find(f"{NS}f")
        formula = formula.text if formula is not None else None

        if cell_type == "s":
            # Shared string.
            value = strings[int(cell.find(f"{NS}v").text)]
        elif cell_type == "inlineStr":
            # In-line string, used in constant_memory mode.
            inline = cell.find(f"{NS}is")
            value = "".join(node.text or "" for node in inline.iter(f"{NS}t"))
        else:
            node = cell.find(f"{NS}v")
            value = node.text if node is not None else None

        cells[ref] = (value, formula, style, cell_type if cell_type == "b" else "")

    zip_file.close()

    return cells


def _get_element(filename, element):
    # Extract an element (such as <cols>) from the worksheet XML as a string.
    zip_file = ZipFile(filename)
    xml_str = zip_file.read("xl/worksheets/sheet1.xml").decode("utf-8")
    zip_file.close()

    match = re.search(rf"<{element}>.*?</{element}>|<{element}/>", xml_str)

    return match.group(0) if match else None


def _get_table_xml(filename):
    zip_file = ZipFile(filename)
    table_xml = zip_file.read("xl/tables/table1.xml")
    zip_file.close()

    return table_xml


def generate_report(filename, constant_memory):
    # A daily-report style export. The same script is run in standard and
    # constant_memory modes. The table is added before any data is written,
    # which is required in constant_memory mode.
    workbook = Workbook(filename, {"constant_memory": constant_memory})
    worksheet = workbook.add_worksheet()

    money_format = workbook.add_format({"num_format": 2})
    date_format = workbook.add_format({"num_format": 14})
    header_format = workbook.add_format({"bold": True, "bg_color": "#D7E4BC"})

    worksheet.add_table(
        0,
        0,
        102,
        5,
        {
            "name": "DailyReport",
            "style": "Table Style Light 11",
            "total_row": 1,
            "columns": [
                {
                    "header": "Item",
                    "header_format": header_format,
                    "total_string": "Total",
                },
                {"header": "Qty", "total_function": "sum"},
                {"header": "Price", "format": money_format},
                {"header": "Amount", "formula": "=[@Qty]*[@Price]"},
                {"header": "Date"},
                {"header": "Flag"},
            ],
        },
    )

    for i in range(1, 102):
        worksheet.write(i, 0, f"item-{i}")
        worksheet.write(i, 1, i)
        worksheet.write(i, 2, i * 1.5, money_format)
        # The Amount column is populated by the table column formula.
        worksheet.write_datetime(i, 4, datetime(2026, 11, 13), date_format)
        worksheet.write_boolean(i, 5, i % 2 == 0)

    worksheet.autofit()

    workbook.close()


def generate_report_with_data_option(filename, constant_memory):
    # A report where the table data is supplied via the add_table() 'data'
    # option, without a header row.
    workbook = Workbook(filename, {"constant_memory": constant_memory})
    worksheet = workbook.add_worksheet()

    data = [[f"item-{i}", i, i * 0.5] for i in range(1, 501)]

    worksheet.add_table(
        2,
        1,
        502,
        3,
        {
            "data": data,
            "header_row": False,
            "style": "Table Style Medium 4",
            "banded_columns": True,
        },
    )

    worksheet.autofit()

    workbook.close()


def generate_report_for_autofit(filename, constant_memory):
    # A report with content that exercises the autofit width calculation for
    # different data types.
    workbook = Workbook(filename, {"constant_memory": constant_memory})
    worksheet = workbook.add_worksheet()

    text_wrap = workbook.add_format({"text_wrap": True})

    worksheet.add_table(
        0,
        0,
        50,
        3,
        {
            "columns": [
                {"header": "Short"},
                {"header": "A much longer header text"},
                {"header": "Multiline"},
                {"header": "Numbers"},
            ]
        },
    )

    for i in range(1, 51):
        worksheet.write(i, 0, "x" * (i % 7))
        worksheet.write(i, 1, "some longer text " + str(i))
        worksheet.write(i, 2, "Foo\nBamboo\nBar", text_wrap)
        worksheet.write(i, 3, i * 12345)

    worksheet.autofit()
    worksheet.autofit(300)

    workbook.close()


class TestConstantMemoryEquivalence(unittest.TestCase):
    """
    Test that files created in constant_memory mode match the equivalent
    standard mode files.
    """

    maxDiff = None

    def setUp(self):
        self.test_dir = "xlsxwriter/test/comparison/output/"
        os.makedirs(self.test_dir, exist_ok=True)

    def tearDown(self):
        for filename in self._filenames():
            if os.path.exists(filename):
                os.remove(filename)

    def _filenames(self):
        return [
            self.test_dir + name + suffix
            for name in ("cm_report", "cm_data_opt", "cm_autofit")
            for suffix in ("_stream.xlsx", "_normal.xlsx")
        ]

    def _assert_equivalent(self, stream_file, normal_file):
        # The table definition must be identical.
        self.assertEqual(_get_table_xml(stream_file), _get_table_xml(normal_file))

        # The column widths must be identical.
        self.assertEqual(
            _get_element(stream_file, "cols"), _get_element(normal_file, "cols")
        )

        # The dimensions must be identical.
        self.assertEqual(
            _get_element(stream_file, "dimension"),
            _get_element(normal_file, "dimension"),
        )

        # Every cell (header, data, formula and total cells) must match.
        stream_cells = _get_cells(stream_file)
        normal_cells = _get_cells(normal_file)

        self.assertEqual(normal_cells, stream_cells)

    def test_report_with_totals_and_formulas(self):
        """Test a report with a total row and a formula column."""
        stream_file = self.test_dir + "cm_report_stream.xlsx"
        normal_file = self.test_dir + "cm_report_normal.xlsx"

        generate_report(stream_file, True)
        generate_report(normal_file, False)

        self._assert_equivalent(stream_file, normal_file)

    def test_report_with_data_option(self):
        """Test a report using the add_table() 'data' option."""
        stream_file = self.test_dir + "cm_data_opt_stream.xlsx"
        normal_file = self.test_dir + "cm_data_opt_normal.xlsx"

        generate_report_with_data_option(stream_file, True)
        generate_report_with_data_option(normal_file, False)

        self._assert_equivalent(stream_file, normal_file)

    def test_report_autofit_widths(self):
        """Test that autofit column widths match in both modes."""
        stream_file = self.test_dir + "cm_autofit_stream.xlsx"
        normal_file = self.test_dir + "cm_autofit_normal.xlsx"

        generate_report_for_autofit(stream_file, True)
        generate_report_for_autofit(normal_file, False)

        self._assert_equivalent(stream_file, normal_file)
