# Copyright (c) 2010-2019 openpyxl

import pytest

# compatibility imports


# package imports
from openpyxl.styles import numbers
from openpyxl.reader.excel import load_workbook


@pytest.mark.parametrize("cell, number_format",
                    [
                        ('A1', numbers.FORMAT_GENERAL),
                        ('A2', numbers.FORMAT_DATE_XLSX14),
                        ('A3', numbers.FORMAT_NUMBER_00),
                        ('A4', numbers.FORMAT_DATE_TIME3),
                        ('A5', numbers.FORMAT_PERCENTAGE_00),
                    ]
                    )
def test_read_general_style(datadir, cell, number_format):
    datadir.join("genuine").chdir()
    wb = load_workbook('empty-with-styles.xlsx')
    ws = wb["Sheet1"]
    assert ws[cell].number_format == number_format


def test_read_no_theme(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook('libreoffice_nrt.xlsx')
    assert wb
