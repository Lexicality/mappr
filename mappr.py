from __future__ import generator_stop

import typing  # noqa: F401

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.cell import Cell as xlCell
from openpyxl.worksheet import Worksheet as xlSheet
# from openpyxl.styles import Color, colors


def _valid_coord(sheet: xlSheet, x: str, y: int):
    col = column_index_from_string(x)
    return (1 <= col and col <= sheet.max_column) and (1 <= y and y <= sheet.max_row)


class Cell:
    def __init__(self, cell: xlCell):
        self.cell = cell

    def is_void(self) -> bool:
        return (
            not self.cell.has_style or
            self.cell.fill.bgColor == 'FFFFFFFF' or
            self.cell.fill.bgColor == '00FFFFFF'
        )


wb = load_workbook('test.xlsx')

sheet = wb.active
cell = sheet['H4']
