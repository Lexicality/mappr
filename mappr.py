from __future__ import generator_stop

import typing  # noqa: F401

from openpyxl import load_workbook
from openpyxl.cell import Cell as xlCell
# from openpyxl.styles import Color, colors


def is_void(cell: xlCell) -> bool:
    return not cell.has_style or cell.fill.bgColor == 'FFFFFFFF' or cell.fill.bgColor == '00FFFFFF'


wb = load_workbook('test.xlsx')

sheet = wb.active
cell = sheet['H4']
