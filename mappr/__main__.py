from openpyxl import load_workbook

from .cell import Cell

wb = load_workbook('test.xlsx')

sheet = wb.active
cell = sheet['H4']
mycell = Cell(cell)
