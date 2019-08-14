from openpyxl import Workbook
import os


def get_workbook_and_worksheet():
    wb = Workbook()
    ws = wb.active
    return wb, ws

