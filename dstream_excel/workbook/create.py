from openpyxl import Workbook
import os

from morningstar import holdings_command


def create_xlsx_with_holding_command(folder, *holding_args, **holding_kwargs):
    wb, ws = _get_workbook_and_worksheet()
    _fill_with_holding_command(ws, *holding_args, **holding_kwargs)

    filepath = os.path.join(folder, get_workbook_name_from_holding_args(holding_args))
    wb.save(filepath)

    return os.path.abspath(filepath)

def get_workbook_name_from_holding_args(holding_args):
    return f'{holding_args[0].upper()} {holding_args[1].upper()} holdings.xlsx'

##### Helper functions ####

def _fill_with_holding_command(ws, *holding_args, **holding_kwargs):
    ws['A1'] = holdings_command(*holding_args, **holding_kwargs)

def _get_workbook_and_worksheet():
    wb = Workbook()
    ws = wb.active
    return wb, ws

