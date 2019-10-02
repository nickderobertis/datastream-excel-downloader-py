import time
from pythoncom import com_error
import xlwings as xw


def get_empty_ws():
    _get_app_starting_excel_if_necessary()
    wb = _get_workbook_opening_if_necessary()
    return wb.sheets('Sheet1')


def _get_app_starting_excel_if_necessary():
    if len(xw.apps) == 0:
        xw.apps.add()
        time.sleep(5)  # it takes time for plugins to be fully active
    return list(xw.apps)[0]


def _get_workbook_opening_if_necessary():
    try:
        return xw.books['Book1']
    except com_error:
        old_names = [bk.name for bk in xw.books]
        xw.books.add()
        time.sleep(5)  # it takes time for plugins to be fully active
        new_name = [bk.name for bk in xw.books if bk.name not in old_names][0]
        return xw.books[new_name]
