from win32com.client import constants
from pywintypes import com_error
import time
import traceback, sys
import pythoncom


from exceldriver.tools import _restart_excel_with_addins_and_attach
from .wait import (
    _wait_for_datastream_result,
    WorkbookClosedException,
    DatastreamDataErrorException,
    _get_cell_by_index
)
from exceldriver.exceptions import NoExcelWorkbookException

def populate_datastream_for_file(filepath, excel, retries_remaining=3, close_workbook=False, index=0):
    """

    Private function has main functionality. This is a wrapper to add retries afer com errors
    """

    # Necessary to be called in each new thread or process which uses COM (communicate with Microsoft products)
    pythoncom.CoInitialize()

    # Even if things are going normally, restart every 500 worksheets as there is a memory leak
    if index % 500 == 0 and retries_remaining == 3:
        excel = _restart_excel_with_addins_and_attach(start_sleep=60)

    # Stop retries
    if retries_remaining <= 0:
        print(fr'ERROR: Could not process {filepath}. Skipping and moving to "..\failed".')
        return excel, False

    try:
        # If we are retrying, need to close the workbook before trying to populate
        if close_workbook:
            excel.CutCopyMode = False
            time.sleep(1)
            if excel.ActiveWorkbook:
                excel.ActiveWorkbook.Close(SaveChanges=False)
            time.sleep(5)

        _populate_datastream_for_file(filepath, excel)
        return excel, True
    except (com_error, WorkbookClosedException, DatastreamDataErrorException, NoExcelWorkbookException) as e:
        print(f'Error {e} populating {filepath}. Will wait 30 seconds, restart Excel, and try again.')
        traceback.print_tb(sys.exc_info()[2])
        time.sleep(30)
        excel = _restart_excel_with_addins_and_attach(start_sleep=60)
        return populate_datastream_for_file(filepath, excel, retries_remaining=retries_remaining - 1, close_workbook=True)


def _populate_datastream_for_file(filepath, excel):
    wb = excel.Workbooks.Open(filepath)
    _run_datastream_func(excel)
    successful = _wait_for_datastream_result(excel)
    _copy_paste_values(excel, wb)
    _relabel_date(excel, wb)
    excel.ActiveWorkbook.Close(SaveChanges=True)
    return successful


def _run_datastream_func(excel):
    dstream_func_cell = _get_cell_by_index(excel, 1, 1)
    dstream_func_cell.Activate()
    dstream_func_cell.Calculate()


def _copy_paste_values(excel, wb, range='A1:J10000'):
    ws = wb.Sheets('Sheet')
    ws.Range(range).Copy()
    ws.Range(range.split(':')[0]).PasteSpecial(Paste=constants.xlPasteValues, Operation=constants.xlNone)
    excel.CutCopyMode = False


def _relabel_date(excel, wb, cell_range: str ='A1'):
    """
    Currently date is coming in with odd name, replace with Date
    """
    ws = wb.Sheets('Sheet')
    ws.Range(cell_range).Value = 'Date'