import time

from dstream_excel.dstream.workbook.exceptions import WorkbookClosedException
from .exceptions import DatastreamDataErrorException

def _wait_for_datastream_result(excel):
    finished = False
    while not finished:
        time.sleep(1)
        finished = _datastream_is_done(excel)

    return True

def _datastream_is_done(excel):
    if _datastream_not_working(excel):
        raise DatastreamDataErrorException('Datastream plugin did not properly pull values.')

    return _data_in_b2(excel)

def _datastream_not_working(excel):
    a1 = _get_cell_value_by_index(excel, 1, 1)
    error_starts_with = ('ERR','#ERROR','Offline','0')

    errors = []
    for pat in error_starts_with:
        if str(a1).lower().startswith(pat):
            errors.append(True)

    return any(errors)


def _data_in_b2(excel):
    b2 = _get_cell_value_by_index(excel, 2, 2)
    return b2 is not None

def _get_cell_value_by_index(excel, x, y):
    try:
        result = excel.ActiveSheet.Cells(x, y).Value
        return result
    except AttributeError:
        raise WorkbookClosedException('Workbook was not open when trying to populate values.')


def _get_cell_by_index(excel, x, y):
    try:
        result = excel.ActiveSheet.Cells(x, y)
        return result
    except AttributeError:
        raise WorkbookClosedException('Workbook was not open when trying to populate values.')