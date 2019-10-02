import time
from xlwings.main import Sheet
from dstream_excel.dstream.workbook.exceptions import (
    DatastreamDataErrorException,
    DatastreamFunctionShouldBeReRunException
)


def wait_for_datastream_result(ws: Sheet):
    finished = False
    while not finished:
        time.sleep(1)
        finished = _datastream_is_done(ws)

    return True


def _datastream_is_done(ws: Sheet):
    if _datastream_function_needs_to_be_re_run(ws):
        raise DatastreamFunctionShouldBeReRunException
    if _datastream_not_working(ws):
        raise DatastreamDataErrorException('Datastream plugin did not properly pull values.')

    return _data_in_b2(ws)


def _datastream_not_working(ws: Sheet):
    a1 = ws.range('A1').value
    error_starts_with = ('ERR','#ERROR','Offline','0')

    errors = []
    for pat in error_starts_with:
        if str(a1).lower().startswith(pat.lower()):
            errors.append(True)

    return any(errors)


def _datastream_function_needs_to_be_re_run(ws: Sheet):
    a1 = ws.range('A1').value
    if a1 == -2146826259:  # for some reason this value is coming when cell shows #NAME
        return True
    elif '#NAME' in a1:
        return True

    return False


def _data_in_b2(ws: Sheet):
    b2 = ws.range('B2').value
    return b2 is not None
