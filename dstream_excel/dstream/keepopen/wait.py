import time
from dstream_excel.dstream.workbook.exceptions import DatastreamDataErrorException


def wait_for_datastream_result(ws):
    finished = False
    while not finished:
        time.sleep(1)
        finished = _datastream_is_done(ws)

    return True


def _datastream_is_done(ws):
    if _datastream_not_working(ws):
        raise DatastreamDataErrorException('Datastream plugin did not properly pull values.')

    return _data_in_b2(ws)


def _datastream_not_working(ws):
    a1 = ws.range('A1').value
    error_starts_with = ('ERR','#ERROR','Offline','0')

    errors = []
    for pat in error_starts_with:
        if str(a1).lower().startswith(pat.lower()):
            errors.append(True)

    return any(errors)


def _data_in_b2(ws):
    b2 = ws.range('B2').value
    return b2 is not None
