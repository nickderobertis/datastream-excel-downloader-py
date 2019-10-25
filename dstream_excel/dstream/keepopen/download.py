from typing import Sequence
import uuid
import os
import time
import pandas as pd
from xlwings.main import Sheet
from processfiles.objs import ObjectProcessTracker
from dstream_excel.dstream.keepopen.extract import long_firm_df_from_ds_multi_firm_df, get_multi_firm_df_from_sheet
from dstream_excel.dstream.workbook.functions import DatastreamExcelFunction
from dstream_excel.dstream.keepopen.wait import wait_for_datastream_result
from dstream_excel.dstream.keepopen.excel import get_empty_ws
from dstream_excel.dstream.workbook.exceptions import (
    DatastreamFunctionShouldBeReRunException,
    DatastreamDataErrorException
)
from exceldriver.columns import get_n_cols_after_col
from pythoncom import com_error


def download_datastream_save_to_csvs(symbols: Sequence[str], variables: Sequence[str], out_folder: str = 'inprogress',
                                     num_symbols_per_call: int = 50, restart: bool = False,
                                     **ds_func_kwargs):
    """
    Opens a workbook if necessary, then repeatedly switches out Datastream function, saves the values in
    a csv, then goes to next Datastream function.

    :param symbols: Datastream symbols representing companies. Available by selecting a list through filters
        in Eikon, then getting an output of that list.
    :param variables: Time-series variables to download
    :param out_folder: Folder in which in process XLSX files will be generated
    :param num_symbols_per_call: Number of firms to gather in one pull. Setting this higher should speed up data
        collection, but may also make it less stable. If many variables and a longer history is being pulled,
        a smaller number may be necessary.
    :param restart: Whether to force a restart of download. If the process is stopped, it will be continued
        where it left off if restart=False, and start from the beginning if restart=True.
    :param ds_func_kwargs: can pass begin, end, or freq. All of these arguments take day, month, and year as
        D, M, and Y. freq can be set to D, M, or Y. begin and end can be set to just Y but combined with
        a time modifier, e.g. begin='-5Y'. Don't pass end to have the data go up to current.
    :return: None
    """
    ws = get_empty_ws()
    tracker = ObjectProcessTracker(symbols, restart=restart)
    for symbols_chunk in tracker.obj_generator(chunk=num_symbols_per_call):
        _download_datastream_save_to_csvs(ws, symbols_chunk, variables, out_folder=out_folder, **ds_func_kwargs)


def _download_datastream_save_to_csvs(ws: Sheet, symbols: Sequence[str], variables: Sequence[str],
                                      out_folder: str = 'inprogress',
                                      **ds_func_kwargs) -> pd.DataFrame:
    _write_ds_func_wait_for_result(
        ws,
        symbols,
        variables,
        **ds_func_kwargs
    )
    wide_df = get_multi_firm_df_from_sheet(ws)
    df = long_firm_df_from_ds_multi_firm_df(wide_df, symbols, len(variables))

    # Random output name
    out_name = str(uuid.uuid4()) + '.csv'
    out_path = os.path.join(out_folder, out_name)
    if not os.path.exists(out_folder):
        os.makedirs(out_folder)
    df.to_csv(out_path)

    # Wipe output to prepare for next pull
    end_data_col = get_n_cols_after_col('A', len(variables) * len(symbols))
    ws.range(f'A:{end_data_col}').value = ''

    return df


def _write_ds_func_wait_for_result(ws: Sheet, symbols: Sequence[str], variables: Sequence[str], retries: int = 3,
                                   **ds_func_kwargs):
    dsef = DatastreamExcelFunction(symbols)
    func_str = dsef.time_series(variables, **ds_func_kwargs)
    while retries > 0:
        try:
            ws.range('A1').value = func_str
            wait_for_datastream_result(ws)
        except (DatastreamFunctionShouldBeReRunException, com_error):
            time.sleep(0.5)
            retries -= 1
            if retries <= 0:
                raise DatastreamDataErrorException('could not populate datastream values.')
            continue
        break



