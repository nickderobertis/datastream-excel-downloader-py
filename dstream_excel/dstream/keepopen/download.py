from typing import Sequence
import uuid
import os
import pandas as pd
from xlwings.main import Sheet
from processfiles.objs import ObjectProcessTracker
from dstream_excel.dstream.keepopen.extract import long_firm_df_from_ds_multi_firm_df, get_multi_firm_df_from_sheet
from dstream_excel.dstream.workbook.functions import DatastreamExcelFunction
from dstream_excel.dstream.keepopen.wait import wait_for_datastream_result
from dstream_excel.dstream.keepopen.excel import get_empty_ws


def download_datastream_save_to_csvs(symbols: Sequence[str], variables: Sequence[str], out_folder: str = 'inprogress',
                                     num_symbols_per_call: int = 50, restart: bool = False,
                                     **ds_func_kwargs) -> pd.DataFrame:
    ws = get_empty_ws()
    tracker = ObjectProcessTracker(symbols, restart=restart)
    for symbols_chunk in tracker.obj_generator(chunk=num_symbols_per_call):
        _download_datastream_save_to_csvs(ws, symbols_chunk, variables, out_folder=out_folder, **ds_func_kwargs)


def _download_datastream_save_to_csvs(ws: Sheet, symbols: Sequence[str], variables: Sequence[str],
                                      out_folder: str = 'inprogress',
                                      **ds_func_kwargs) -> pd.DataFrame:
    dsef = DatastreamExcelFunction(symbols)
    func_str = dsef.time_series(variables, **ds_func_kwargs)
    ws.range('A1').value = func_str
    wait_for_datastream_result(ws)
    wide_df = get_multi_firm_df_from_sheet(ws)
    df = long_firm_df_from_ds_multi_firm_df(wide_df, symbols, len(variables))

    # Random output name
    out_name = str(uuid.uuid4()) + '.csv'
    out_path = os.path.join(out_folder, out_name)
    if not os.path.exists(out_folder):
        os.makedirs(out_folder)
    df.to_csv(out_path)

    return df

