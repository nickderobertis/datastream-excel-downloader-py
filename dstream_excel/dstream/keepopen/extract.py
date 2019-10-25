from typing import Sequence
import time
import pandas as pd
from xlwings.main import Sheet
from pythoncom import com_error
from dstream_excel.dstream.workbook.exceptions import DatastreamDataErrorException


def firm_dfs_from_multi_firm_df(multi_firm_df: pd.DataFrame, symbols: Sequence[str], num_variables: int):
    dfs = []
    for i, symbol in enumerate(symbols):
        bottom_column_index = i * num_variables
        top_column_index = (i + 1) * num_variables
        firm_df = multi_firm_df.iloc[:, bottom_column_index:top_column_index]
        firm_df['Symbol'] = symbol
        dfs.append(firm_df)
    return dfs


def long_firm_df_from_ds_multi_firm_df(multi_firm_df: pd.DataFrame, symbols: Sequence[str], num_variables: int):
    firm_dfs = firm_dfs_from_multi_firm_df(multi_firm_df, symbols, num_variables)
    return pd.concat(firm_dfs)


def get_multi_firm_df_from_sheet(ws: Sheet, table_start_range: str = 'A1', retries: int = 3):
    while retries > 0:
        try:
            df = ws.range(table_start_range).options(pd.DataFrame, expand='table').value
        except com_error:
            time.sleep(1)
            retries -= 1
            if retries <= 0:
                raise DatastreamDataErrorException('could not populate datastream values.')
            continue
        break

    df.index.name = 'Date'
    return df
