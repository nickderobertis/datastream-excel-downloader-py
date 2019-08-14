from typing import Sequence
import argparse
import os

import pandas as pd

from exceldriver.workbook.create import get_workbook_and_worksheet
from .functions import DatastreamExcelFunction
from .filename import _valid_filename_from_symbol


def create_all_xlsx_with_datastream_command(folder: str, symbol_list: Sequence[str], variables: Sequence[str],
                                            **dstream_kwargs):
    """
    Creates XLSX files containing Datastream functions

    The XLSX files will not be populated with data at this point, only the function itself. The files must then
    be populated, then combined.

    :param folder: Folder in which to create Excel files
    :param symbol_list: Datastream symbols representing companies. Available by selecting a list through filters
        in Eikon, then getting an output of that list.
    :param variables: Time-series variables to download
    :param dstream_kwargs: can pass begin, end, or freq. All of these arguments take day, month, and year as
        D, M, and Y. freq can be set to D, M, or Y. begin and end can be set to just Y but combined with
        a time modifier, e.g. begin='-5Y'. Don't pass end to have the data go up to current.
    :return: None
    """
    [create_xlsx_with_datastream_command(folder, symbol, variables, **dstream_kwargs) for symbol in symbol_list]

def create_xlsx_with_datastream_command(folder, symbol, variables, **dstream_kwargs):
    wb, ws = get_workbook_and_worksheet()
    _fill_with_datastream_command(ws, symbol, variables, **dstream_kwargs)

    if not os.path.exists(folder): os.makedirs(folder)

    filepath = os.path.join(folder, f'{_valid_filename_from_symbol(symbol)}.xlsx')
    wb.save(filepath)

    return os.path.abspath(filepath)

def _load_symbols(filepath):
    df = pd.read_csv(filepath, usecols=['Symbol'])
    return df['Symbol'].unique().tolist()

##### Helper functions ####

def _fill_with_datastream_command(ws, symbols, variables, **dstream_kwargs):
    ws['A1'] = DatastreamExcelFunction(symbols).time_series(variables, **dstream_kwargs)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Create XLSX files with datastream functions ready to be opened for pulling data')
    parser.add_argument('-f','--folder', required=False, default=r'C:\Users\derobertisna.UFAD\Dropbox (Personal)\UF\Andy\ETF Project\Data\Datastream\inprogress')
    parser.add_argument('-s','--symbols-path', required=False, default=r'C:\Users\derobertisna.UFAD\Dropbox (Personal)\UF\Andy\ETF Project\Data\Datastream\japan equities list.csv')
    parser.add_argument('-v', '--variables', required=False, default='UP;RI;NOSH;VO')
    parser.add_argument('-b', '--begin', required=False, default='-20Y')
    parser.add_argument('-e', '--end', required=False, default='')

    args = parser.parse_args()

    symbols = _load_symbols(args.symbols_path)
    variables = args.variables.split(';')

    create_all_xlsx_with_datastream_command(
        args.folder,
        symbols,
        variables,
        begin=args.begin,
        end=args.end
    )