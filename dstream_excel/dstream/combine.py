import argparse
import os
import re

import pandas as pd

from dstream_excel.tools.ext_pandas import _get_outpath_and_df_of_headers, _append_df_to_csv
from dstream_excel.tracker.files import FileProcessTracker
from dstream_excel.dstream.workbook.filename import _original_symbol_from_filename_symbol


def combine_all_datastream_xlsx(infolder: str, outpath: str = 'all_data.csv', restart: bool = True):
    """
    Combine populated data in XLSX files into single csv file.

    :param infolder: Folder containing populated XLSX files
    :param outpath: Full file path to output csv to
    :param restart: Whether to force a restart of combining files. If the process is stopped, it will be continued
        where it left off if restart=False, and start from the beginning if restart=True.
    :return: None
    """

    file_tracker = FileProcessTracker(folder=infolder, restart=restart, file_types=('xlsx',))
    outpath, df_of_headers = _get_outpath_and_df_of_headers(outpath)
    all_columns = [col for col in df_of_headers.columns]

    for file in file_tracker.file_generator():
        df_of_headers, all_columns = _append_datastream_xlsx_to_csv(file, outpath, df_of_headers, all_columns)

def _append_datastream_xlsx_to_csv(file, outpath, df_of_headers, all_columns):
    df_for_append = pd.read_excel(file)  # load new data
    _reformat_datastream_df_for_append(df_for_append, file)
    df_of_headers, all_columns = _append_df_to_csv(df_for_append, df_of_headers, outpath, all_columns)

    return df_of_headers, all_columns

def _reformat_datastream_df_for_append(df, file):
    """
    Note: inplace
    """
    _clean_datastream_df(df)
    df['Ticker'] = _datastream_filepath_to_iq_id(file)

    return df


def _clean_datastream_df(df):
    """
    Note: inplace
    """
    _drop_unneded_data(df)
    _rename_datastream_cols(df)


def _drop_unneded_data(df):
    """
    Note: inplace
    """
    error_cols = [col for col in df.columns if '#ERROR' in col]
    df.drop(error_cols, axis=1, inplace=True)
    data_cols = [col for col in df.columns if '-' in col]
    df.dropna(subset=data_cols, inplace=True)

def _rename_datastream_cols(df):
    """
    Note: inplace
    """
    rename_dict = _rename_dict_for_datastream_cols(df)
    df.rename(columns=rename_dict, inplace=True)


def _rename_dict_for_datastream_cols(df):
    rename_dict = {
        'Name': 'Date'
    }
    data_cols = [col for col in df.columns if '-' in col]
    data_rename = {col: _new_name_for_datastream_column(col) for col in data_cols}
    rename_dict.update(data_rename)
    return rename_dict


def _new_name_for_datastream_column(col):
    return col.split('-')[-1].strip()


def _datastream_filepath_to_iq_id(filepath):
    filename = os.path.basename(filepath) #strips folders, etc.
    pattern = re.compile(r'([\S\s]*)([.]xlsx)')
    ds_filename_id = pattern.match(filename).group(1)
    ds_id = _original_symbol_from_filename_symbol(ds_filename_id)
    return ds_id

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Populate data in XLSX files containing datastream functions')
    parser.add_argument('-f', '--folder', required=False,
                        default=r'C:\Users\derobertisna.UFAD\Dropbox (Personal)\UF\Andy\ETF Project\Data\Datastream\inprogress')
    parser.add_argument('-o', '--outpath', required=False,
                        default=r'C:\Users\derobertisna.UFAD\Dropbox (Personal)\UF\Andy\ETF Project\Data\Datastream\all datastream data.csv')

    # Default is to restart creation of data
    parser.add_argument('-n', '--no-restart', action='store_true', default=False)

    args = parser.parse_args()
    restart = (not args.no_restart)

    combine_all_datastream_xlsx(args.folder, args.outpath, restart=restart)