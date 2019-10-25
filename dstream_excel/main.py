from typing import Sequence, Optional
from dstream_excel.dstream.keepopen.download import download_datastream_save_to_csvs
from dstream_excel.dstream.keepopen.combine import combine_all_datastream_csv


def download_datastream(work_folder: str, symbol_list: Sequence[str], variables: Sequence[str],
                        num_symbols_per_call: Optional[int] = 20, outpath: str = 'data.csv',
                        restart: bool = False, **dstream_kwargs):
    """
    Function for full workflow to download data from Datastream.

    Opens a workbook if necessary, then repeatedly switches out Datastream function, saves the values in
    a csv, then goes to next Datastream function. Finally, combines the created csv files into a single csv.

    This function is ultimately just a wrapper which calls:
    :py:func:`.download_datastream_save_to_csvs` to download create many CSV files containing all the data
    :py:func:`.combine_all_datastream_csv` to combine the many CSV files created in the process into a single CSV

    :param work_folder: Folder in which in process CSV files will be generated
    :param symbol_list: Datastream symbols representing companies. Available by selecting a list through filters
        in Eikon, then getting an output of that list.
    :param variables: Time-series variables to download
    :param num_symbols_per_call: Number of firms to gather in one pull. Setting this higher should speed up data
        collection, but may also make it less stable. If many variables and a longer history is being pulled,
        a smaller number may be necessary.
    :param outpath: Full file path to output csv to
    :param restart: Whether to force a restart of download. If the process is stopped, it will be continued
        where it left off if restart=False, and start from the beginning if restart=True.
    :param dstream_kwargs: can pass begin, end, or freq. All of these arguments take day, month, and year as
        D, M, and Y. freq can be set to D, M, or Y. begin and end can be set to just Y but combined with
        a time modifier, e.g. begin='-5Y'. Don't pass end to have the data go up to current.
    :return: None
    """
    print(f'Populating XLSX file, saving to CSV files in {work_folder}')
    download_datastream_save_to_csvs(
        symbol_list,
        variables,
        out_folder=work_folder,
        num_symbols_per_call=num_symbols_per_call,
        restart=restart,
        **dstream_kwargs,
    )
    print(f'Combining all CSV files in {work_folder} into a single CSV file {outpath}')
    combine_all_datastream_csv(work_folder, outpath=outpath, restart=True)
