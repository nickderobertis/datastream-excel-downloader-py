from typing import Sequence
from dstream_excel.dstream.downloader import populate_all_files_in_folder
from dstream_excel.dstream.combine import combine_all_datastream_xlsx
from dstream_excel.dstream.workbook.create import create_all_xlsx_with_datastream_command


def download_datastream(work_folder: str, symbol_list: Sequence[str], variables: Sequence[str],
                        outpath: str = 'data.csv', **dstream_kwargs):
    """
    Function for full workflow to download data from Datastream.

    Creates XLSX files with Datastream functions, opens those workbooks to populate the data, then
    combines the XLSX files into a single CSV.

    :param work_folder: Folder in which in process XLSX files will be generated
    :param symbol_list: Datastream symbols representing companies. Available by selecting a list through filters
        in Eikon, then getting an output of that list.
    :param variables: Time-series variables to download
    :param outpath: Full file path to output csv to
    :param dstream_kwargs: can pass begin, end, or freq. All of these arguments take day, month, and year as
        D, M, and Y. freq can be set to D, M, or Y. begin and end can be set to just Y but combined with
        a time modifier, e.g. begin='-5Y'. Don't pass end to have the data go up to current.
    :return: None
    """
    print(f'Creating XLSX files with functions in {work_folder}')
    create_all_xlsx_with_datastream_command(work_folder, symbol_list, variables, **dstream_kwargs)
    print(f'Populating XLSX files with data in {work_folder}')
    populate_all_files_in_folder(work_folder)
    print(f'Combining all XLSX files in {work_folder} into a single CSV file {outpath}')
    combine_all_datastream_xlsx(work_folder, outpath=outpath)