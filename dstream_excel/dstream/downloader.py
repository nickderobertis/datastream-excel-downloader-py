import argparse
import os
from concurrent.futures import ThreadPoolExecutor, TimeoutError

from dstream_excel.fileops import get_path_of_failed_folder_add_if_necessary, move_file_to_failed_folder
from exceldriver.tools import _start_excel_with_addins_and_attach, _restart_excel_with_addins_and_attach
from dstream_excel.tracker.files import FileProcessTracker
from .workbook.populate import populate_datastream_for_file


def populate_all_files_in_folder(folder: str, restart: bool = True):
    """
    Downloads data into Excel files which have been set up with Datastream functions.

    Opens excel files which already have Datastream functions, allows the data to populate, checks that
    it was populated, then closes the file and continues on to the next one.

    :param folder: Folder containing Excel files with Datastream functions pre-filled
    :param restart: Whether to force a restart of populating files. If the process is stopped, it will be continued
        where it left off if restart=False, and start from the beginning if restart=True.
    :return: None
    """
    excel = _start_excel_with_addins_and_attach(sleep=60)
    file_tracker = FileProcessTracker(folder=folder, restart=restart, file_types=('xlsx',))
    failed_folder = get_path_of_failed_folder_add_if_necessary(folder)

    with ThreadPoolExecutor(max_workers=1) as e:
        for i, file in enumerate(file_tracker.file_generator()):

            excel, successful = _try_to_get_result_if_fail_restart_excel(e, i, file, excel)

            if not successful:
                move_file_to_failed_folder(file, failed_folder)

def _try_to_get_result_if_fail_restart_excel(threadpool, i, file, excel, tries_remaining=3):

    if tries_remaining <= 0:
        print(fr'ERROR: Timed out processing {file}. Skipping and moving to "..\failed".')
        return excel, False

    # Submit result on threadpool asynchronously
    async_result = threadpool.submit(populate_datastream_for_file, file, excel, index=i + 1)

    # Try to get the result, waiting up to two minutes.
    try:
        excel, successful = async_result.result(timeout=300)
    except TimeoutError:
        excel = _restart_excel_with_addins_and_attach()
        return _try_to_get_result_if_fail_restart_excel(threadpool, i, file, excel, tries_remaining=tries_remaining - 1)


    return excel, successful



if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Populate data in XLSX files containing datastream functions')
    parser.add_argument('-f', '--folder', required=False,
                        default=r'C:\Users\derobertisna.UFAD\Dropbox (Personal)\UF\Andy\ETF Project\Data\Datastream\inprogress')
    parser.add_argument('-r', '--restart', action='store_true', default=False)
    args = parser.parse_args()

    folder = os.path.abspath(args.folder)

    populate_all_files_in_folder(folder, restart=args.restart)