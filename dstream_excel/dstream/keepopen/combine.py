import shutil
from dstream_excel.tools.ext_pandas import append_csv_to_csv
from processfiles.files import FileProcessTracker


def combine_all_datastream_csv(infolder: str, outpath: str = 'all_data.csv', restart: bool = False):
    """
    Combine populated data in csv files into single csv file. This is normally run as a part of
    :py:func:`.download_datastream`, but if for some reason your CSV files don't get combined at the end,
    or you want to combine them to see a test sample while the download is still running,
    then you can run this manually.

    :param infolder: Folder containing populated csv files
    :param outpath: Full file path to output csv to
    :param restart: Whether to force a restart of combining files. If the process is stopped, it will be continued
        where it left off if restart=False, and start from the beginning if restart=True.
    :return: None
    """
    tracker = FileProcessTracker(infolder, restart=restart, file_types=('csv',))
    for i, file in enumerate(tracker.file_generator()):
        if i == 0:
            shutil.copy(file, outpath)
            continue
        append_csv_to_csv(file, outpath)
