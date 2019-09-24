"""
Use this tool to drive Excel using the Thompson Reuters Eikon plugin to download Datastream data. Useful
for downloading large amounts of data.
"""
from dstream_excel.dstream.downloader import populate_all_files_in_folder
from dstream_excel.dstream.combine import combine_all_datastream_xlsx
from dstream_excel.dstream.workbook.create import create_all_xlsx_with_datastream_command
from dstream_excel.main import download_datastream