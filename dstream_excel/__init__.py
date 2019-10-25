"""
Use this tool to drive Excel using the Thompson Reuters Eikon plugin to download Datastream data. Useful
for downloading large amounts of data.
"""
from dstream_excel.main import download_datastream
from dstream_excel.dstream.keepopen.combine import combine_all_datastream_csv