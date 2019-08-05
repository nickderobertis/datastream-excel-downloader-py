.. _tutorial:

Getting started with dstream_excel
**********************************

Install
=======

Install this package via::

    pip install dstream_excel

You must also be running Windows, have Excel installed, have Eikon installed,
and have the Eikon plugin for Excel installed.

Overview
=========

This package downloads the data in three main steps:

1. Creates an XLSX workbook for each company containing the Excel function
for the Eikon Excel plugin to download the Datastream data

2. Opens each workbook, one by one, allowing the data to populate, then
closing and saving the workbook.

3. Reads the data from all the generated workbooks and combines into
one CSV file.

Usage
=========

There are four main functions in the package.
:py:func:`.download_datastream` is the main function, which
is just a wrapper to call the other three main functions which complete
the three steps described above: :py:func:`.create_all_xlsx_with_datastream_command`,
:py:func:`.populate_all_files_in_folder`, and :py:func:`.combine_all_datastream_xlsx`.


This is a simple example::

    from dstream_excel import download_datastream

    download_datastream(
        'output_folder',
        ['@AAPL','@MSFT'],  # Datastream-style tickers
        ['UP', 'RI', 'NOSH', 'VO'],  # Variable names from Datastream
        outpath='data.csv',
        freq='D',
        begin='-1M'
    )


Troubleshooting
================

Hopefully the :py:func:`.download_datastream` function works end-to-end. But
the second step where the files populated may cause Excel to fail. There is
some logic in the package to keep restarting Excel, but this may eventually
fail as well. If this happens, get your Excel working manually again (may
require a restart or re-enabling the Eikon plugin), then you can run
:py:func:`.populate_all_files_in_folder` while passing `restart=False` to
continue where it left off. Repeat this as many times as needed, then at
the end you can run :py:func:`.combine_all_datastream_xlsx`.

For example::

    from dstream_excel import (
        populate_all_files_in_folder,
        combine_all_datastream_xlsx
    )

    # Run as many times as needed, fixing Excel in between
    populate_all_files_in_folder(
        'output_folder',
        restart=False
    )

    # Once populate can actually finish, then run this
    combine_all_datastream_xlsx(
        'output_folder',
        outpath='data.csv'
    )