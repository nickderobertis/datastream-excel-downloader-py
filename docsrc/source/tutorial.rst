.. _tutorial:

Getting started with dstream_excel
**********************************

Install
=======

Install this package via::

    pip install dstream_excel

You must also be running Windows, have Excel installed, have Eikon installed,
and have the Eikon plugin for Excel installed.

There is one setting in the Excel plugin that makes or breaks this
automation. Go to the Thompson Reuters Datatream tab, click Options.
In Calculation Options, make sure "Suppress Auto Calculation on Open"
is unchecked.

Also ensure that when you go to the Thomson Reuters and Thomson Reuters Datastream
tabs, that the buttons are not grayed out. If they are, you have to go
to the Thompson Reuters tab and click the "Offline" button. This should
launch Eikon on your system and possibly require you to log in. It may
be necessary to click the "Offline" button again after that. Once you are
logged in, all the buttons should be highlighted on both tabs. Then you can
close Excel and begin using :code:`dstream_excel`.

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
        begin='-1Y'
    )


You may see errors relating to calling Excel and that Excel has been terminated.
There is retry logic built into the package as Excel does not respond very
consistently in this way, so Excel may be terminated and restarted many
times in the process of downloading.

A `failed` folder will be created and any XLSX that were unable
to pull data after several retries will be moved here so that they can be
re-run later.

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