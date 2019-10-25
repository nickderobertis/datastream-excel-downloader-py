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

This package downloads the data in four main steps:

1. Opens a blank Excel workbook

2. Runs the Datastream Excel plugin functions for sets of companies in each call,
   by default 20 firms at a time.

3. After each function, it extracts the data and saves it into a CSV file

4. After all data is collected, it reads the data from all the generated CSV files
   and combines into a single CSV file.

Usage
=========

There are three main functions in the package.
:py:func:`.download_datastream` is the main function, which
is just a wrapper to call the other two main functions which complete
the four steps described above: :py:func:`.download_datastream_save_to_csvs`
and :py:func:`.combine_all_datastream_csv`.


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



How Fast is it?
==================

This is actually the second approach I have taken in this library to downloading the
data. This newer approach is 5-10x as fast as the old one. It allows downloading
thousands of firms' data overnight.

It is possible to squeeze a bit more speed by increasing `num_symbols_per_call`,
which is how many firms are being downloaded in one Excel function call. The
default is 20, but you may try increasing it, especially if you are not downloading
many `variables`. For example::

    from dstream_excel import download_datastream

    download_datastream(
        'output_folder',
        ['@AAPL','@MSFT'],
        ['UP', 'RI', 'NOSH', 'VO'],
        outpath='data.csv',
        freq='D',
        begin='-1Y',
        num_symbols_per_call=40  # can keep going higher until it gets unstable
    )

Once you increase this too much, the Excel function calls will start failing.

Troubleshooting
================

Failure during Download
-------------------------

Hopefully the :py:func:`.download_datastream` function works end-to-end. But
the steps 2 and 3 where the data is downloaded may fail. If this happens, get your
Excel working manually again (may
require a restart or re-enabling the Eikon plugin), then you can run
:py:func:`.download_datastream` with `restart=False` to
continue where it left off. `restart=False` is the default so you don't actually
need to pass it, but make sure that if you used `restart=True` before that you
have removed it. All the other options should be the same as your
original call to the function.

For example::

    from dstream_excel import download_datastream

    download_datastream(
        'output_folder',
        ['@AAPL','@MSFT'],
        ['UP', 'RI', 'NOSH', 'VO'],
        outpath='data.csv',
        freq='D',
        begin='-1Y',

        # optional, but don't set it to True if you want it to pick up where it left off!
        restart=False
    )

It's Failing a Lot
~~~~~~~~~~~~~~~~~~~~

The stability of the apporoach depends on how much data is being downloaded in each
function call. This depends both on `num_symbols_per_call` and the number of
`variables` being downloaded. I have set the default `num_symbols_per_call` to be
reasonable for up to 10-20 `variables`. If you are downloading more, you may need
to decrease `num_symbols_per_call`::

    from dstream_excel import download_datastream

    download_datastream(
        'output_folder',
        ['@AAPL','@MSFT'],
        ['UP', 'RI', 'NOSH', 'VO'],
        outpath='data.csv',
        freq='D',
        begin='-1Y',
        num_symbols_per_call=10  # keep setting lower if you are running into lots of crashing
    )

I have planned for a future version to automatically set a good default
`num_symbols_per_call` based on the number of variables being passed, but for
now it is 20 by default for any number of variables.

Failure during Combining
--------------------------

Sometimes it may happen that the data download finishes fine, but it fails during
the combining CSV step. So you are left with a lot of CSV files, not one single
file to work with. If this happens, you can manually run the
:py:func:`.combine_all_datastream_csv` function which is usually called as part of
:py:func:`.download_datastream`. For example::

    from dstream_excel import combine_all_datastream_csv

    combine_all_datastream_csv('output_folder', outpath='data.csv')

Note that calling it this way, it will attempt to resume from where it left off.
If for some reason this still doesn't work, you can also try::

    combine_all_datastream_csv('output_folder', outpath='data.csv', restart=True)

The `restart=True` when used in :py:func:`.combine_all_datastream_csv` is merely
talking about the combining process, not the download process. It will not remove
or overwrite anything you have downloaded. It will just start the combining process
over from scratch.