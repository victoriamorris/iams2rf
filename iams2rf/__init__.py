#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Main package for iams2rf."""

# Import required modules
from .main import *
# import datetime
# import sys

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'


def iams2rf_snapshot2sql(iams_snapshot_path, db_path, debug=False):
    """Convert the IAMS Published Snapshot to an SQL database.

    :rtype: object
    :param iams_snapshot_path: Path to the IAMS Published Snapshot.
    :param db_path: Path to save the SQL database.
    :param debug: Display additional output to assist with debugging.
    """

    converter = IAMS2SQL(debug)
    if debug:
        print('Creating instance of IAMS2SQL class with the following parameters:')
        print('iams_snapshot_path: {}'.format(str(iams_snapshot_path)))
        print('db_path: {}'.format(str(db_path)))
    converter.iams2rf_snapshot2sql(iams_snapshot_path, db_path)


def iams2rf_sql2rf(db_path, request_path, output_folder, debug=False):
    """Search for records within an SQL database created using snapshot2sql
    and convert to Researcher Format

    :rtype: object
    :param db_path: Path to the SQL database.
    :param request_path: Path to Outlook message containing details of the request.
    :param output_folder: Folder to save Researcher Format output files.
    :param debug: Display additional output to assist with debugging.
    """

    converter = SQL2RF(debug)
    if debug:
        print('Creating instance of SQL2RF class with the following parameters:')
        print('db_path: {}'.format(str(db_path)))
        print('request_path: {}'.format(str(request_path)))
        print('output_folder: {}'.format(str(output_folder)))
    converter.iams2rf_sql2rf(db_path, request_path, output_folder)
