#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Script to search for records within an SQL database created using snapshot2sql
and convert to Researcher Format."""

# Import required modules
import getopt
from iams2rf import *

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'


def usage():
    """Function to print information about the script"""
    print('========================================')
    print('sql2rf')
    print('IAMS data extraction for Researcher Format')
    print('========================================')
    print('This utility searches an SQL database of IAMS records')
    print('created using the utility snapshot2sql')
    print('and converts matching records to Researcher Format')
    print('\nCorrect syntax is:')
    print('sql2rf -d DB_PATH -r REQUEST_PATH [OPTIONS]')
    print('\nSearch DB_PATH for records meeting criteria in REQUEST_PATH.')
    print('    -d    Path to the SQL database')
    print('    -r    Path to Outlook message containing details of the request')
    print('\nUse quotation marks (") around arguments which contain spaces')
    print('\nIf REQUEST_PATH is not specified you will be given the option to set parameters for the output')
    print('\nOptions:')
    print('    -o       OUTPUT_FOLDER to save output files.')    
    print('    --debug  Debug mode.')
    print('    --help   Show this message and exit.')
    exit_prompt()


def main(argv=None):
    if argv is None:
        name = str(sys.argv[1])

    db_path, request_path, output_folder = '', '', ''
    debug = False

    try:
        opts, args = getopt.getopt(argv, 'd:r:o:', ['db_path=', 'request_path=', 'output_folder=', 'debug', 'help'])
    except getopt.GetoptError as err:
        exit_prompt('Error: {}'.format(err))
    if opts is None or not opts:
        usage()
    for opt, arg in opts:
        if opt == '--help': usage()
        elif opt == '--debug': debug = True
        elif opt in ['-d', '--db_path']: db_path = arg
        elif opt in ['-r', '--request_path']: request_path = arg
        elif opt in ['-o', '--output_folder']: output_folder = arg
        else: exit_prompt('Error: Option {} not recognised'.format(opt))

    iams2rf_sql2rf(db_path, request_path, output_folder, debug)

    print('\n\nAll processing complete')
    print('----------------------------------------')
    print(str(datetime.datetime.now()))
    sys.exit()


if __name__ == '__main__':
    main(sys.argv[1:])
