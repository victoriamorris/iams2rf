#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Script to convert the IAMS Published Snapshot to an SQL database."""

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
    print('snapshot2sql')
    print('IAMS data preparation for Researcher Format')
    print('========================================')
    print('This utility converts the IAMS Published Snapshot to an SQL database \n')
    print('\nCorrect syntax is:')
    print('snapshot2sql -i IAMS_SNAPSHOT_PATH -d DB_PATH [OPTIONS]')
    print('\nConvert IAMS_SNAPSHOT_PATH to an SQL database at DB_PATH.')
    print('    -i    Path to the IAMS Published Snapshot')
    print('    -d    Path to save the SQL database')
    print('\nUse quotation marks (") around arguments which contain spaces')
    print('\nOptions:')
    print('    --debug  Debug mode.')
    print('    --help   Show this message and exit.')
    exit_prompt()


def main(argv=None):
    if argv is None:
        name = str(sys.argv[1])

    iams_snapshot_path, db_path = '', ''
    debug = False

    try: opts, args = getopt.getopt(argv, 'i:d:', ['iams_snapshot_path=', 'db_path=', 'debug', 'help'])
    except getopt.GetoptError as err:
        exit_prompt('Error: {}'.format(err))
    if opts is None or not opts:
        usage()
    for opt, arg in opts:
        if opt == '--help':
            usage()
        elif opt == '--debug':
            debug = True
        elif opt in ['-i', '--iams_snapshot_path']:
            iams_snapshot_path = arg
        elif opt in ['-d', '--db_path']:
            db_path = arg
        else: exit_prompt('Error: Option {} not recognised'.format(opt))

    if debug:
        print('IAMS_SNAPSHOT_PATH: {}'.format(str(iams_snapshot_path)))
        print('DB_PATH: {}'.format(str(db_path)))

    iams2rf_snapshot2sql(iams_snapshot_path, db_path, debug)

    print('\n\nAll processing complete')
    print('----------------------------------------')
    print(str(datetime.datetime.now()))
    sys.exit()
        

if __name__ == '__main__':
    main(sys.argv[1:])
