#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Script to prepare config files for selection of MARC records to convert
to Researcher Format"""

# Import required modules
import getopt
from marc2rf import *

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'


def usage():
    """Function to print information about the script"""
    print('========================================')
    print('write_rf_config')
    print('MARC record selection for Researcher Format')
    print('========================================')
    print('This utility prepares config files for selection of MARC records ')
    print('to convert to Researcher Format')
    print('Correct syntax is:')
    print('write_rf_config -r REQUEST_PATH [OPTIONS]\n')
    print('\nWrite config files for criteria in REQUEST_PATH.')
    print('    -r    Path to Outlook message containing details of the request')
    print('\nUse quotation marks (") around arguments which contain spaces')
    print('\nOptions:')
    print('    -o       OUTPUT_FOLDER to save output files.')
    print('    --debug  Debug mode.')
    print('    --help   Show this message and exit.')
    exit_prompt()


def main(argv=None):
    if argv is None:
        name = str(sys.argv[1])

    request_path, output_folder = '', ''
    debug = False

    try:
        opts, args = getopt.getopt(argv, 'r:o:', ['request_path=', 'output_folder=', 'debug', 'help'])
    except getopt.GetoptError as err:
        exit_prompt('Error: {}'.format(err))
    if opts is None or not opts:
        usage()
    for opt, arg in opts:
        if opt == '--help': usage()
        elif opt == '--debug': debug = True
        elif opt in ['-r', '--request_path']: request_path = arg
        elif opt in ['-o', '--output_folder']: output_folder = arg
        else: exit_prompt('Error: Option {} not recognised'.format(opt))

    marc2rf_write_rf_config(request_path, output_folder, debug)

    print('\n\nAll processing complete')
    print('----------------------------------------')
    print(str(datetime.datetime.now()))
    sys.exit()

if __name__ == '__main__':
    main(sys.argv[1:])
