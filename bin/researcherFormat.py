#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Script to convert MARC records to Researcher Format"""

import getopt
from marc2rf import *

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'


def usage():
    """Function to print information about the script"""
    print('========================================')
    print('researcherFormat')
    print('MARC record conversion for Researcher Format')
    print('========================================')
    print('This utility transforms a file of MARC records to Researcher Format')
    print('Correct syntax is:')
    print('researcherFormat -i MARC_PATH -r REQUEST_PATH -o OUTPUT_FOLDER [OPTIONS]\n')
    print('\nConvert MARC_PATH to Researcher Format with parameters set in REQUEST_PATH.')
    print('    -i    Path to file of MARC records')
    print('    -r    Path to Outlook message containing details of the request')
    print('    -o    Folder to save Researcher Format output files')
    print('\nUse quotation marks (") around arguments which contain spaces')
    print('\nIf REQUEST_PATH is not specified you will be given the option to set parameters for the output')
    print('\nOptions:')
    print('\nAt most one of ...')
    print('    -d       Default transformation.')
    print('    -b       Default transformation for BNB records.')
    print('    -c       Default transformation for Music records.')
    print('    -e       Default transformation for ESTC records.')
    print('    -f       Default transformation for FRBRization.')
    print('    -m       Use MARC fields instead of column headings.')
    print('    -n       Default transformation for Newspaper records.')
    print('\nAny of ...')
    print('    -o       OUTPUT_FOLDER to save output files.')
    print('    --debug  Debug mode.')
    print('    --help   Show this message and exit.')
    exit_prompt()


def main(argv=None):
    if argv is None:
        name = str(sys.argv[1])

    marc_path, request_path, output_folder, options = '', '', '', ''
    debug = False

    try:
        opts, args = getopt.getopt(argv, 'i:r:o:dbcefmn', ['request_path=', 'output_folder=', 'debug', 'help'])
    except getopt.GetoptError as err:
        exit_prompt('Error: {}'.format(err))
    if opts is None or not opts:
        usage()
    for opt, arg in opts:
        if opt == '--help': usage()
        elif opt == '--debug': debug = True
        elif opt in ['-i', '--marc_path']: marc_path = arg
        elif opt in ['-r', '--request_path']: request_path = arg
        elif opt in ['-o', '--output_folder']: output_folder = arg
        elif opt in ['-d', '-b', '-c', '-e', '-f', '-m', '-n']: options += opt
        else: exit_prompt('Error: Option {} not recognised'.format(opt))

    if len(re.sub(r'[^a-z]','',options)) > 1:
        exit_prompt('Error: too many optional parameters specified')

    marc2rf_researcherFormat(marc_path, request_path, output_folder, options, debug)

    print('\n\nAll processing complete')
    print('----------------------------------------')
    print(str(datetime.datetime.now()))
    sys.exit()

if __name__ == '__main__':
    main(sys.argv[1:])
