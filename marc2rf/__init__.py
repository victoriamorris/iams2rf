#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Main package for marc2rf."""

# Import required modules
from .main import *

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'


def marc2rf_write_rf_config(request_path, output_folder, debug=False):
    """Prepare config files for selection of MARC records to convert
    to Researcher Format.

    :rtype: object
    :param request_path: Path to Outlook message containing details of the request.
    :param output_folder: Folder to save config files.
    :param debug: Display additional output to assist with debugging.
    """

    config = ConfigWriter(debug)
    if debug:
        print('Preparing config files with the following parameters:')
        print('request_path: {}'.format(str(request_path)))
        print('output_folder: {}'.format(str(output_folder)))
    config.marc2rf_write_rf_config(request_path, output_folder)


def marc2rf_researcherFormat(marc_path, request_path, output_folder, options, debug=False):
    """Convert MARC records to Researcher Format.

    :rtype: object
    :param marc_path: Path to file of MARC records.
    :param request_path: Path to Outlook message containing details of the request.
    :param output_folder: Folder to save Researcher Format output files.
    :param options: Options to set default transformation parameters.
    :param debug: Display additional output to assist with debugging.
    """

    converter = Converter(marc_path, request_path, output_folder, options, debug)
    if debug:
        print('Converting MARC records with the following parameters:')
        print('marc_path: {}'.format(str(marc_path)))
        print('request_path: {}'.format(str(request_path)))
        print('output_folder: {}'.format(str(output_folder)))
        print('options: {}'.format(str(options)))
    converter.marc2rf_researcherFormat()

