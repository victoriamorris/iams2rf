#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Main module for marc2rf."""

# Import required modules
# These should all be contained in the standard library
from collections import OrderedDict
import copy
import datetime
import gc
import glob
import locale
import os
import regex as re
import sys
import unicodedata

# Modules specific to Researcher Format
from marc2rf.lookup import *
from marc2rf.marc_data import *
from marc2rf.cleaning_functions import *

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'

# Set locale to assist with sorting
locale.setlocale(locale.LC_ALL, '')

# Disable garbage collection (except when called)
# Automatic garbage collection disrupts MultiRegex replacements
gc.disable()

# ====================
#       Classes
# ====================


class Output:

    def __init__(self, profile=None, initiate=False):
        if profile != 'F':
            self.headings = {
                'ID': 'BL record ID',
                'RT': 'Type of resource',
                'CT': 'Content type',
                'MT': 'Material type',
                'BN': 'BNB number',
                'LC': 'LC number',
                'OC': 'OCLC number',
                'ES': 'ESTC citation number',
                'AK': 'Archival Resource Key',
                'IB': 'ISBN',
                'IS': 'ISSN',
                'IL': 'ISSN-L',
                'IM': 'International Standard Music Number (ISMN)',
                'IR': 'International Standard Recording Code (ISRC)',  # NEW
                'IA': 'International Article Number (IAN)',   # NEW
                'PN': 'Publisher number',
                'AA': 'Name',
                'AD': 'Dates associated with name',
                'AT': 'Type of name',
                'AR': 'Role',
                'II': 'ISNI',
                'VF': 'VIAF',
                'AN': 'All names',
                'TT': 'Title',
                'TU': 'Uniform title',
                'TK': 'Key title',
                'TV': 'Variant titles',
                'S1': 'Preceding titles',
                'S2': 'Succeeding titles',
                'SE': 'Series title',
                'SN': 'Number within series',
                'PC': 'Country of publication',
                'PP': 'Place of publication',
                'PB': 'Publisher',
                'PD': 'Date of publication',
                'PU': 'Date of publication (not standardised)',
                'PJ': 'Projected date of publication',
                'PG': 'Publication date range',
                'P1': 'Publication date one',
                'P2': 'Publication date two',
                'FA': 'Free text information about dates of publication',
                'HF': 'First date held',
                'HL': 'Last date held',
                'HA': 'Free text information about holdings',
                'FC': 'Current publication frequency',
                'FF': 'Former publication frequency',
                'ED': 'Edition',
                'DS': 'Physical description',
                'SC': 'Scale',   # NEW
                'JK': 'Projection',   # NEW
                'CD': 'Coordinates',   # NEW
                'MF': 'Musical form',   # NEW
                'MG': 'Musical format',   # NEW
                'PR': 'Price',
                'DW': 'Dewey classification',
                'LN': 'Library of Congress classification',  # NEW
                'SM': 'BL shelfmark',
                'SD': 'DSC shelfmark',
                'SO': 'Other shelfmark',
                'BU': 'Burney?',
                'IO': 'India Office?',
                'CL': 'Formerly held at Colindale?',
                'SU': 'Topics',
                'G1': 'First geographical subject heading',
                'G2': 'Subsequent geographical subject headings',
                'CG': 'General area of coverage',
                'CC': 'Coverage: Country',
                'CF': 'Coverage: Region',
                'CY': 'Coverage: City',
                'GE': 'Genre',
                'LA': 'Languages',
                'CO': 'Contents',
                'AB': 'Abstract',
                'NN': 'Notes',
                'CA': 'Additional notes for cartographic materials',  # NEW
                'MA': 'Additional notes for music',  # NEW
                'PV': 'Provenance',
                'RF': 'Referenced in',
                'NL': 'Link to digitised resource',
                '8F': '852 holdings flag',
                'ND': 'NID',
                'EL': 'Encoding level',
                'SX': 'Status',
            }
        # Columns for FRBRization
        if profile == 'F':
            self.headings = {
                '001': '001',
                '130': '130',
                '240': '240',
                '245': '245',
                '246': '246',
                '100, 110, 111': '100, 110, 111',
                '700, 710, 711': '700, 710, 711',
                '250, 254': '250, 254',
                '260, 264': '260, 264',
                '015': '015',
                '020': '020',
                '300': '300',
                '336': '336',
                '337': '337',
                '338': '338',
                '500': '500',
                '600, 610, 611, 630, 648, 650, 651, 653': '600, 610, 611, 630, 648, 650, 651, 653',
                '852': '852',
                '985': '985',
            }
            self.values = OrderedDict([
                ('001', True), ('130', True), ('240', True), ('245', True), ('246', True), ('100, 110, 111', True),
                ('700, 710, 711', True), ('250, 254', True), ('260, 264', True), ('015', True), ('020', True),
                ('300', True), ('336', True), ('337', True), ('338', True), ('500', True),
                ('600, 610, 611, 630, 648, 650, 651, 653', True), ('852', True), ('985', True),
            ])
        elif profile == 'B':
            self.values = OrderedDict([
                ('DW', True), ('TT', True), ('PJ', True), ('IB', True), ('MT', True), ('AA', True), ('AD', True),
                ('AN', True), ('SU', True), ('LA', True), ('PP', True), ('PB', True), ('PR', True), ('DS', True),
                ('NN', True), ('SE', True), ('SN', True), ('CO', True), ('BN', True), ('ID', True),
                # END OF FIELDS USED BY THIS PROFILE
                ('RT', False), ('CT', False), ('LC', False), ('OC', False), ('ES', False), ('AK', False), ('IS', False),
                ('IL', False), ('IM', False), ('IR', False), ('IA', False), ('PN', False), ('AT', False), ('AR', False),
                ('II', False), ('VF', False), ('TU', False), ('TK', False), ('TV', False), ('S1', False), ('S2', False),
                ('PC', False), ('PD', False), ('PU', False), ('PG', False), ('P1', False), ('P2', False), ('FA', False),
                ('HF', False), ('HL', False), ('HA', False), ('FC', False), ('FF', False), ('ED', False), ('SC', False),
                ('JK', False), ('CD', False), ('MF', False), ('MG', False), ('LN', False), ('SM', False), ('SD', False),
                ('SO', False), ('BU', False), ('IO', False), ('CL', False), ('G1', False), ('G2', False), ('CG', False),
                ('CC', False), ('CF', False), ('CY', False), ('GE', False), ('AB', False), ('CA', False), ('MA', False),
                ('PV', False), ('RF', False), ('NL', False), ('8F', False), ('ND', False), ('EL', False), ('SX', False),
            ])
        elif profile == 'C':
            self.headings['AA'] = 'Composer'
            self.headings['AD'] = 'Composer life dates'
            self.headings['TU'] = 'Standardised title'
            self.values = OrderedDict([
                ('ID', True), ('AA', True), ('AD', True), ('TT', True), ('TU', True), ('TV', True), ('AN', True),
                ('PD', True), ('PU', True), ('PC', True), ('PP', True), ('PB', True), ('NN', True), ('MA', True),
                ('CO', True), ('RF', True), ('SU', True), ('ED', True), ('DS', True), ('MF', True), ('MG', True),
                ('SE', True), ('SN', True), ('IB', True), ('IM', True), ('PN', True), ('SM', True),
                # END OF FIELDS USED BY THIS PROFILE
                ('RT', False), ('CT', False), ('MT', False), ('BN', False), ('LC', False), ('OC', False), ('ES', False),
                ('AK', False), ('IS', False), ('IL', False), ('IR', False), ('IA', False), ('AT', False), ('AR', False),
                ('II', False), ('VF', False), ('TK', False), ('S1', False), ('S2', False), ('PJ', False), ('PG', False),
                ('P1', False), ('P2', False), ('FA', False), ('HF', False), ('HL', False), ('HA', False), ('FC', False),
                ('FF', False), ('SC', False), ('JK', False), ('CD', False), ('PR', False), ('DW', False), ('LN', False),
                ('SD', False), ('SO', False), ('BU', False), ('IO', False), ('CL', False), ('G1', False), ('G2', False),
                ('CG', False), ('CC', False), ('CF', False), ('CY', False), ('GE', False), ('LA', False), ('AB', False),
                ('CA', False), ('PV', False), ('NL', False), ('8F', False), ('ND', False), ('EL', False), ('SX', False),
            ])
        elif profile == 'E':
            self.values = OrderedDict([
                ('ID', True), ('ES', True), ('RT', True), ('AA', True), ('AD', True), ('AT', True), ('AR', True),
                ('AN', True), ('TT', True), ('PC', True), ('PP', True), ('PB', True), ('PD', True), ('DS', True),
                ('SU', True), ('LA', True), ('NN', True), ('PV', True), ('RF', True),
                # END OF FIELDS USED BY THIS PROFILE
                ('CT', False), ('MT', False), ('BN', False), ('LC', False), ('OC', False), ('AK', False), ('IB', False),
                ('IS', False), ('IL', False), ('IM', False), ('IR', False), ('IA', False), ('PN', False), ('II', False),
                ('VF', False), ('TU', False), ('TK', False), ('TV', False), ('S1', False), ('S2', False), ('SE', False),
                ('SN', False), ('PU', False), ('PJ', False), ('PG', False), ('P1', False), ('P2', False), ('FA', False),
                ('HF', False), ('HL', False), ('HA', False), ('FC', False), ('FF', False), ('ED', False), ('SC', False),
                ('JK', False), ('CD', False), ('MF', False), ('MG', False), ('PR', False), ('DW', False), ('LN', False),
                ('SM', False), ('SD', False), ('SO', False), ('BU', False), ('IO', False), ('CL', False), ('G1', False),
                ('G2', False), ('CG', False), ('CC', False), ('CF', False), ('CY', False), ('GE', False), ('CO', False),
                ('AB', False), ('CA', False), ('MA', False), ('NL', False), ('8F', False), ('ND', False), ('EL', False),
                ('SX', False),
            ])
        elif profile == 'N':
            self.headings['ID'] = 'Title ID'
            self.values = OrderedDict([
                ('ID', True), ('OC', True), ('LC', True), ('BN', True), ('IS', True), ('IL', True), ('TV', True),
                ('S1', True), ('S2', True), ('MT', True), ('G1', True), ('G2', True), ('PP', True), ('PC', True),
                ('CG', True), ('CY', True), ('P1', True), ('P2', True), ('FA', True), ('ED', True), ('FC', True),
                ('HF', True), ('HL', True), ('HA', True), ('BU', True), ('IO', True), ('CL', True), ('NL', True),
                # END OF FIELDS USED BY THIS PROFILE
                ('RT', False), ('CT', False), ('ES', False), ('AK', False), ('IB', False), ('IM', False), ('IR', False),
                ('IA', False), ('PN', False), ('AA', False), ('AD', False), ('AT', False), ('AR', False), ('II', False),
                ('VF', False), ('AN', False), ('TT', False), ('TU', False), ('TK', False), ('SE', False), ('SN', False),
                ('PB', False), ('PD', False), ('PU', False), ('PJ', False), ('PG', False), ('FF', False), ('DS', False),
                ('SC', False), ('JK', False), ('CD', False), ('MF', False), ('MG', False), ('PR', False), ('DW', False),
                ('LN', False), ('SM', False), ('SD', False), ('SO', False), ('SU', False), ('CC', False), ('CF', False),
                ('GE', False), ('LA', False), ('CO', False), ('AB', False), ('NN', False), ('CA', False), ('MA', False),
                ('PV', False), ('RF', False), ('8F', False), ('ND', False), ('EL', False), ('SX', False),
            ])
        elif profile == 'R':
            self.values = OrderedDict([
                ('ID', True), ('SD', True), ('SO', True), ('IB', True), ('IS', True), ('AA', True), ('AD', True),
                ('AT', True), ('AR', True), ('AN', True), ('TT', True), ('TV', True), ('SE', True), ('SN', True),
                ('PC', True), ('PP', True), ('PB', True), ('PD', True), ('PU', True), ('ED', True), ('DS', True),
                ('DW', True), ('SU', True), ('GE', True), ('LA', True), ('NN', True), ('RT', True), ('CT', True),
                ('MT', True), ('BN', True), ('LC', True), ('OC', True),
                # END OF FIELDS USED BY THIS PROFILE
                ('ES', False), ('AK', False), ('IL', False), ('IM', False), ('IR', False), ('IA', False), ('PN', False),
                ('II', False), ('VF', False), ('TU', False), ('TK', False), ('S1', False), ('S2', False), ('PJ', False),
                ('PG', False), ('P1', False), ('P2', False), ('FA', False), ('HF', False), ('HL', False), ('HA', False),
                ('FC', False), ('FF', False), ('SC', False), ('JK', False), ('CD', False), ('MF', False), ('MG', False),
                ('PR', False), ('LN', False), ('SM', False), ('BU', False), ('IO', False), ('CL', False), ('G1', False),
                ('G2', False), ('CG', False), ('CC', False), ('CF', False), ('CY', False), ('CO', False), ('AB', False),
                ('CA', False), ('MA', False), ('PV', False), ('RF', False), ('NL', False), ('8F', False), ('ND', False),
                ('EL', False), ('SX', False),
            ])

        else:
            self.values = OrderedDict([
                ('ID', True),  ('RT', True),  ('CT', False), ('MT', False), ('BN', True),  ('LC', False), ('OC', False),
                ('ES', False), ('AK', False), ('IB', True),  ('IS', False), ('IL', False), ('IM', False), ('IR', False),
                ('IA', False), ('PN', False), ('AA', True),  ('AD', True),  ('AT', True),  ('AR', True),  ('II', False),
                ('VF', False), ('AN', True),  ('TT', True),  ('TU', False), ('TK', False), ('TV', True),  ('S1', False),
                ('S2', False), ('SE', True),  ('SN', True),  ('PC', True),  ('PP', True),  ('PB', True),  ('PD', True),
                ('PU', False), ('PJ', False), ('PG', False), ('P1', False), ('P2', False), ('FA', False), ('HF', False),
                ('HL', False), ('HA', False), ('FC', False), ('FF', False), ('ED', True),  ('DS', True),  ('SC', False),
                ('JK', False), ('CD', False), ('MF', False), ('MG', False), ('PR', False), ('DW', True),  ('LN', False),
                ('SM', True),  ('SD', False), ('SO', False), ('BU', False), ('IO', False), ('CL', False), ('SU', True),
                ('G1', False), ('G2', False), ('CG', False), ('CC', False), ('CF', False), ('CY', False), ('GE', True),
                ('LA', True),  ('CO', False), ('AB', False), ('NN', True),  ('CA', False), ('MA', False), ('PV', False),
                ('RF', False), ('NL', False), ('8F', False), ('ND', False), ('EL', False), ('SX', False),
            ])
        if initiate:
            # Do I want to exclude more fields here?
            if profile == 'A':
                for v in self.values:
                    self.values[v] = True
            if profile == 'S':
                for v in ['8F', 'NL', 'P1', 'P2', 'SD', 'SO', 'SX']:
                    self.values[v] = False
        else:
            for v in self.values:
                self.values[v] = set()


class Converter(object):
    """A class for converting records.

    :param marc_path: Path to file of MARC records.
    :param request_path: Path to Outlook message containing details of the request.
    :param output_folder: Folder to save Researcher Format output files.
    :param options: Options to set default transformation parameters.
    :param debug: Display additional output to assist with debugging.
    """

    def __init__(self, marc_path, request_path, output_folder, options, debug=False):
        self.marc_path = marc_path
        self.request_path = request_path
        self.output_folder = output_folder
        self.options = re.sub(r'[^a-z]','',options)
        self.debug = debug
        self.header = '========================================\n' \
                      'researcherFormat\n' \
                      'MARC record conversion for Researcher Format\n' \
                      '========================================\n' \
                      'This program transforms a file of MARC records to Researcher Format\n'
        self.sources = set()
        self.output_fields = Output(initiate=True)
        self.profile = None
        # Parameters for type of records to be included
        self.main_cat, self.bnb, self.estc, self.iams, self.codes = True, False, False, False, False
        # Parameters for output files to be included
        self.file_records, self.file_titles, self.file_names, self.file_topics, self.file_classification = \
            False, False, False, False, False
        self.fields_present, self.nid_urls = {}, {}

    def show_header(self):
        if self.header:
            print(self.header)

    def write_readme(self):
        if self.profile in ['B', 'F', 'M', 'N', 'R']: return None

        sources = OrderedDict([
            ('bnb', [self.bnb, 'the British National Bibliography (BNB)', 'http://bnb.bl.uk (BNB)']),
            ('estc', [self.estc, 'the English Short Title Catalogue (ESTC)', 'http://estc.bl.uk (ESTC)']),
            ('main_cat', [self.main_cat, 'our British Library print collections', 'http://explore.bl.uk (published items)']),
            ('iams', [self.iams, 'our British Library manuscript collections', 'http://searcharchives.bl.uk (unpublished items)']),
        ])

        readme = open(os.path.join(self.output_folder, 'Readme.txt'), mode='w', encoding='utf-8', errors='replace')
        readme.write('README - Metadata subsets from the British Library\n\n'
                     '===================================================================================================='
                     '\n\nGENERAL INFORMATION ABOUT THE DATASETS\n\n')
        if self.profile == 'C':
            readme.write('The British Library \'Big Data Music Project\' dataset is a set of .csv (comma separated '
                         'value) files containing metadata relating to printed and manuscript sheet music. ')
        else:
            readme.write('This Researcher Format dataset is a set of .csv (comma separated value) files containing '
                         'metadata for resources from ')
            if sum([self.bnb, self.estc, self.main_cat, self.iams]) == 0:
                readme.write('our British Library collections.\n\n')
            else:
                readme.write(' and '.join((', '.join(sources[s][1] for s in sources if sources[s][0]).replace(
                    '{}, {}'.format(sources['main_cat'], sources['iams']),
                    'our British Library print and manuscript collections')).rsplit(', ', 1)) + '.\n\n')
        readme.write('There are {} files; each provides a different view of the data. '.format(
            str(sum([self.file_records, self.file_titles, self.file_names, self.file_topics, self.file_classification]))))
        readme.write('Each record in a .csv file contains the identifier ')
        if self.output_fields.values['ID'] or self.output_fields.values['ES']:
            readme.write('(')
            if self.output_fields.values['ID']:
                readme.write('BL record ID')
                if self.output_fields.values['ES']: readme.write(' and ')
            if self.output_fields.values['ES']: readme.write('ESTC citation number')
            readme.write(') ')
        readme.write('for the full metadata record held in the British Library\'s online catalogues. ')
        if self.estc:
            readme.write('The English Short Title Catalogue (ESTC) lists items published mainly in the British Isles '
                         'and North America between 1473 and 1800, from the collections of the British Library and '
                         'over 2,000 other libraries. ')
        readme.write('The online catalogue{} can be accessed at '.format(
            's' if sum([self.bnb, self.estc, self.main_cat, self.iams]) > 1 else ''))
        if sum([self.bnb, self.estc, self.main_cat, self.iams]) == 0:
            readme.write('http://explore.bl.uk.\n\n')
        else:
            readme.write(', and '.join((', '.join(sources[s][2] for s in sources if sources[s][0])).rsplit(', ', 1)) + '.\n\n')
        if self.estc:
            readme.write('===================================================================================================='
                         '\n\nRIGHTS\n\n'
                         'This dataset is being made available to support the researcher\'s own project; it is not for '
                         're-supply to any third party for commercial purposes.\n\n'
                         'Please contact estc@bl.uk before sharing this dataset with other researchers.\n\n'
                         'Please acknowledge the ESTC in any publications.\n\n'
                         '====================================================================================================\n\n')
        readme.write('CSV FILES\n\n')
        if self.file_records:
            readme.write('records.csv.\nA list of all resources. Includes: ')
            readme.write(', '.join(self.output_fields.headings[v] for v in self.output_fields.values if self.output_fields.values[v]))
            if self.profile != 'C':
                readme.write('.\nNote: we have used the column heading \'Name\' rather than \'Author\' to reflect the '
                             'fact that the names associated with a resource may be editors, artists, etc.\n\n')
        if self.file_names:
            readme.write('names.csv.\nA list of all names (including authors and editors, personal names and '
                         'organisations) associated with the resources. Includes: Name, Dates associated with name, '
                         'Type of name, Role')
            for v in ['II', 'VF']:
                if self.output_fields.values[v]:
                    readme.write(', {}'.format(self.output_fields.headings[v]))
            readme.write(', Other names, {}'.format(', '.join(self.output_fields.headings[v] for v in self.output_fields.values
                                                              if self.output_fields.values[v] and v not in ['AA', 'AD', 'AT', 'AR', 'II', 'VF', 'AN'])))
            readme.write('.\nNote: when a resource is associated with multiple names, each name appears as a separate'
                         ' entry within this file.\n\n')
        if self.file_titles:
            readme.write('titles.csv.\nA list of all titles (such as main titles, standardised titles, and variant'
                         ' titles) associated with the resources. Includes: Title, Other titles')
            readme.write(', '.join(self.output_fields.headings[v] for v in self.output_fields.values if self.output_fields.values[v]
                                   and v not in ['TT', 'TV', 'TU', 'TK']))
            readme.write('.\nNote: when a resource has multiple titles, each title appears as a separate entry within'
                         ' this file.\n\n')
        if self.file_topics:
            readme.write('topics.csv.\nA list of all topical/subject terms (including general terms, geographical'
                         ' terms, personal names, meetings/conferences and organisations) associated with the resources.'
                         ' Includes: Topic, Type of topic')
            readme.write(', '.join(
                self.output_fields.headings[v] for v in self.output_fields.values if self.output_fields.values[v]
                and v != 'SU'))
            readme.write('.\nNote: when a resource has multiple topical terms, each term appears as a separate entry'
                         ' within this file.\n\n')
        if self.file_classification:
            readme.write('classification.csv.\nA list of all Dewey classification numbers associated with the resources. Includes: Dewey classification')
            readme.write(', '.join(
                self.output_fields.headings[v] for v in self.output_fields.values if self.output_fields.values[v]
                and v != 'DW'))
            readme.write('.\nNote: when a resource has multiple Dewey classification numbers, each classification'
                         ' appears as a separate entry within this file.\n\n')
        readme.write('===================================================================================================='
                     '\n\nFORMAT OF THE DATA\n\n'
                     'Header row: The first row is a header row containing the name of the value e.g. \'Place of publication\'.'
                     '\n\n'
                     'Repeating values: Some cells may contain repeats of values separated with a delimiter e.g. '
                     '\'London ; New York\' in Place of publication. The two places are separated with the delimiter \';\'.'
                     '\n\n'
                     'Multiple facets: Some cells may contain multiple facets separated with a delimiter e.g. '
                     '\'Civil rights--History\' in Topics, where the sub-facet is separated with the delimiter \'--\'.'
                     '\n\n'
                     '===================================================================================================='
                     '\n\nSUPPORTING INFORMATION\n\n'
                     'Import issues: BL record IDs begin with at least one 0 (zero). Some import utilities may strip '
                     'these leading zeros. If you wish to open the files in Excel, for example, you will need to set '
                     'the data format to \'Text\' when importing the file.'
                     '\n\n'
                     'Character set issues: The raw bibliographic metadata is held in a library character set format, '
                     'MARC8. We have exported the csv files as UTF-8 but, depending on which utility you use to import '
                     'or analyse the data, there may be some instances where there are character set problems. These '
                     'may be circumvented if Unicode UTF-8 is an import option for the data format. For instance, if '
                     'you wish to open the files in Excel, you will need to import the files as \'data from text\', and '
                     'specify the character encoding as UTF-8; if you open the files in Excel without following this '
                     'import procedure, it is likely that letters with diacritics and other special characters will not '
                     'display properly.'
                     '\n\n'
                     'Data cleaning: Cataloguing rules and procedures have changed over time so there may be some '
                     'variations in the detail of each entry. Some cleaning of the data has been carried out to remove '
                     'unnecessary punctuation from the entries; however, given the varying nature of the original '
                     'catalogue records there may be some instances where this has not been successful. In many cases '
                     'these examples will be obvious because they will occur at the beginning of the table, prior to '
                     'the beginning of the alphabetical listing.\n\n')
        if self.estc or self.iams:
            readme.write('Variations in data from different sources: The metadata in the tables has come from different '
                         'sources. Some of the columns in the table will not be relevant to all data sources. For '
                         'example, records from the {} not have BNB numbers, ISBNs, or Dewey '
                         'classifications.\n\n'.format({(True, True): 'manuscript collections and ESTC may',
                                                        (True, False): 'ESTC may',
                                                        (False, True): 'manuscript collections will'}[(self.estc, self.iams)]))
        if self.iams:
            readme.write('Manuscript collections: Many manuscript collections are hierarchical. This hierarchy cannot '
                         'easily be represented within a CSV file. Within the online catalogue '
                         '(http://searcharchives.bl.uk) it is possible to navigate between related records within a '
                         'manuscript collection.\n\n')
        readme.write('Information about the Dewey classification system can be found at '
                     'http://en.wikipedia.org/wiki/List_of_Dewey_Decimal_classes.'
                     '\n\n'
                     'Information about importing CSV data into Microsoft Excel can be found at '
                     'https://support.office.com/en-ZA/article/import-data-using-the-text-import-wizard-40c6d5e6-41b0-4575-a54e-967bbe63a048.'
                     '\n\n'
                     'For further information or to comment, please contact metadata@bl.uk.')
        readme.close()

    def convert_record(self, record):
        """Function to convert a single MARC record to Researcher Format."""
        output = Output(self.profile)
        if self.profile == 'M':
            output.values = copy.deepcopy(self.fields_present)
            for v in output.values:
                output.values[v] = set()

        start_year, end_year = '', ''

        if self.profile == 'M':
            # Keep all MARC fields without cleaning
            for field in record.fields:
                if field.tag in marc_fields:
                    try:
                        if field.is_control_field():
                            # Don't strip subfield codes from control fields
                            output.values[field.tag].add(clean(field.data, space=False))
                        elif self.codes:
                            # Don't strip subfield codes if codes option is selected
                            output.values[field.tag].add((' '.join('$' + subfield[0] + clean(subfield[1], space=False) for subfield in field)).strip())
                        else:
                            # Strip subfield codes from data fields if codes option is not selected
                            output.values[field.tag].add((' '.join(clean(subfield[1], space=False) for subfield in field)).strip())
                    except: print('\nError [please report code 002 to VM]: {0}\n'.format(str(sys.exc_info())))

        elif self.profile == 'F':
            # Keep all relevant MARC fields without cleaning
            for field in record.fields:
                for v in self.output_fields.values:
                    if field.tag in v:
                        try:
                            if field.is_control_field():
                                # Don't strip subfield codes from control fields
                                output.values[v].add(clean(field.data, space=False))
                            elif field.tag == '985':
                                if re.search(r'dis[\s\-]*ag*reg', str(field).lower()) is not None \
                                        and re.search(r'dis[\s\-]*ag*reg', str(field).lower()) != '':
                                    output.values[v].add((' '.join('$' + subfield[0] + clean(subfield[1], space=False) for subfield in field)).strip())
                            else:
                                # Don't strip subfield codes
                                output.values[v].add((' '.join('$' + subfield[0] + clean(subfield[1], space=False) for subfield in field)).strip())
                        except: print('\nError [please report code 002FRBR to VM]: {0}\n'.format(str(sys.exc_info())))

        else:
            # Keep only selected fields with cleaning
            mtq = set()

            # LDR
            # CT    # Content type (from LDR/06)
            if content_types_ldr.get(record.leader[6]):
                output.values['CT'].add(content_types_ldr.get(record.leader[6]))
            # RT    # Type of resource (from LDR/07)
            if resource_types.get(record.leader[7]):
                output.values['RT'].add(resource_types.get(record.leader[7]))
            # EL    # Encoding level (from LDR/17)
            if encoding_levels.get(record.leader[17]):
                output.values['EL'].add(encoding_levels.get(record.leader[17]))

            # 001
            # ID    # BL record ID
            for field in record.get_fields('001'):
                if field.data != '':
                    output.values['ID'].add(clean(field.data))

            # 007
            # MT    # Material type
            for field in record.get_fields('007'):
                if field.data != '' and field.data[0].lower() == 'h':
                    output.values['MT'].add('Microfilm')

            # 008
            # PD    # Publication date
            #       # Serial start and end year
            # P1    # Publication date one
            # P2    # Publication date two
            # PC    # Publication country
            # LA    # Languages
            # FC    # Current publication frequency

            for field in record.get_fields('008'):

                try: date = re.sub(r'[^0-9]', '', field.data[7:11])
                except: pass
                else:
                    if len(date) == 4 and date != '9999':
                        start_year = date
                        output.values['P1'].add(date)

                try: date = re.sub(r'[^0-9]', '', field.data[11:15])
                except: pass
                else:
                    if len(date) == 4 and date != '9999':
                        end_year = date
                        output.values['P2'].add(date)

                date_type = field.data[6] if len(field.data) >= 6 else 's'

                # Records with a single publication date
                if date_type in ['e', 'p', 'r', 's', 't']:
                    if start_year != '':
                        output.values['PD'].add(start_year)

                # Records with a range of publication dates
                elif date_type in ['c', 'd', 'i', 'k', 'm', 'q', 'u']:
                    if start_year != '' and end_year != '':
                        output.values['PD'].add('{}-{}'.format(start_year, end_year))
                    elif start_year != '':
                        output.values['PD'].add('{}-'.format(start_year))
                    elif end_year != '':
                        output.values['PD'].add('-{}'.format(end_year))

                try: country = countries.get(re.sub(r'[^a-z]', '', field.data[15:18]), '')
                except: pass
                else: output.values['PC'].add(country)

                if self.profile == 'N':
                    try: frequency = frequencies.get(re.sub(r'[^a-z]', '', field.data[18]), '')
                    except: pass
                    else: output.values['FC'].add(frequency)

                try: language = languages.get(re.sub(r'[^a-z]', '', field.data[35:38]), '')
                except: pass
                else: output.values['LA'].add(language)

            # 010
            # LC    # LC number
            for field in record.get_fields('010'):
                for subfield in field.get_subfields('a'):
                    if subfield != '':
                        output.values['LC'].add(subfield)

            # 015
            # BN    # BNB number
            for field in record.get_fields('015'):
                if 'bnb' in ' '.join(field.get_subfields('2')).lower():
                    for subfield in field.get_subfields('a'):
                        if subfield != '':
                            output.values['BN'].add(subfield)

            # 019
            # IO    # India Office?
            for field in record.get_fields('019'):
                if field.indicator1 == '2':
                    for subfield in field.get_subfields('a'):
                        if subfield[:3].lower() == 'iol':
                            output.values['IO'].add('Y')

            # 020
            # IB    # ISBN
            # PR    # Price
            # Material type qualifier
            for field in record.get_fields('020'):
                for subfield in field.get_subfields('a'):
                    subfield = re.sub(r'[^0-9X]', '', subfield.upper())
                    if is_isbn_10(subfield):
                        output.values['IB'].add(isbn_convert(subfield))
                    if is_isbn_13(subfield):
                        output.values['IB'].add(subfield)

                for subfield in field.get_subfields('a'):
                    try:
                        subfield = clean(subfield.split('(')[1].split(')')[0].replace('corrected', ''))
                    except:
                        pass
                    else:
                        if subfield.lower() not in 'hbk pbk hardback paperback pack':
                            mtq.add(subfield)

                for subfield in field.get_subfields('c', cleaning=False):
                    # set cleaning=False to avoid stripping $ from prices in dollars
                    subfield = re.sub(r'[^0-9£$\u20AC.]', '', subfield.upper().replace('EUR', u'\u20AC'))
                    if re.sub(r'[^0-9]', '', subfield) != '' and len(subfield) <= 9:
                        output.values['PR'].add(subfield)

            # 022
            # IS    # ISSN
            # IL    # ISSN-L
            for field in record.get_fields('022'):
                for subfield in field.get_subfields('a'):
                    subfield = re.sub(r'[^0-9X]', '', subfield.upper())
                    if len(subfield) == 8:
                        output.values['IS'].add(subfield[:4] + '-' + subfield[4:])

                for subfield in field.get_subfields('l'):
                    subfield = re.sub(r'[^0-9X]', '', subfield.upper())
                    if len(subfield) == 8:
                        output.values['IL'].add(subfield[:4] + '-' + subfield[4:])

            # 024
            # IM    # International Standard Music Number (ISMN)
            # IR    # International Standard Recording Code (ISRC)
            # IA    # International Article Number (IAN)
            for field in record.get_fields('024'):
                if field.indicator1 in ['0', '2', '3']:
                    for subfield in field.get_subfields('a'):
                        subfield = re.sub(r'[^0-9X]', '', subfield.upper())
                        if subfield != '':
                            if field.indicator1 == '0':
                                output.values['IR'].add(subfield)
                            elif field.indicator1 == '2':
                                output.values['IM'].add(subfield)
                            elif field.indicator1 == '3':
                                output.values['IA'].add(subfield)

            # 028
            # PN    # Publisher number
            for field in record.get_fields('028'):
                for subfield in field.get_subfields('a'):
                    if subfield != '':
                        output.values['PN'].add(subfield)

            # 034
            # SC    # Scale
            # CD    # Coordinates
            # CA    # Additional notes for cartographic materials
            for field in record.get_fields('034'):
                scale = '-'.join(sorted('1:{}'.format(re.sub(r'[^0-9]', '', subfield))
                                        for subfield in field.get_subfields('b')
                                        if re.sub(r'[^0-9]', '', subfield) != ''))
                if scale != '':
                    output.values['SC'].add(scale)
                coordinates = ', '.join(field.get_subfields('d', 'e', 'f', 'g', cleaning=False))
                if coordinates != '':
                    output.values['CD'].add(scale)
                notes = ''
                for subfield in field:
                    if subfield[0] in ['c', 'h', 'j', 'k', 'm', 'n', 'p', 'r', 's', 't', 'x', 'y', 'z'] \
                            and subfield[1].strip() != '':
                        notes = add_string({'c': ' Constant ratio linear vertical scale: ',
                                            'h': ' Angular scale: ',
                                            'j': ' Declination - northern limit: ',
                                            'k': ' Declination - southern limit: ',
                                            'm': 'Right ascension - eastern limit: ',
                                            'n': 'Right ascension - western limit: ',
                                            'p': 'Equinox: ',
                                            'r': 'Distance from earth: ',
                                            's': 'G-ring latitude: ',
                                            't': 'G-ring longitude: ',
                                            'x': 'Beginning date: ',
                                            'y': 'Ending date: ',
                                            'z': 'Name of extraterrestrial body: '}[subfield[0]] + subfield[1].strip(),
                                           notes, '. ')
                if notes != '':
                    output.values['CA'].add(notes)

            # 035
            # ES    # ESTC number
            # OC    # OCLC number
            for field in record.get_fields('035'):
                for subfield in field.get_subfields('a'):
                    if '(CU-RivES)' in subfield:
                        output.values['ES'].add(subfield.replace('(CU-RivES)', ''))
                    elif '(OCoLC)' in subfield:
                        output.values['OC'].add(subfield.replace('(OCoLC)', ''))

            # 041
            # LA    # Languages
            for field in record.get_fields('041'):
                for subfield in field.get_subfields('a', 'b', 'd', 'e', 'f', 'g', 'j', 'm'):
                    try: subfield = languages.get(subfield, '')
                    except: pass
                    else:
                        if subfield != '':
                            output.values['LA'].add(subfield)

            # 047
            # MF    # Musical form
            for field in record.get_fields('047'):
                if field.indicator2 == ' ' or 'marcmuscomp' in '|'.join(field.get_subfields('2')):
                    for subfield in field.get_subfields('a'):
                        try: subfield = musical_forms.get(subfield, '')
                        except: pass
                        else:
                            if subfield != '':
                                output.values['MF'].add(subfield)

            # 050
            # LN    # Library of Congress classification
            for field in record.get_fields('050'):
                for subfield in field.get_subfields('a'):
                    subfield = re.sub(r'[^0-9A-Z.]', '', subfield.upper())
                    if subfield != '':
                        output.values['LN'].add(subfield)

            # 082
            # DW    # Dewey classification
            for field in record.get_fields('082'):
                for subfield in field.get_subfields('a'):
                    subfield = re.sub(r'[^0-9.]', '', subfield)
                    if '.' in subfield:
                        subfield = ('0000' + subfield.split('.', 1)[0])[-3:] + '.' \
                                   + subfield.split('.')[1].replace('.', '')
                    output.values['DW'].add(subfield)

            # 100 110 111
            # AN    # All names
            # AA    # Name
            # AD    # Dates
            # AR    # Role
            # AT    # Type of name
            # II    # ISNI
            # VF    # VIAF
            for field in record.get_fields('100', '110', '111'):
                name = get_name_parts(field)
                output.values['AN'].add(name)
                if len(output.values['AA']) == 0:
                    output.values['AA'].add(name[0])
                    output.values['AD'].add(name[1])
                    output.values['AT'].add(name[2])
                    output.values['AR'].add(name[3])
                    output.values['II'].add(name[4])
                    output.values['VF'].add(name[5])

            # 130 240
            # TU    # Uniform title
            # TV    # Variant titles
            # ED    # Edition (for newspapers)
            for field in record.get_fields('130', '240'):
                title = ''
                for subfield in field:
                    code, content = subfield[0], clean(subfield[1])
                    if code == 'a':
                        title = add_string(content, title, ' ')
                    elif code in ['d', 'f', 'r']:
                        title = add_string(content, title, ', ')
                    elif code in ['h', 'k', 'l', 'm', 'n', 'o', 'p', 's', 't']:
                        title = add_string(content, title, '. ')
                title = quick_clean(title)
                if title != '':
                    output.values['TU'].add(title)
                    output.values['TV'].add(title)

                if self.profile == 'N':
                    for subfield in field:
                        code, content = subfield[0], clean_250(subfield[1])
                        if code in ['f', 'l', 's']:
                            output.values['ED'].add(content)

            # 222 245 246 247 730 740
            # TK    # Key title for serials
            # TT    # Title
            # TV    # Variant titles
            # See also 700 $t
            for field in record.get_fields('222', '245', '246', '247', '730', '740'):
                title = ''
                for subfield in field:
                    code, content = subfield[0], clean(subfield[1])
                    if code == 'a': title = add_string(content, title, ' ')
                    elif code == 'b': title = add_string(content, title, ' : ')
                    elif code in ['n', 'p']: title = add_string(content, title, '. ')
                title = quick_clean(title)
                if title != '' and not title.startswith('$u'):
                    if field.tag == '245':
                        output.values['TT'].add(title)
                    elif field.tag == '222':
                        output.values['TK'].add(title)
                    output.values['TV'].add(title)

            # 250
            # ED    # Edition
            for field in record.get_fields('250'):
                edition = ''
                for subfield in field:
                    code, content = subfield[0], clean_250(subfield[1])
                    if code == 'a':
                        if len(content) >= 2: content = content[0].upper() + content[1:]
                        edition = add_string(content, edition, ' ')
                    elif code == 'b':
                        edition = add_string(content, edition, ', ')
                edition = quick_clean(edition)
                if edition != '':
                    output.values['ED'].add(edition)

            # 254
            # ED    # Edition
            for field in record.get_fields('254'):
                for subfield in field.get_subfields('a'):
                    if subfield != '':
                        output.values['ED'].add(subfield)

            # 255
            # JK    # Projection
            # CA    # Additional notes for cartographic materials
            for field in record.get_fields('255'):
                for subfield in field.get_subfields('a'):
                    if subfield != '' and 'not given' not in subfield:
                        output.values['CA'].add(subfield)
                for subfield in field.get_subfields('b'):
                    subfield = re.sub(r'(?<![a-z])proj\.?(?![a-z])', 'projection', subfield, flags=re.IGNORECASE)
                    if subfield != '':
                        output.values['JK'].add(subfield)
                for subfield in field.get_subfields('c', cleaning=False):
                    if subfield != '':
                        output.values['CA'].add('Coordinates: {}'.format(subfield))
                for subfield in field:
                    if subfield[0] in ['d', 'e', 'f', 'g'] and subfield[1] != '':
                        output.values['CA'].add({'d': 'Zone: ',
                                                 'e': 'Equinox: ',
                                                 'f': 'Outer G-ring: ',
                                                 'g': 'Exclusion G-ring: '}[subfield[0]] + subfield[1])

            # 260 264
            # PP    # Place of publication
            # PB    # Publisher
            # PD    # Publication date
            # PU    # Publication date (uncleaned)
            for field in record.get_fields('260', '264'):
                for subfield in field.get_subfields('a'):
                    subfield = clean_26X(subfield)
                    # Test for a date at the end of the subfield
                    rx = re.compile('(,\s+|^)(©|\u00A9|c)?[0-9\-]{4}[0-9\-]*($|\s*\()')
                    if rx.search(subfield):
                        if not quick_clean(rx.search(subfield).group(0)) == '':
                            output.values['PU'].add(quick_clean(rx.search(subfield).group(0)))
                            subfield = quick_clean(subfield.replace(rx.search(subfield).group(0), ''))

                    if ':' in subfield:
                        publishers, states, places = clean_publication_places(subfield.split(':', 1)[0], output.values['PC'])
                        for item in publishers: output.values['PB'].add(item.strip())
                        for item in states: output.values['PC'].add(item.strip())
                        for item in places: output.values['PP'].add(item.strip())
                        subfield = subfield.split(':', 1)[1]
                        publishers, states, places = clean_publisher_names(subfield)
                        for item in publishers: output.values['PB'].add(item.strip())
                        for item in states: output.values['PC'].add(item.strip())
                        for item in places: output.values['PP'].add(item.strip())
                    else:
                        publishers, states, places = clean_publication_places(subfield, output.values['PC'])
                        for item in publishers: output.values['PB'].add(item.strip())
                        for item in states: output.values['PC'].add(item.strip())
                        for item in places: output.values['PP'].add(item.strip())

                for subfield in field.get_subfields('b'):
                    subfield = clean_26X(subfield)
                    publishers, states, places = clean_publisher_names(subfield)
                    for item in publishers: output.values['PB'].add(item.strip())
                    for item in states: output.values['PC'].add(item.strip())
                    for item in places: output.values['PP'].add(item.strip())

                for subfield in field.get_subfields('c'):
                    subfield = clean_26X(subfield)
                    subfield = quick_clean(re.sub(r'[.\[\]?:;]', '', re.sub(r'\[sic\.?\]', '', subfield, flags=re.IGNORECASE)))
                    if subfield != '':
                        output.values['PU'].add(subfield)
                        subfield = re.sub(r'[^0-9]', '', subfield)
                        # Publication date in 260/264 is only used if no date found in 008
                        if subfield != '' and len(subfield) >= 4 and len(output.values['PD']) == 0:
                            output.values['PD'].add(subfield[0:4])

            # 263
            # PJ    # Projected publication date
            for field in record.get_fields('263'):
                for subfield in field.get_subfields('a'):
                    subfield = re.sub(r'[^0-9]', '', subfield)
                if len(subfield) == 6:
                    output.values['PJ'].add(subfield[0:4] + '-' + subfield[4:])
                elif len(subfield) == 4:
                    output.values['PJ'].add(subfield)

            # 300
            # DS    # Physical description
            for field in record.get_fields('300'):
                description = clean_300(field)
                output.values['DS'].add(description)

            # 310 321
            # FC    # Current publication frequency
            # FF    # Former publication frequency
            for field in record.get_fields('310', '321'):
                frequency = ''
                dates = set()
                for subfield in field.get_subfields('a'):
                    subfield = clean_310(subfield)
                    frequency = add_string(subfield, frequency, ' ')
                for subfield in field.get_subfields('b'):
                    dates.add(get_date_range(subfield))
                dates = ' ; '.join(dates).replace('- ; -', '-')
                if len(re.sub(r'[^;]', '', dates)) == 1:
                    dates = re.sub(r'\s+', ' ', re.sub(r'([^;]*);([^;]*)', r'\2 ; \1', dates)).replace(
                        '- ; -', '-').strip()
                dates = quick_clean(dates)
                if dates != '' and frequency != '':
                    frequency = add_string('(' + dates + ')', frequency, ' ')
                if frequency != '':
                    if field.tag == '310' or self.profile == 'N':
                        output.values['FC'].add(frequency)
                    elif field.tag == '321':
                        output.values['FF'].add(frequency)

            # 336
            # CT    # Content type
            # See also end of record processing for values from leader
            for field in record.get_fields('336'):
                if 'rdacontent' in ' '.join(field.get_subfields('2')).lower():
                    for subfield in field.get_subfields('a', 'b'):
                        if len(subfield) == 3:
                            subfield = content_types.get(subfield.lower(), '')
                        else: subfield = subfield.capitalize()
                        if subfield != '' and subfield in content_types.values():
                            output.values['CT'].add(subfield)

            # 338
            # MT    # Material type
            for field in record.get_fields('338'):
                if 'rdacarrier' in ' '.join(field.get_subfields('2')).lower():
                    for subfield in field.get_subfields('a', 'b'):
                        if len(subfield) == 2:
                            subfield = material_types.get(subfield.lower(), '')
                        else:
                            subfield = subfield.capitalize()
                            if subfield == 'Online': subfield = 'Online resource'
                        if subfield != '' and subfield in material_types.values():
                            output.values['MT'].add(subfield)

            # 362
            # PG    # Publication date range (for serials)
            pg = set()
            for field in record.get_fields('362'):
                date_range = clean_362(field, start_year, end_year)
                pg.add(date_range)
            pg = ' ; '.join(pg).replace('- ; -', '-')
            if len(re.sub(r'[^;]', '', pg)) == 1:
                pg = re.sub(r'\s+', ' ', re.sub(r'([^;]*);([^;]*)', r'\2 ; \1', pg)).replace('- ; -', '-').strip()
            for sub_range in pg.split(';'):
                output.values['PG'].add(sub_range.strip())
                output.values['FA'].add(sub_range.strip())

            # 348
            # MG    # Musical format
            for field in record.get_fields('348'):
                for subfield in field.get_subfields('a'):
                    if subfield != '':
                        output.values['MG'].add(subfield)

            # 382
            # MA    # Additional notes for music
            for field in record.get_fields('382'):
                for subfield in field.get_subfields('a'):
                    if subfield != '':
                        output.values['MG'].add('Medium of performance: {}'.format(subfield))
                for subfield in field.get_subfields('b'):
                    if subfield != '':
                        output.values['MG'].add('Soloist: {}'.format(subfield))

            # 383
            # MA    # Additional notes for music
            for field in record.get_fields('383'):
                if quick_clean(str(field)) != '':
                    output.values['MA'].add(quick_clean(str(field)))

            # 384
            # MA    # Additional notes for music
            for field in record.get_fields('384'):
                for subfield in field.get_subfields('a'):
                    if subfield != '':
                        output.values['MG'].add('Key: {}'.format(subfield))

            # 490
            # SE    # Series title
            # SN    # Number within series
            for field in record.get_fields('490'):
                title, number = get_series_parts(field)
                if title != '':
                    output.values['SE'].add(title)
                    if number != '':
                        number = number + ' [' + title + ']'
                        output.values['SN'].add(number)

            # 500 515
            # NN    # Notes
            for field in record.get_fields('500', '515'):
                notes = ''
                for subfield in field.get_subfields('a'):
                    subfield = clean_500(subfield)
                    notes = add_string(subfield, notes, ' ')
                output.values['NN'].add(notes)
                if field.tag == '515' and 'micro' not in notes.lower():
                    output.values['FA'].add(notes)

            # 505
            # CO    # Contents
            for field in record.get_fields('505'):
                contents = ''
                for subfield in field.get_subfields('a', 'g', 'r', 't'):
                    contents = add_string(subfield, contents, ' ')
                output.values['CO'].add(contents)

            # 510
            # RF    # References
            for field in record.get_fields('510'):
                references = ''
                for subfield in field.get_subfields('a', 'b', 'c', 'u', 'x'):
                    subfield = clean_510(subfield)
                    references = add_string(subfield, references, ' ')
                output.values['RF'].add(references)

            # 520
            # AB    # Abstract
            for field in record.get_fields('520'):
                abstract = ''
                for subfield in field.get_subfields('a', 'b', 'c'):
                    abstract = add_string(subfield, abstract, ' ')
                output.values['AB'].add(abstract)

            # 561
            # PV    # Provenance
            for field in record.get_fields('561'):
                if field.indicator1 != '0':  # 1st indicator 0 denotes private notes
                    for subfield in field.get_subfields('a'):
                        output.values['PV'].add(subfield)

            # 600 610 611 630 650 651 653
            # SU    # Topics
            # GE    # Genre
            # G1    # First geographical subject heading
            # G2    # Subsequent geographical subject headings
            for field in record.get_fields('600', '610', '611', '630', '650', '651', '653'):
                topic = get_topic_parts(field)
                for genre in topic[2]:
                    output.values['GE'].add(genre)
                topic = topic[0], topic[1]
                output.values['SU'].add(topic)

            for field in record.get_fields('651'):
                for subfield in field.get_subfields('a', 'z'):
                    if len(output.values['G1']) == 0:
                        output.values['G1'].add(subfield)
                    elif subfield not in output.values['G1']:
                        output.values['G2'].add(subfield)
                    # For newspapers, geographical subject headings
                    # are also used to detect country and place of publication
                    if self.profile == 'N' and len(output.values['PC']) == 0:
                        for country in countries.values():
                            if country in subfield: output.values['PC'].add(country)
                        subfield = subfield.title()
                        if subfield in PLACES_ENGLAND:
                            output.values['PC'].add('England')
                        elif subfield in PLACES_IRELAND:
                            output.values['PC'].add('Ireland')
                        elif subfield in PLACES_N_IRELAND:
                            output.values['PC'].add('Northern Ireland')
                        elif subfield in PLACES_SCOTLAND:
                            output.values['PC'].add('Scotland')
                        elif subfield in PLACES_WALES:
                            output.values['PC'].add('Wales')
                        elif subfield in PLACES_US:
                            output.values['PC'].add('United States of America')
                        elif 'Great Britain' in subfield:
                            output.values['PC'].add('United Kingdom')
                        for city in PLACES:
                            if city in subfield: output.values['PP'].add(city)

            # 648
            # SU    # Topics
            for field in record.get_fields('648'):
                for subfield in field.get_subfields('a'):
                    topic = subfield, 'chronological term'
                    output.values['SU'].add(topic)

            # 655
            # GE    # Genre
            for field in record.get_fields('655'):
                for subfield in field.get_subfields('a'):
                    # Check for information about place of publication at the end of the field
                    if '.- ' in subfield:
                        publishers, states, places = clean_publication_places(subfield.split('.- ', 1)[1], output.values['PC'])
                        for item in states: output.values['PC'].add(item.strip())
                        for item in places: output.values['PP'].add(item.strip())
                    subfield = clean_genre(subfield)
                    output.values['GE'].add(subfield)

            # 700 710 711
            # AN    # All names
            for field in record.get_fields('700', '710', '711'):
                name = get_name_parts(field)
                output.values['AN'].add(name)

            # 700
            # TV    # Variant titles
            for field in record.get_fields('700'):
                for subfield in field.get_subfields('t'):
                    output.values['TV'].add(subfield)

            # 752
            # CC    # Coverage: Country
            # CF    # Coverage: Region
            # CY    # Coverage: City
            # CG    # Covegage: General
            # PP    # Place of publication
            # CL    # Formerly held at Colindale?
            for field in record.get_fields('752'):
                if 'blnpn' in ' '.join(field.get_subfields()).lower():
                    place = ''
                    if '2' in field:
                        output.values['CL'].discard('Maybe')
                        output.values['CL'].add('Y')
                    # $a - Country or larger entity (R)
                    for subfield in field.get_subfields('a'):
                        subfield = expand_place_abbreviations(subfield, output.values['PC'])
                        place = add_string(subfield, place, '--')
                        output.values['CC'].add(subfield)
                        if subfield in countries.values(): output.values['PC'].add(subfield)
                    # $b - First-order political jurisdiction (NR)
                    for subfield in field.get_subfields('b'):
                        subfield = expand_place_abbreviations(subfield, output.values['PC'])
                        place = add_string(subfield, place, '--')
                        output.values['CF'].add(subfield)
                    # $c - Intermediate political jurisdiction (R)
                    for subfield in field.get_subfields('c'):
                        subfield = expand_place_abbreviations(subfield, output.values['PC'])
                        place = add_string(subfield, place, '--')
                    # $d - City (NR)
                    for subfield in field.get_subfields('d'):
                        subfield = expand_place_abbreviations(subfield, output.values['PC'])
                        output.values['CY'].add(subfield)
                        if subfield in PLACES_ENGLAND:
                            output.values['PC'].add('England')
                        elif subfield in PLACES_IRELAND:
                            output.values['PC'].add('Ireland')
                        elif subfield in PLACES_N_IRELAND:
                            output.values['PC'].add('Northern Ireland')
                        elif subfield in PLACES_SCOTLAND:
                            output.values['PC'].add('Scotland')
                        elif subfield in PLACES_WALES:
                            output.values['PC'].add('Wales')
                        elif subfield in PLACES_US:
                            output.values['PC'].add('United States of America')
                    output.values['CG'].add(place)
                else:
                    for subfield in field.get_subfields('a'):
                        subfield = expand_place_abbreviations(subfield, output.values['PC'])
                        if subfield in countries.values(): output.values['PC'].add(subfield)
                    for subfield in field.get_subfields('b', 'c', 'd'):
                        subfield = expand_place_abbreviations(subfield, output.values['PC'])
                        output.values['PP'].add(subfield)

            # 760 762 770 772 773 774 775 776
            # NN    # Additional information for serials
            for field in record.get_fields('760', '762', '770', '772', '773', '774', '775', '776'):
                notes = ''
                preamble = {
                    '760': 'Main series: ',
                    '762': 'Subseries: ',
                    '770': 'Special issue: ',
                    '772': 'Supplement parent: ',
                    '773': 'Host item: ',
                    '774': 'Constituent unit: ',
                    '775': 'Other edition: ',
                    '776': 'Additional physical form: ',
                    '880': '',
                }[field.tag]
                for subfield in field.get_subfields('a'):
                    notes = add_string(subfield, notes, ' ')
                for subfield in field.get_subfields('b', 'c', 'd', 'g', 'h', 'i', 'k', 'm', 'n', 'o', 'r', 's', 't', 'z'):
                    notes = add_string(subfield, notes, '. ')
                for subfield in field.get_subfields('x'):
                    subfield = re.sub(r'[^0-9xX]', '', subfield)
                    if len(subfield) == 8:
                        notes = add_string('(ISSN: {}-{})'.format(subfield[:4], subfield[4:]), notes, ' ')
                if notes != '':
                    notes = preamble + notes
                    output.values['NN'].add(notes)

            # 780
            # NN    # Additional information for serials
            # S1    # Preceding titles
            for field in record.get_fields('780'):
                notes = ''
                try:
                    preamble = {
                        '0': 'Continues: ',
                        '1': 'Continues in part: ',
                        '2': 'Supersedes: ',
                        '3': 'Supersedes in part: ',
                        '4': 'Formed by a union of titles including: ',
                        '5': 'Absorbed: ',
                        '6': 'Absorbed in part: ',
                        '7': 'Separated from: ',
                    }[field.indicator1]
                except:
                    preamble = ''
                for subfield in field.get_subfields('a'):
                    notes = add_string(subfield, notes, ' ')
                for subfield in field.get_subfields('b', 'c', 'd', 'g', 'h', 'i', 'k', 'm', 'n', 'o', 'r', 's', 't', 'z'):
                    notes = add_string(subfield, notes, '. ')
                for subfield in field.get_subfields('x'):
                    subfield = re.sub(r'[^0-9xX]', '', subfield)
                    if len(subfield) == 8:
                        notes = add_string('(ISSN: {}-{})'.format(subfield[:4], subfield[4:]), notes, ' ')
                if notes != '':
                    notes = preamble + notes
                    output.values['NN'].add(notes)
                    output.values['S1'].add(notes)

            # 785
            # NN    # Additional information for serials
            # S2    # Succeeding titles
            for field in record.get_fields('785'):
                notes = ''
                try:
                    preamble = {
                        '0': 'Continued by: ',
                        '1': 'Continued in part by: ',
                        '2': 'Superseded by: ',
                        '3': 'Superseded in part by: ',
                        '4': 'Absorbed by: ',
                        '5': 'Absorbed in part by: ',
                        '6': 'Split into: ',
                        '7': 'Merged with/to: ',
                        '8': 'Changed back to: ',
                    }[field.indicator1]
                except:
                    preamble = ''
                for subfield in field.get_subfields('a'):
                    notes = add_string(subfield, notes, ' ')
                for subfield in field.get_subfields('b', 'c', 'd', 'g', 'h', 'i', 'k', 'm', 'n', 'o', 'r', 's', 't', 'z'):
                    notes = add_string(subfield, notes, '. ')
                for subfield in field.get_subfields('x'):
                    subfield = re.sub(r'[^0-9xX]', '', subfield)
                    if len(subfield) == 8:
                        notes = add_string('(ISSN: {}-{})'.format(subfield[:4], subfield[4:]), notes, ' ')
                if notes != '':
                    notes = preamble + notes
                    output.values['NN'].add(notes)
                    output.values['S2'].add(notes)

            # 852 and 979
            # SM    # Shelfmark
            # SD    # DSC Shelfmark
            # SO    # Other Shelfmark
            # BU    # Burney?
            # IO    # India Office?
            for field in record.get_fields('852', '979'):
                shelfmark = ''
                if self.profile == 'N':
                    for subfield in field.get_subfields('b', 'c', 'h', 'j', 'k'):
                        subfield = clean_852(subfield)
                        shelfmark = add_string(subfield, shelfmark, ' ')
                    if 'burney' in shelfmark.lower():
                        output.values['BU'].add('Y')
                    if 'IOL' in shelfmark:
                        output.values['IO'].add('Y')
                else:
                    for subfield in field.get_subfields('h', 'j'):
                        subfield = clean_852(subfield)
                        shelfmark = add_string(subfield, shelfmark, ' ')
                output.values['SM'].add(shelfmark)
                if 'dsc' in ' '.join(field.get_subfields('b')).lower():
                    output.values['SD'].add(shelfmark)
                else:
                    output.values['SO'].add(shelfmark)

            # 856
            # NL    # Link to digitised resource
            for field in record.get_fields('856'):
                for subfield in field.get_subfields('u'):
                    if 'http://www.britishnewspaperarchive.co.uk' in subfield:
                        output.values['NL'].add(subfield)

            # 866
            # HF    # First date held
            # HL    # Last date held
            for field in record.get_fields('866'):
                for subfield in field.get_subfields('a'):
                    subfield = subfield.replace('.', ' ')
                    subfield = re.sub(
                        r'(^newspaper library\s*:?|^newspapers\s*:|[\s\(:\[]*print\s*(copies|issues)?\s*(is|are)?\s*not\s*(made)?\s*available\s*(for\s*conservations?\s*reas[on]*s)?\s*(whe(n|re)\s*(an\s*alternative\s*(format)?\s*(version)?|a\s*micrf?ofilm\s*alternative)\s*exists?)?\.?\]?)',
                        '', subfield, flags=re.IGNORECASE)
                    subfield = re.sub(
                        r'[\s\(see separate record for microfilm holdings|:\[]*(\-*\s*microfilm (is|will be) available (at a later date)?|(see)?\s*all editions microfilm is available for \'kentish express\')[\.:,\)\]]*',
                        '', subfield, flags=re.IGNORECASE)
                    subfield = quick_clean(subfield)
                    if 'micro' not in subfield.lower():
                        output.values['HA'].add(expand_abbreviations(subfield, plurals=False, case=False))
                    subfield = re.sub(
                        r'[\s\(:\[]*(newspapers|newspaper\s*library|(please)?\s*see\s*(also)?\s*sep[ae]rate\s*record\s*for\s*(microfilm|print)\s*holdings|(please)?\s*see\s*(microfilm|print)\s*record|(microfilm|print)\s*holdings\s*only|some\s*issues\s*are\s*held\s*in\s*(microfilm|print)\s*only|(microfilm|print)\s*is\s*available|microfilms?\s*of\s*varied\s*quality\s*with\s*imperfect\s*holdings|for\s*holdings\s*of\s*this\s*title,?\s*(please)?\s*see\s*record|(print|microfilm)\s*for\s*(this\s*title\s*is\s*available\s*on\s*that)\s*for\s*\'[^\']+\'|see\s*\'[^\']+\'\s*(microfilm)?\s*record\s*for\s*(print|microfilm)\s*holdings|for\s*issues\s*(to)?\s*[0-9]{4}\s*(onwards)?,?\s*see\s*(microfilm|print)|ISSN\s*[0-9]{4}\-[0-9]{4})\s*[\.:,\)\]]*',
                        '', subfield, flags=re.IGNORECASE)
                    subfield = re.sub(
                        r'[\s\(:\[]*(all\s*editions\s*microfilm|microfilm\s*will\s*be\s*available\s*at\s*a\s*later\s*date|this\s*record\s*has\s*holdings\s*for\s*both\s*print\s*and\s*microfilm\s*versions|(for)?\s*(earlier|later)?\s*issues\s*(to)?\s*[0-9,\-\s]*\s*(available\s*[io]n|(please)?\s*see)\s*(microfilm|print)\s*(holdings)?)\s*[\.:,\)\]]*',
                        '', subfield, flags=re.IGNORECASE)
                    subfield = re.sub(
                        r'[\s\(:\[]*(n\s*s\s*|(new|original)\s*series|nuova\s*serie|feest\-?nummer|nouvelle\s*s.rie|extra\s*no|no\s*di\s*propaganda|numero\s*(unico|sp.cimen|extraordinario|de\s*reprise)|(centenary\s*souvenir|centennial)\s*number|cyfres\s*newydd|ekstranummer|(pilot|preview|registration)\s*issues?|print|proefnummer|foglio\s*unico|supplement\s*only|sic|see|etc|weekly\s*eds?|extraordinary\s*numbers?)[\.:,\)\]]*',
                        '', subfield, flags=re.IGNORECASE)
                    subfield = re.sub(r'(?<![a-z])(no|yr|year|vol)\s+[0-9\-,\s]+', '', subfield, flags=re.IGNORECASE)
                    subfield = quick_clean(subfield)
                    (first, last) = get_holdings_years(subfield, start_year, end_year)
                    output.values['HF'].add(first)
                    output.values['HL'].add(last)

            # 903, AQN
            # Check whether there is 903 $9 AQN $a L7 (for newspapers not retained)
            # 8F    # 852 holdings flag
            for field in record.get_fields('903'):
                if '9' in field.subfields:
                    output.values['8F'].add('L7')

            for field in record.get_fields('AQN'):
                for subfield in field.get_subfields('a'):
                    if subfield.upper()[:2] == 'L7':
                        output.values['8F'].add('L7')

            if 'L7' not in output.values['8F'] and len(output.values['SM']) > 0:
                output.values['8F'].add('Y')

            # 907, CAT, 920, LEO
            # CL    # Formerly held at Colindale?
            # See also 752
            for field in record.get_fields('907', 'CAT', '920', 'LEO'):
                for subfield in field.get_subfields('a'):
                    if (subfield.upper()[:8] == 'NPL-LOCK' or subfield.upper()[:5] == 'MP35.') \
                            and 'Y' not in output.values['CL']:
                        output.values['CL'].add('Maybe')

            # 932, STA, LDD
            # SX    # Record status
            for field in record.get_fields('932'):
                for subfield in field.get_subfields('a'):
                    subfield = subfield.lower()
                    output.values['SX'].add(subfield)

            for field in record.get_fields('STA', 'LDD'):
                for subfield in field.get_subfields():
                    subfield = subfield.lower().replace('-', '')
                    output.values['SX'].add(subfield)

            # 944, NID
            # ND    # NID (Newspaper ID)
            # NL    # Link to digitised resource
            for field in record.get_fields('944', 'NID'):
                for subfield in field.get_subfields('a'):
                    subfield = re.sub(r'[^0-9]', '', subfield)
                    output.values['ND'].add(subfield)
                    if subfield in self.nid_urls:
                        for item in self.nid_urls[subfield]:
                            output.values['NL'].add(item)

            # Remove null values from output
            for item in output.values:
                if '' in output.values[item]:
                    output.values[item].remove('')
            gc.collect()

            # Material type qualifier
            if self.profile not in ['F', 'M', 'N']:
                if len(output.values['MT']) > 0 and len(mtq) > 0 and 'Online resource' in output.values['MT']:
                    output.values['MT'].remove('Online resource')
                    output.values['MT'].add('Online resource (' + ' ; '.join(mtq) + ')')

            # Dewey classification should not be empty for BNB/CIP records
            if self.profile == 'B' and len(output.values['DW']) == 0:
                output.values['DW'].add('Not yet available')

        return output

    def marc2rf_researcherFormat(self):
        """Convert MARC records to Researcher Format."""
        self.show_header()
        records, names, titles, topics, classification, mfile = None, None, None, None, None, None

        # Check file locations
        marc_folder, marc_file, marc_ext = check_file_location(self.marc_path, 'MARC records', '.lex', True)
        if self.request_path != '':
            request_folder, request_file, request_ext = check_file_location(self.request_path, 'request message', '.msg', True)
        if self.output_folder != '':
            try:
                if not os.path.exists(self.output_folder):
                    os.makedirs(self.output_folder)
            except os.error:
                exit_prompt('Error: Could not create folder for output files')
        if len(self.options) > 1:
            exit_prompt('Error: too many optional parameters specified')

        # --------------------
        # Parameters seem OK => start program
        # --------------------

        # Display confirmation information about the transformation
        print('Input file: {}'.format(marc_file + marc_ext))
        if self.request_path != '':
            print('Request message: {}'.format(request_file + request_ext))
        if self.output_folder != '':
            print('Output folder: {}'.format(self.output_folder))
        if self.debug:
            print('Debug mode')

        # If request message has been specified, use this to determine transformation parameters
        if self.request_path != '':
            print('\nProcessing request file ...')
            print('----------------------------------------')
            print(str(datetime.datetime.now()))
            self.output_fields = Output(initiate=True)
            msgfile = open(os.path.join(request_folder, request_file + request_ext), mode='r', encoding='utf-8',
                           errors='replace')
            for filelineno, line in enumerate(msgfile):
                if 'Coded parameters for your transformation' in line: break
            for filelineno, line in enumerate(msgfile):
                if 'End of coded parameters' in line: break
                if '=' in line:
                    line = clean_msg(line)
                    p, vals = line.split('=', 1)[0].strip(), re.sub(r'^ ', '', line.split('=', 1)[-1])
                    if p != '' and vals != '':
                        if self.debug:
                            try: print('Parameter: {}\nValues: {}\n'.format(p, vals))
                            except: pass
                        if p == 'o':
                            for v in re.sub(r'[^a-zA-Z0-9|]', '.', vals).split('|'):
                                if v in self.output_fields.values:
                                    self.output_fields.values[v] = True
                        elif p == 'v' and (re.sub(r'[^rtnscRTNSC|]', '', vals) == vals):
                            vals = vals.lower()
                            self.file_records = 'r' in vals
                            self.file_titles = 't' in vals
                            self.file_names = 'n' in vals
                            self.file_topics = 's' in vals
                            self.file_classification = 'c' in vals
                        elif p == 's' and (re.sub(r'[^beimBEIM|]', '', vals) == vals):
                            vals = vals.upper()
                            self.sources = set(re.sub(r'[^BEIM]', '', vals.upper()))
                            if 'I' in self.sources:
                                self.output_fields.headings['PD'] = 'Date of creation/publication'
                                self.output_fields.headings['PU'] = 'Date of publication (not standardised)'
            msgfile.close()

        else:
            self.sources.add('M')
            if self.options in ['b', 'c' 'd', 'e' 'f', 'm', 'n']:
                self.profile = self.options.upper()
            else:
                # User selects columns to include
                print('\n')
                print('----------------------------------------')
                print('Select one of the following options: \n')
                print('   D     Default columns \n')
                print('   A     All columns \n')
                print('   B     Columns for BNB/CIP records \n')
                print('   C     Columns for Music records \n')
                print('   E     Columns for ESTC records \n')
                print('   F     FRBRize records \n')
                print('   M     Use MARC fields instead of column headings \n')
                print('   N     Columns for Newspaper records \n')
                print('   R     Columns for RIN project \n')
                print('   S     Select columns to include \n\n')
                self.profile = input('Choose an option:').upper()
                while self.profile not in ['D', 'A', 'B', 'C', 'E', 'F', 'M', 'N', 'R', 'S']:
                    self.profile = input('Sorry, your choice was not recognised. Please enter D, A, B, C, E, F, M, N, R, or S:').upper()

        if self.profile in ['D', 'A', 'B', 'C', 'E', 'F', 'N', 'R', 'S']:
            self.output_fields = Output(profile=self.profile, initiate=True)

        # All columns and Default columns
        if self.profile in ['D', 'A']:
            self.bnb = get_boolean('Does the input data contain BNB records? (Y/N):')
            self.iams = get_boolean('Is the output going to be combined with records from IAMS? (Y/N):')
            self.estc = get_boolean('Does the input data contain ESTC records? (Y/N):')

            # Remove IAMS-specific columns if no IAMS records are to be included
            if self.iams:
                for v in ['AK', 'PV', 'RF']:
                    self.output_fields.values[v] = True
                self.output_fields.headings['PD'] = 'Date of creation/publication'
                self.output_fields.headings['PU'] = 'Date of creation/publication (not standardised)'
            else:
                self.output_fields.values['AK'] = False

            # Remove BNB-specific columns
            for v in ['II', 'VF']:
                self.output_fields.values[v] = False

            # Remove ESTC-specific columns if no ESTC records are present
            if not self.estc:
                self.output_fields.values['ES'] = False

            # Remove other context-specific columns
            for v in ['8F', 'BU', 'CG', 'CL', 'EL', 'FA', 'G1', 'G2', 'HA', 'HF', 'HL', 'IO', 'ND', 'NL',
                      'P1', 'P2', 'PJ', 'SD', 'SO', 'SX']:
                self.output_fields.values[v] = False

        # Columns for ESTC records
        elif self.profile == 'E':
            self.estc, self.main_cat = True, False

        # FRBRization
        elif self.profile == 'F':
            q880 = False

        # Using MARC fields as column headers
        elif self.profile == 'M':
            q880 = False

            # User chooses whether to include subfield codes
            print('\n')
            self.codes = get_boolean('Do you wish to retain subfield codes ($a, $b etc.)? (Y/N):')

        # User selects columns to include
        elif self.profile == 'S':
            self.bnb = get_boolean('Does the input data contain BNB records? (Y/N):')
            self.iams = get_boolean('Is the output going to be combined with records from IAMS? (Y/N):')
            self.estc = get_boolean('Does the input data contain ESTC records? (Y/N):')
            if self.iams:
                self.output_fields.headings['PD'] = 'Date of creation/publication'
                self.output_fields.headings['PU'] = 'Date of creation/publication (not standardised)'

            print('\nChoose the optional columns: \n')

            for v in self.output_fields.values:
                opt_text = self.output_fields.headings[v]
                # Skip some choices depending upon the values of parameters already set
                if (self.iams and v in ['AK', 'PV']) or (self.bnb and v == 'BN') or (self.estc and v == 'ES'):
                    self.output_fields.values[v] = True
                elif v in ['_8F', 'NL', 'P1', 'P2', 'SD', 'SO', 'SX', 'II', 'VF'] or \
                        (not self.output_fields.values['AA'] and v in ['AD', 'AT', 'AR', 'II', 'VF']) or \
                        (not self.output_fields.values['SE'] and v == 'SN') or \
                        (not self.output_fields.values['PD'] and v == 'PU'):
                    self.output_fields.values[v] = False
                else:
                    # Add additional explanatory text to column heading when presenting user with choices
                    if v == 'AA':
                        opt_text += ' (of first author)'
                    elif v in ['FA', 'FC', 'FF', 'HA', 'HF', 'HL', 'IL', 'IS', 'PG', 'TK']:
                        opt_text += ' (relevant to serials only)'
                    elif v in ['BU', 'CG', 'CL', 'IO', 'ND']:
                        opt_text += ' (relevant to newspapers only)'
                    elif v in ['IM', 'MF', 'MG', 'MA']:
                        opt_text += ' (relevant to music only)'
                    elif v in ['SC', 'JK', 'CD', 'CA']:
                        opt_text += ' (relevant to cartographic materials only)'
                    self.output_fields.values[v] = get_boolean('Include {0}? (Y/N):'.format(opt_text))

        if self.profile not in ['F', 'M'] and self.output_fields.values['PD'] and self.output_fields.values['PU']:
            self.output_fields.headings['PD'] += ' (standardised)'

        if self.profile in ['B', 'F', 'M', 'N', 'R']:
            self.file_records = True
        elif self.profile == 'C':
            self.file_records, self.file_titles, self.file_names = True, True, True
        else:
            print('\n----------------------------------------')
            print('Select output files to include:\n')
            self.file_records = get_boolean('Include the Records file? (Y/N):')
            self.file_titles = get_boolean('Include the Titles file? (Y/N):')
            self.file_names = get_boolean('Include the Names file? (Y/N):')
            self.file_topics = get_boolean('Include the Topics file? (Y/N):')
            self.file_classification = get_boolean('Include the Classification file? (Y/N):')

        self.write_readme()

        if self.file_records and self.profile != 'M':
            records_header = '"' + \
                             '","'.join(self.output_fields.headings[v] for v in self.output_fields.values
                                        if self.output_fields.values[v]) + \
                             '"\n'
            if self.profile == 'B': records_name = 'BNB.csv'
            elif self.profile == 'F': records_name = marc_file + '_FRBRized.csv'
            elif self.profile == 'M': records_name = marc_file + '.csv'
            else: records_name = 'records.csv'
            records = open(os.path.join(self.output_folder, records_name), mode='w', encoding='utf-8', errors='replace')
            records.write(records_header)

        if self.file_names:
            names_header = '"' + '","'.join(s for s in ['Name', 'Dates associated with name', 'Type of name', 'Role',
                                                        '","'.join(self.output_fields.headings[v] for v in ['II', 'VF']
                                                                   if self.output_fields.values[v]),
                                                        'Other names',
                                                        '","'.join(self.output_fields.headings[v] for v in self.output_fields.values
                                                                   if self.output_fields.values[v] and v not in ['AA', 'AD', 'AT', 'AR', 'II', 'VF', 'AN'])
                                                        ] if s != '') + '"\n'
            names = open(os.path.join(self.output_folder, 'names.csv'), mode='w', encoding='utf-8', errors='replace')
            names.write(names_header)

        if self.file_titles:
            titles_header = '"' + '","'.join(['Title', 'Other titles',
                                              '","'.join(
                                                  self.output_fields.headings[v] for v in self.output_fields.values
                                                  if self.output_fields.values[v] and v not in ['TT', 'TV', 'TU', 'TK'])
                                              ]) + '"\n'
            titles = open(os.path.join(self.output_folder, 'titles.csv'), mode='w', encoding='utf-8', errors='replace')
            titles.write(titles_header)

        if self.file_topics:
            topics_header = '"' + '","'.join(['Topic', 'Type of topic',
                                              '","'.join(self.output_fields.headings[v] for v in self.output_fields.values
                                                         if self.output_fields.values[v] and v != 'SU')
                                              ]) + '"\n'
            topics = open(os.path.join(self.output_folder, 'topics.csv'), mode='w', encoding='utf-8', errors='replace')
            topics.write(topics_header)

        if self.file_classification:
            classification_header = '"' + '","'.join(['Dewey classification',
                                                      '","'.join(self.output_fields.headings[v] for v in self.output_fields.values
                                                                 if self.output_fields.values[v] and v != 'DW')
                                                      ]) + '"\n'
            classification = open(os.path.join(self.output_folder, 'classification.csv'), mode='w', encoding='utf-8', errors='replace')
            classification.write(classification_header)

        if self.profile == 'M':
            records = open(os.path.join(self.output_folder, marc_file + '.csv'), mode='w', encoding='utf-8', errors='replace')
            # Check which MARC fields are present
            record_count = 0
            print('\nChecking which MARC fields are present ...')
            print('----------------------------------------')
            print(str(datetime.datetime.now()))

            mfile = open(os.path.join(marc_folder, marc_file + marc_ext), 'rb')
            reader = MARCReader(mfile)
            for record in reader:
                record_count += 1
                print('\r{0} MARC records processed'.format(str(record_count)), end='\r')
                for field in record.fields:
                    if field.tag not in self.fields_present and field.tag in marc_fields:
                        self.fields_present[field.tag] = []
            mfile.close()
            records.write('"' + '","'.join(tag for tag in sorted(self.fields_present) if tag != 'STA') + '"\n')
            records.write(
                '"' + '","'.join(marc_fields[tag] for tag in sorted(self.fields_present) if tag != 'STA') + '"\n')
            print('\n')

        if self.profile == 'N':
            # Build index of NID identifiers and URLs linking to digitized resources
            record_count = 0
            print('\nBuilding NID index ...')
            print('----------------------------------------')
            print(str(datetime.datetime.now()))

            mfile = open(os.path.join(marc_folder, marc_file + marc_ext), 'rb')
            reader = MARCReader(mfile)
            for record in reader:
                record_count += 1
                print('\r{0} MARC records processed'.format(str(record_count)), end='\r')

                # 944, NID
                # ND    # NID (Newspaper ID)
                for f1 in record.get_fields('944', 'NID'):
                    for sa in f1.get_subfields('a'):
                        sa = re.sub(r'[^0-9]', '', sa)
                        for f2 in record.get_fields('856'):
                            if sa not in self.nid_urls: self.nid_urls[sa] = set()
                            for su in f2.get_subfields('u'):
                                if 'http://www.britishnewspaperarchive.co.uk' in su:
                                    self.nid_urls[sa].add(su)
            mfile.close()
            print('\n')

        # --------------------
        # Main transformation
        # --------------------

        # Open MARC file
        print('\nStarting transformation ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))

        record_count = 0
        if self.debug:
            print('Opening file: {}'.format(str(os.path.join(marc_folder, marc_file + marc_ext))))
        mfile = open(os.path.join(marc_folder, marc_file + marc_ext), 'rb')
        reader = MARCReader(mfile)
        for record in reader:
            record_count += 1
            print('\r{0} MARC records processed'.format(str(record_count)), end='\r')
            output = self.convert_record(record)

            # Write record to output file

            if self.profile == 'F':
                records.write('"' + '","'.join((' ; '.join(str(p) for p in output.values[tag]).strip())
                                               for tag in output.values) + '"\n')

            elif self.profile == 'M':
                if not (any(s in ''.join(output.values['STA']).lower() for s in
                            ['deleted', 'suppressed', 'prepublication'])) and len(output.values['001']) > 0:
                    records.write('"' + '","'.join((' ; '.join(str(p) for p in output.values[tag]).strip())
                                                   for tag in sorted(output.values) if tag != 'STA') + '"\n')

            elif self.profile == 'N':
                # Limit newspaper selection to UK, Ireland and current UK dependencies
                # Records must have shelfmarks and not have 'L7' in field AQN $a or 903 $9 (indicated by 8F)
                if not (any(s in ''.join(output.values['SX']).lower() for s in
                            ['deleted', 'suppressed', 'prepublication'])) \
                        and len(output.values['ID']) > 0 and 'Y' in output.values['8F'] \
                        and (any(s in ' '.join(output.values['PC']).lower() for s in
                                 ['akrotiri', 'alderney', 'anguilla', 'ascension', 'bermuda', 'cayman',
                                  'channel island', 'dhekelia', 'falkland', 'gibraltar', 'guernsey', 'isle of man',
                                  'montserrat', 'pitcairn', 'saint helena', 'sark', 'south georgia', 'south sandwich',
                                  'tristan da cunha', 'turks and caicos', 'britain', 'british', 'united kingdom',
                                  'england', 'wales', 'scotland', 'ireland']) or len(output.values['PC']) == 0):
                    # Material type (column label Carrier type) must be empty for Newspaper records. 2016-07-06
                    output.values['MT'] = []
                    # Delimiter for Newspaper records is | but for all other outputs is ;
                    output_string = '"'
                    for v in self.output_fields.values:
                        if v in output.values and self.output_fields.values[v]:
                            try: output_string += '|'.join(p for p in sorted(output.values[v]) if p != '') + '","'
                            except: print('\nError in newspaper records: {}\n{}\n'.format(v, str(sys.exc_info())))
                    output_string += '\n'
                    output_string = output_string.replace(',"\n', '\n')
                    records.write(output_string)
                    gc.collect()

            else:

                if not (any(s in ''.join(output.values['SX']).lower() for s in
                            ['deleted', 'suppressed', 'prepublication'])) \
                        and len(output.values['ID']) > 0 \
                        and not (len(''.join(output.values['TT'])) <= 5 and len(output.values['AA']) == 0 and len(output.values['PD']) == 0):

                    if self.file_records:
                        output_string = '"'
                        for v in self.output_fields.values:
                            if v in output.values and self.output_fields.values[v]:
                                if v == 'AN':
                                    s = ''
                                    for item in output.values['AN']:
                                        if item[0] != '':
                                            name = item[0]
                                            for i in [1, 3, 4, 5]:
                                                if item[i] != '': name += ', ' + item[i]
                                            if item[2] != '': name = name + ' [' + item[2] + ']'
                                            s = add_string(name, s, ' ; ')
                                    output_string += s + '","'
                                elif v == 'SU':
                                    s = ''
                                    for item in output.values['SU']:
                                        if item[0] != '': topic = str(item[0])
                                        s = add_string(topic, s, ' ; ')
                                    output_string += s + '","'
                                elif v == 'TV':
                                    try:
                                        output_string += ' ; '.join(p for p in sorted(output.values['TV'])
                                                                    if (p != '' and p not in output.values['TT'])) + '","'
                                    except:
                                        print('\nError [please report code 003 to VM]: {}\n'.format(str(sys.exc_info())))
                                else:
                                    try:
                                        output_string += ' ; '.join(p for p in sorted(output.values[v]) if p != '') + '","'
                                    except:
                                        print('\nError in records: {}\n{}\n'.format(v, str(sys.exc_info())))

                        output_string += '\n'
                        output_string = output_string.replace(',"\n', '\n')
                        records.write(output_string)
                        gc.collect()

                    if self.file_names:
                        for item in output.values['AN']:
                            if item[0] != '':
                                output_string = '"'
                                output_string += item[0] + '","'  # Name
                                if item[1] != '': output_string += item[1]  # Dates associated with name
                                output_string += '","'
                                if item[2] != '': output_string += item[2]  # Type of name
                                output_string += '","'
                                if item[3] != '': output_string += item[3]  # Name role
                                output_string += '","'
                                if self.bnb and self.output_fields.values['II']:
                                    if item[4] != '': output_string += item[4]  # ISNI
                                    output_string += '","'
                                if self.bnb and self.output_fields.values['VF']:
                                    if item[5] != '': output_string += item[5]  # VIAF
                                    output_string += '","'
                                s = ''
                                for item2 in output.values['AN']:
                                    if item2[0] != '':
                                        name = item2[0]
                                        for i in [1, 3, 4, 5]:
                                            if item2[i] != '': name += ', ' + item2[i]
                                        if item2[2] != '': name = name + ' [' + item2[2] + ']'
                                        s = add_string(name, s, ' ; ')
                                output_string += s + '","'
                                for v in self.output_fields.values:
                                    if v in output.values and self.output_fields.values[v]:
                                        if v == 'SU':
                                            s = ''
                                            for item3 in output.values['SU']:
                                                if item3[0] != '': topic = str(item3[0])
                                                s = add_string(topic, s, ' ; ')
                                            output_string += s + '","'
                                        elif v == 'TV':
                                            try:
                                                output_string += ' ; '.join(p for p in sorted(output.values['TV'])
                                                                            if (p != '' and p not in output.values['TT'])) + '","'
                                            except:
                                                print('\nError [please report code 004 to VM]: {}\n'.format(str(sys.exc_info())))
                                        elif v not in ['AN', 'AA', 'AD', 'AT', 'AR', 'II', 'VF']:
                                            try:
                                                output_string += ' ; '.join(p for p in sorted(output.values[v]) if p != '') + '","'
                                            except:
                                                print('\nError in names: {}\n{}\n'.format(v, str(sys.exc_info())))
                                output_string += '\n'
                                output_string = output_string.replace(',"\n', '\n')
                                names.write(output_string)

                    if self.file_titles:
                        for item in output.values['TV']:
                            output_string = '"' + item + '","'
                            output_string += ' ; '.join(p for p in sorted(output.values['TV']) if p != '' and p != item) + '","'
                            for v in self.output_fields.values:
                                if v in output.values and self.output_fields.values[v]:
                                    if v == 'AN':
                                        s = ''
                                        for item2 in output.values['AN']:
                                            if item2[0] != '':
                                                name = item2[0]
                                                for i in [1, 3, 4, 5]:
                                                    if item2[i] != '': name += ', ' + item2[i]
                                                if item2[2] != '': name = name + ' [' + item2[2] + ']'
                                                s = add_string(name, s, ' ; ')
                                        output_string += s + '","'
                                    elif v == 'SU':
                                        s = ''
                                        for item3 in output.values['SU']:
                                            if item3[0] != '': topic = str(item3[0])
                                            s = add_string(topic, s, ' ; ')
                                        output_string += s + '","'
                                    elif v not in ['TK', 'TT', 'TU', 'TV']:
                                        try:
                                            output_string += ' ; '.join(p for p in sorted(output.values[v]) if p != '') + '","'
                                        except:
                                            print('\nError in titles: {}\n{}\n'.format(v, str(sys.exc_info())))
                            output_string += '\n'
                            output_string = output_string.replace(',"\n', '\n')
                            titles.write(output_string)

                    if self.file_topics:
                        for item in output.values['SU']:
                            if item[0] != '':
                                output_string = '"' + item[0] + '","'
                                if item[1] != '': output_string += item[1]
                                output_string += '","'
                                for v in self.output_fields.values:
                                    if v in output.values and self.output_fields.values[v]:
                                        if v == 'AN':
                                            s = ''
                                            for item2 in output.values['AN']:
                                                if item2[0] != '':
                                                    name = item2[0]
                                                    for i in [1, 3, 4, 5]:
                                                        if item2[i] != '': name += ', ' + item2[i]
                                                    if item2[2] != '': name = name + ' [' + item2[2] + ']'
                                                    s = add_string(name, s, ' ; ')
                                            output_string += s + '","'
                                        elif v == 'TV':
                                            try:
                                                output_string += ' ; '.join(p for p in sorted(output.values['TV'])
                                                                            if (p != '' and p not in output.values['TT'])) + '","'
                                            except:
                                                print('\nError [please report code 005 to VM]: {}\n'.format(str(sys.exc_info())))
                                        elif v not in ['SU']:
                                            try:
                                                output_string += ' ; '.join(p for p in sorted(output.values[v]) if p != '') + '","'
                                            except:
                                                print('\nError in topics: {}\n{}\n'.format(v, str(sys.exc_info())))
                                output_string += '\n'
                                output_string = output_string.replace(',"\n', '\n')
                                topics.write(output_string)

                    if self.file_classification:
                        for item in output.values['DW']:
                            if item != '':
                                output_string = '"' + str(item) + '","'
                                for v in self.output_fields.values:
                                    if v in output.values and self.output_fields.values[v]:
                                        if v == 'AN':
                                            s = ''
                                            for item2 in output.values['AN']:
                                                if item2[0] != '':
                                                    name = item2[0]
                                                    for i in [1, 3, 4, 5]:
                                                        if item2[i] != '': name += ', ' + item2[i]
                                                    if item2[2] != '': name = name + ' [' + item2[2] + ']'
                                                    s = add_string(name, s, ' ; ')
                                            output_string += s + '","'
                                        elif v == 'SU':
                                            s = ''
                                            for item3 in output.values['SU']:
                                                if item3[0] != '': topic = str(item3[0])
                                                s = add_string(topic, s, ' ; ')
                                            output_string += s + '","'
                                        elif v == 'TV':
                                            try:
                                                output_string += ' ; '.join(p for p in sorted(output.values['TV'])
                                                                            if (p != '' and p not in output.values['TT'])) + '","'
                                            except:
                                                print('\nError [please report code 006 to VM]: {}\n'.format(str(sys.exc_info())))
                                        elif v not in ['DW']:
                                            try:
                                                output_string += ' ; '.join(p for p in sorted(output.values[v]) if p != '') + '","'
                                            except:
                                                print('\nError in classification: {}\n{}\n'.format(v, str(sys.exc_info())))
                                output_string += '\n'
                                output_string = output_string.replace(',"\n', '\n')
                                classification.write(output_string)

                    gc.collect()

        # Close files
        for file in [records, names, titles, topics, classification, mfile]:
            try: file.close()
            except: pass


class ConfigWriter(object):
    """A class for writing config files.

    :param debug: Display additional output to assist with debugging.
    """

    def __init__(self, debug=False):
        self.debug = debug
        self.header = '========================================\n' + \
                      'write_rf_config\n' + \
                      'MARC record selection for Researcher Format\n' + \
                      '========================================\n' + \
                      'This utility prepare config files for selection of MARC records\n' + \
                      'to convert to Researcher Format\n'
        self.parameters = {
            'cp1':  set(),     # Publication country codes
            'd1':   '',        # Start date
            'd2':   '',        # End date
            'dw':   set(),     # Dewey range
            'l1':   set(),     # Language codes
            'or1':  set(),     # Optional requirements (007 position 00 codes)
            'or2':  set(),     # Optional requirements (LDR position 06 codes)
            'os1':  False,     # Include only resources with a BNB number
            'os2':  False,     # Include only resources with a British Library shelfmark
            's':    set(),     # Sources
            'txt':  set(),     # Search strings
            }

    def show_header(self):
        if self.header:
            print(self.header)

    def marc2rf_write_rf_config(self, request_path, output_folder):
        """Prepare config files for selection of MARC records to convert
        to Researcher Format.

        :param request_path: Path to Outlook message containing details of the request.
        :param output_folder: Folder to save config files.
        """
        self.show_header()

        # Check file locations
        request_folder, request_file, request_ext = check_file_location(request_path, 'request message', '.msg', True)
        if output_folder != '':
            try:
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)
            except os.error:
                exit_prompt('Error: Could not create folder for output files')

        # --------------------
        # Parameters seem OK => start program
        # --------------------

        # Display confirmation information about the transformation
        print('Request message: {}'.format(request_file + request_ext))
        if output_folder != '':
            print('Output folder: {}'.format(output_folder))
        if self.debug:
            print('Debug mode')

        # Process input file
        print('\nProcessing request file ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))

        msgfile = open(os.path.join(request_folder, request_file + request_ext), mode='r', encoding='utf-8',
                       errors='replace')
        for filelineno, line in enumerate(msgfile):
            if 'Coded parameters for your transformation' in line: break
        for filelineno, line in enumerate(msgfile):
            if 'End of coded parameters' in line: break
            if '=' in line:
                line = clean_msg(line)
                p, vals = line.split('=', 1)[0].strip(), re.sub(r'^ ', '', line.split('=', 1)[-1])
                if p != '' and vals != '':
                    if self.debug:
                        try: print('Parameter: {}\nValues: {}\n'.format(p, vals))
                        except: pass
                    if p == 'txt':
                        for v in re.sub(r'[^\x00-\x7F]', '.', vals).split('|'):
                            self.parameters[p].add(v)
                    elif p in ['cp1', 'l1']:
                        for v in re.sub(r'[^a-z|\s]', '', vals.lower()).split('|'):
                            if 2 <= len(v) <= 3:
                                self.parameters[p].add((v + ' ')[:3])
                    elif p == 'dw':
                        for v in re.sub(r'[^0-9.\-;]', '', vals).split(';'):
                            self.parameters['dw'].add(v)
                    elif p in ['d1', 'd2'] and len(re.sub(r'[^0-9]', '', vals)) >= 4:
                        self.parameters[p] = re.sub(r'[^0-9]', '', vals)[:4]
                    elif p in ['os1', 'os2']:
                        if re.sub(r'[^a-z]', '', vals.lower()) == 'on':
                            self.parameters[p] = True
                    elif p == 's':
                        self.parameters[p] = set(re.sub(r'[^a-zA-Z]', '', vals))
                    elif p == 'or':
                        if 'u' in vals:
                            self.parameters['or2'] = set('cdj')
                        self.parameters['or1'] = set(re.sub(r'[^temidv]', '', vals))
        msgfile.close()

        if 'I' in self.parameters['s']:
            self.parameters['s'].discard('I')
            # IAMS config file selectIAMS.cfg is a flag to indicate that IAMS selection is required
            with open('selectIAMS.cfg', mode='w', encoding='utf-8', errors='replace') as f:
                f.write('IAMS selection required')

        selection_criteria = ['', '']

        # Dates
        if self.parameters['d1'] != '' and self.parameters['d2'] == '':
            self.parameters['d2'] = '9999'
        if self.parameters['d2'] != '' and self.parameters['d1'] == '':
            self.parameters['d1'] = '0000'
        if self.parameters['d1'] != '' and self.parameters['d2'] != '' \
                and is_number(self.parameters['d1']) and is_number(self.parameters['d2']) \
                and float(self.parameters['d1']) <= float(self.parameters['d2']):
            output_string = ''
            output_string = add_string(
                '(FIELD 008 POSITION 7 EQUALS \"[eprst]\") AND (FIELD 008 POSITION 8-11 IN_RANGE \"{}-{}\")'.format(
                    self.parameters['d1'], self.parameters['d2']), output_string, '\n OR ', brackets=True)
            output_string = add_string(
                '(FIELD 008 POSITION 7 EQUALS \"[cdikmqu]\") AND (FIELD 008 POSITION 8-11 LESS_THAN \"{}\") AND (FIELD 008 POSITION 12-15 GREATER_THAN \"{}\")'.format(
                    self.parameters['d2'], self.parameters['d1']), output_string, '\n OR ', brackets=True)
            selection_criteria[0] = add_string(output_string, selection_criteria[0], '\nAND\n', brackets=True)

        # Language
        if self.parameters['l1']:
            selection_criteria[0] = add_string('\n OR '.join('( FIELD 008 POSITION 36-38 CONTAINS "{}" )'.format(c)
                                                             for c in self.parameters['l1']),
                                               selection_criteria[0], '\nAND\n', brackets=True)

        # Country of publication
        if self.parameters['cp1']:
            selection_criteria[0] = add_string('\n OR '.join('( FIELD 008 POSITION 16-18 CONTAINS "{}" )'.format(c)
                                                             for c in self.parameters['cp1']),
                                               selection_criteria[0], '\nAND\n', brackets=True)

        # BNB number
        if self.parameters['os1']:
            selection_criteria[0] = add_string('(FIELD 015 EXISTS)\nAND\n(FIELD 015 SUBFIELD 2 EQUALS "bnb")',
                                               selection_criteria[0], '\nAND\n')

        # BL shelfmark
        if self.parameters['os2']:
            selection_criteria[0] = add_string('(FIELD 852 SUBFIELD h EXISTS) OR (FIELD 852 SUBFIELD j EXISTS)',
                                               selection_criteria[0], '\nAND\n', brackets=True)

        # Types of resource
        if self.parameters['or1']:
            selection_criteria[0] = add_string('\n OR '.join('(FIELD 007 POSITION 1 CONTAINS "{}")'.format(c)
                                                             for c in self.parameters['or1']),
                                               selection_criteria[0], '\nAND\n', brackets=True)
        if self.parameters['or2']:
            selection_criteria[0] = add_string('\n OR '.join('(FIELD 000 POSITION 7 CONTAINS "{}")'.format(c)
                                                             for c in self.parameters['or2']),
                                               selection_criteria[0], '\nAND\n', brackets=True)

        # Dewey
        if self.parameters['dw']:
            output_string = ''
            for class_range in self.parameters['dw']:
                if '-' in class_range:
                    # Save start and end of range
                    start_dewey, end_dewey = class_range.split('-', 1)
                    start_dewey = normalize_dewey(start_dewey)
                    end_dewey = normalize_dewey(end_dewey)
                    if self.debug:
                        print('Start of Dewey range: {}'.format(start_dewey))
                        print('End of Dewey range: {}'.format(end_dewey))
                    # Check end of range exceeds start of range
                    if start_dewey != '' and end_dewey != '':
                        j = 0
                        if float(end_dewey) >= float(start_dewey):
                            # Find first position where start and end ranges differ
                            # Count starts from 0 and includes decimal point
                            if start_dewey[0] == end_dewey[0]:
                                while j in range(len(start_dewey)) and start_dewey[j] == end_dewey[j]:
                                    j += 1
                            if self.debug:
                                print('Start and end ranges differ at position {}'.format(str(j)))

                            k = len(start_dewey) - 1
                            while k > j:
                                if self.debug: print('k: {}'.format(str(k)))
                                if k != 3:
                                    if k == len(start_dewey) - 1:
                                        last_digit = start_dewey[k]
                                    elif k in range(len(start_dewey)):
                                        last_digit = str(int(start_dewey[k]) + 1)
                                    else:
                                        last_digit = '0'

                                    number = '{}[{}-9]'.format(start_dewey[:k], last_digit).replace('[9-9]', '9')
                                    if 0 <= k <= 1:
                                        number += '[0-9]' * (2 - k)
                                    number = normalize_dewey(number, escapes=True)

                                    if number != '' and last_digit != '10':
                                        output_string = add_string(
                                            '(FIELD 082 SUBFIELD a CONTAINS "^{}")'.format(number), output_string, '\n OR ')
                                k -= 1

                            if k == j:
                                if self.debug: print('k: {}'.format(str(k)))
                                if k != 3:
                                    number = ''
                                    if k == len(start_dewey) - 1:
                                        number = re.sub(r'\[(\d)-\1\]', '\1', '{}[{}-{}]'.format(
                                            start_dewey[:k], str(int(start_dewey[k])), str(int(end_dewey[k]))))
                                    elif int(start_dewey[k]) + 1 <= int(end_dewey[k]) - 1:
                                        number = re.sub(r'\[(\d)-\1\]', '\1', '{}[{}-{}]'.format(
                                            start_dewey[:k], str(int(start_dewey[k]) + 1), str(int(end_dewey[k]) - 1)))

                                    if 0 <= k <= 1:
                                        number += '[0-9]' * (2 - k)
                                    number = normalize_dewey(number, escapes=True)

                                    if number != '':
                                        output_string = add_string(
                                            '(FIELD 082 SUBFIELD a CONTAINS "^{}")'.format(number), output_string, '\n OR ')

                            k = len(end_dewey) - 1
                            while k > j:
                                if self.debug: print('k: {}'.format(str(k)))
                                if k != 3:
                                    if k == len(end_dewey) - 1:
                                        last_digit = end_dewey[k]
                                    elif k in range(len(end_dewey)):
                                        last_digit = str(int(end_dewey[k]) - 1)
                                    else:
                                        last_digit = '9'

                                    number = '{}[0-{}]'.format(end_dewey[:k], last_digit).replace('[0-0]', '0')
                                    if 0 <= k <= 1:
                                        number += '[0-9]' * (2 - k)
                                    number = normalize_dewey(number, escapes=True)

                                    if number != '' and last_digit != '-1':
                                        output_string = add_string(
                                            '(FIELD 082 SUBFIELD a CONTAINS "^{}")'.format(number), output_string, '\n OR ')
                                k -= 1
                else:
                    class_range = normalize_dewey(class_range, escapes=True)
                    if class_range != '':
                        output_string = add_string('(FIELD 082 SUBFIELD a CONTAINS "^{}")'.format(class_range), output_string, '\n OR ')

            selection_criteria[0] = add_string(output_string, selection_criteria[0], '\nAND\n', brackets=True)

        # Search strings
        if self.parameters['txt']:
            output_string = ''
            for search_string in self.parameters['txt']:
                search_string = search_string.strip()
                if search_string != '' and not is_IAMS_id(search_string):
                    if '$' in search_string:
                        temp = ''
                        subfields = search_string.split('$')
                        for j in range(len(subfields)):
                            subfields[j] = escape_regex_chars(clean_search_string(subfields[j]))
                            if len(subfields[j]) >= 2:
                                temp = add_string('(FIELD 100-499 SUBFIELD {} CONTAINS CASE_INSENSITIVE "{}")'.format(
                                    subfields[j][0], clean_search_string(subfields[j][1:])), temp, ' AND ')
                        output_string = add_string(temp, output_string, '\n OR ', brackets=True)
                        output_string = add_string(temp.replace('FIELD 100-499 SUBFIELD', 'FIELD 600-799 SUBFIELD'), output_string, '\n OR ', brackets=True)
                    else:
                        search_string = escape_regex_chars(clean_search_string(search_string))
                        output_string = add_string(
                            '(FIELD 100-499 SUBFIELD. CONTAINS CASE_INSENSITIVE "{}")'.format(search_string), output_string, '\n OR ')
                        output_string = add_string(
                            '(FIELD 600-799 SUBFIELD. CONTAINS CASE_INSENSITIVE "{}")'.format(search_string), output_string, '\n OR ')

            selection_criteria[1] = add_string(output_string, selection_criteria[1], '\nAND\n', brackets=True)

        for s in self.parameters['s']:
            with open(os.path.join(output_folder, 'select{}1.cfg'.format(SOURCES[s])), mode='w', encoding='utf-8', errors='replace') as f:
                f.write(selection_criteria[0])
            with open(os.path.join(output_folder, 'select{}2.cfg'.format(SOURCES[s])), mode='w', encoding='utf-8', errors='replace') as f:
                f.write(selection_criteria[1])

        # Delete empty files
        for file in glob.glob('*.cfg'):
            try:
                if os.stat(file).st_size == 0:
                    os.remove(file)
            except: pass
