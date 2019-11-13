#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Main module for iams2rf."""

# Import required modules
# These should all be contained in the standard library
from collections import OrderedDict
import datetime
import gc
import locale
import os
import re
import sqlite3
import sys
import unicodedata

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'

# Set locale to assist with sorting
locale.setlocale(locale.LC_ALL, '')

# Set threshold for garbage collection (helps prevent the program run out of memory)
gc.set_threshold(400, 5, 5)

# ====================
#  Regular expressions
# ====================

RE_IAMS_ID = re.compile('0[34][0-9]-[0-9]{9}')

REGEXES = {
    'rel_name': re.compile(
        '<RelatedArchiveDescriptionNamedAuthority TargetNumber=["\']+([0-9]{3}-[0-9]{9})["\']+>' +
        '<RelationshipType>(.*?)</RelationshipType>.*?</RelatedArchiveDescriptionNamedAuthority>'),
    'rel_auth': re.compile('<RelatedArchiveDescription(Place|Subject) TargetNumber=["\']+([0-9]{3}-[0-9]{9})["\']+>'),
    'rel_arch': re.compile(
        '<RelatedArchDesc RelationshipNumber=["\']+([0-9]{3}-[0-9]{9})["\']+>' +
        '<NatureRelationshipTargetDescription>(.*?)</NatureRelationshipTargetDescription>' +
        '<NatureRelationshipSourceDescription>(.*?)</NatureRelationshipSourceDescription>.*?</RelatedArchDesc>'),
    'c_name': re.compile(
        '<CorporationName>.*?<CorporateName\s*[^>]*>(.*?)</CorporateName>.*?' +
        '<NameType>Authorised</NameType>.*?</CorporationName>'),
    'c_qualifiers': re.compile(
        '<CorporationName>.*?<AdditionalQualifiers\s*[^>]*>(.*?)</AdditionalQualifiers>.*?' +
        '<NameType>Authorised</NameType>.*?</CorporationName>'),
    'c_jurisdiction': re.compile(
        '<CorporationName>.*?<Jurisdiction\s*[^>]*>(.*?)</Jurisdiction>.*?' +
        '<NameType>Authorised</NameType>.*?</CorporationName>'),
    'c_dates': re.compile(
        '<CorporationName>.*?<DateRange\s*[^>]*>(.*?)</DateRange>.*?' +
        '<NameType>Authorised</NameType>.*?</CorporationName>'),
    'f_surname': re.compile(
        '<FamilyName>.*?<FamilySurname\s*[^>]*>(.*?)</FamilySurname>.*?<NameType>Authorised</NameType>.*?</FamilyName>'),
    'f_epithet': re.compile(
        '<FamilyName>.*?<FamilyEpithet\s*[^>]*>(.*?)</FamilyEpithet>.*?<NameType>Authorised</NameType>.*?</FamilyName>'),
    'f_dates': re.compile(
        '<FamilyName>.*?<DateRange\s*[^>]*>(.*?)</DateRange>.*?<NameType>Authorised</NameType>.*?</FamilyName>'),
    'p_surname': re.compile(
        '<PersonName>.*?<Surname\s*[^>]*>(.*?)</Surname>.*?<NameType>Authorised</NameType>.*?</PersonName>'),
    'p_forename': re.compile(
        '<PersonName>.*?<FirstName\s*[^>]*>(.*?)</FirstName>.*?<NameType>Authorised</NameType>.*?</PersonName>'),
    'p_title': re.compile(
        '<PersonName>.*?<Title\s*[^>]*>(.*?)</Title>.*?<NameType>Authorised</NameType>.*?</PersonName>'),
    'p_epithet': re.compile(
        '<PersonName>.*?<Epithet\s*[^>]*>(.*?)</Epithet>.*?<NameType>Authorised</NameType>.*?</PersonName>'),
    'p_dates': re.compile(
        '<PersonName>.*?<DateRange\s*[^>]*>(.*?)</DateRange>.*?<NameType>Authorised</NameType>.*?</PersonName>'),
    'pl_name': re.compile('<Name>(.*?)</Name>'),
    'pl_localUnit': re.compile('<LocalAdminUnit\s*[^>]*>(.*?)</LocalAdminUnit>'),
    'pl_widerUnit': re.compile('<WiderAdminUnit\s*[^>]*>(.*?)</WiderAdminUnit>'),
    'pl_country': re.compile('<Country\s*[^>]*>(.*?)</Country>'),
    's_text': re.compile('<Entry\s*[^>]*>(.*?)</Entry>'),
    's_type': re.compile('<Type\s*[^>]*>(.*?)</Type>'),
    'b_reference': re.compile('<Reference\s*[^>]*>(.*?)</Reference>'),
    'b_title': re.compile('<Title\s*[^>]*>(.*?)</Title>'),
    'AN': re.compile(
        '<RelatedArchiveDescriptionNamedAuthority TargetNumber=["\']([0-9]{3}-[0-9]{9})["\']>\s*' +
        '<RelationshipType>(.*?)</RelationshipType>'),
    'LA': re.compile('<MaterialLanguage\s+[^>]*>(.*?)</MaterialLanguage>'),
    'S_LANGUAGES': re.compile('<MaterialLanguage .*?LanguageIsoCode=["\']([a-z]+)["\']>'),
    'SU': re.compile('<RelatedArchiveDescription(?:Place|Subject) TargetNumber=["\']([0-9]{3}-[0-9]{9})["\']>'),
    'TV': re.compile('<AdditionalTitle>[^<>]*?<A?Title>(.*?)<\/A?Title>.*?<\/AdditionalTitle>'),
    'OI': re.compile('<ExternalIdentifier>[^<>]*?<Value>(.*?)<\/Value>.*?<Type .*?>(.*?)<\/Type>.*?<\/ExternalIdentifier>'),
    'isni': re.compile('<ExternalIdentifier>[^<>]*?<Value>(.*?)<\/Value>.*?<Type .*?>[^<>]*?ISNI[^<>]*?<\/Type>.*?<\/ExternalIdentifier>'),
    'viaf': re.compile('<ExternalIdentifier>[^<>]*?<Value>(.*?)<\/Value>.*?<Type .*?>[^<>]*?VIAF[^<>]*?<\/Type>.*?<\/ExternalIdentifier>'),

}

# ====================
#    Lookup tables
# ====================

STATUS = {
    '1': 'Draft',
    '3': 'Pending Deletion',
    '4': 'Published',
    '5': 'Deleted',
    '6': 'Loaded',
    '7': 'Approved',
    '8': 'Ready for Review',
    '9': 'Rejected',
}

TYPES = OrderedDict({
    '032': 'Fonds',
    '033': 'SubFonds',
    '034': 'SubSubFonds',
    '035': 'SubSubSubFonds',
    '036': 'Series',
    '037': 'SubSeries',
    '038': 'SubSubSeries',
    '039': 'SubSubSubSeries',
    '040': 'File',
    '041': 'Item',
    '042': 'SubItem',
    '043': 'SubSubItem',
    '044': 'SubSubSubItem',
    '045': 'Corporation',
    '046': 'Family',
    '047': 'Person',
    '048': 'Place',
    '049': 'Subject',
})


# ====================
#   Global variables
# ====================

TABLE_DEFINITIONS = {
    'records': None,
    'subjects': ([
        ('Topic', 'NTEXT'),
        ('TopicType', 'NTEXT')
    ]),
    'names': ([
        ('Name', 'NTEXT'),
        ('NameDates', 'NTEXT'),
        ('NameType', 'NTEXT'),
        ('NameRole', 'NTEXT'),
        ('NameISNI', 'NTEXT'),
        ('NameVIAF', 'NTEXT')
    ]),
    'titles': ([
        ('Title', 'NTEXT')
    ]),
}

# ====================
#       Classes
# ====================


class Output:
    def __init__(self):
        self.values = OrderedDict([
            ('S_LANGUAGES', ['', 'COMPLEX']),
            ('S_DATE1', ['', '||StartDate']),
            ('S_DATE2', ['', '||EndDate']),
            ('_ID', ['N==BL record ID', 'COMPLEX']),  # DIFFERS FROM STANDARD RF
            ('RT', ['N==Type of resource', 'COMPLEX']),
            ('CT', ['N==Content type', '']),
            ('MT', ['N==Material type', '']),
            ('BN', ['N==BNB number', '']),
            ('LC', ['N==LC number', 'COMPLEX']),
            ('OC', ['N==OCLC number', '']),
            ('ES', ['N==ESTC citation number', '']),
            ('AK', ['N==Archival Resource Key', '||MDARK']),
            ('IB', ['N==ISBN', '']),
            ('_IS', ['N==ISSN', '']),  # DIFFERS FROM STANDARD RF
            ('IL', ['N==ISSN-L', '']),
            ('IM', ['N==International Standard Music Number (ISMN)', '']),
            ('IR', ['N==International Standard Recording Code (ISRC)', '']),
            ('IA', ['N==International Article Number (EAN)', '']),
            ('PN', ['N==Publisher number', '']),
            ('OI', ['N==Other identifier', 'COMPLEX']),
            ('AA', ['N==Name', 'COMPLEX']),
            ('AD', ['N==Dates associated with name', 'COMPLEX']),
            ('AT', ['N==Type of name', 'COMPLEX']),
            ('AR', ['N==Role', 'COMPLEX']),
            ('II', ['N==ISNI', 'COMPLEX']),
            ('VF', ['N==VIAF', 'COMPLEX']),
            ('AN', ['N==All names', 'COMPLEX']),
            ('TT', ['Y==Title', '||Title']),
            ('TU', ['N==Uniform title', '']),
            ('TK', ['N==Key title', '']),
            ('TV', ['N==Variant titles', 'COMPLEX']),
            ('S1', ['N==Preceding titles', '']),
            ('S2', ['N==Succeeding titles', '']),
            ('SE', ['N==Series title', '']),
            ('SN', ['N==Number within series', '']),
            ('PC', ['N==Country of publication', '']),
            ('PP', ['N==Place of creation/publication', '||PlaceOfOrigin']), # New field with IAMS release 26
            ('PB', ['N==Publisher', '']),
            ('PD', ['N==Date of creation/publication', '||DateRange']),
            ('PU', ['N==Date of creation/publication (not standardised)', '||DateRange']),
            ('PJ', ['N==Projected date of publication', '']),
            ('PG', ['N==Publication date range', '||DateRange']),
            ('P1', ['N==Publication date one', '']),
            ('P2', ['N==Publication date two', '']),
            ('FA', ['N==Free text information about dates of publication', '']),
            ('HF', ['N==First date held', '']),
            ('HL', ['N==Last date held', '']),
            ('HA', ['N==Free text information about holdings', '']),
            ('FC', ['N==Current publication frequency', '']),
            ('FF', ['N==Former publication frequency', '']),
            ('ED', ['N==Edition', '']),
            ('DS', ['N==Physical description', '||Extent|PhysicalCharacteristics']),
            ('SC', ['N==Scale', '||Scale|ScaleDesignator']),
            ('JK', ['N==Projection', '||Projection']),
            ('CD', ['N==Coordinates', '||DecimalCoordinates|DegreeCoordinates']),
            ('MF', ['N==Musical form', '']),
            ('MG', ['N==Musical format', '']),
            ('PR', ['N==Price', '']),
            ('DW', ['N==Dewey classification', '']),
            ('LN', ['N==Library of Congress classification', '']),
            ('SM', ['N==BL shelfmark', 'COMPLEX']),
            ('SD', ['N==DSC shelfmark', '']),
            ('SO', ['N==Other shelfmark', '']),
            ('BU', ['N==Burney?', '']),
            ('IO', ['N==India Office?', '']),
            ('CL', ['N==Formerly held at Colindale?', '']),
            ('SU', ['N==Topics', 'COMPLEX']),
            ('G1', ['N==First geographical subject heading', 'COMPLEX']),
            ('G2', ['N==Subsequent geographical subject headings', 'COMPLEX']),
            ('CG', ['N==General area of coverage', '']),
            ('CC', ['N==Coverage: Country', '']),
            ('CF', ['N==Coverage: Region', '']),
            ('CY', ['N==Coverage: City', '']),
            ('GE', ['N==Genre', '']),
            ('TA', ['N==Target audience', '']),
            ('LF', ['N==Literary form', '']),
            ('LA', ['N==Languages', 'COMPLEX']),
            ('CO', ['N==Contents', '']),
            ('AB', ['N==Abstract', '']),
            ('NN', ['N==Notes', '||ScopeContent']),
            ('CA', ['N==Additional notes for cartographic materials', '||DecimalLatitude|DecimalLongitude|Latitude|Longitude|Orientation']),
            ('MA', ['N==Additional notes for music', '']),
            ('PV', ['N==Provenance', '||ImmSourceAcquisition|CustodialHistory|AdministrativeContext']),
            ('RF', ['N==Referenced in', '||PublicationNote']),
            ('NL', ['N==Link to digitised resource', '']),
            ('_8F', ['N==852 holdings flag', '']),  # DIFFERS FROM STANDARD RF
            ('ND', ['N==NID', '']),
            ('EL', ['N==Encoding level', '']),
            ('SX', ['N==Status', 'COMPLEX']),
        ])


class Authority:
    def __init__(self, record, a, atype):
        self.a = a
        self.atype = atype

        # Remove <AdditionalTitles> elements to avoid conflict with <Title> element
        record = re.sub(r'<AdditionalTitles>.*?</AdditionalTitles>', '', record)

        for item in self.a:
            self.a[item] = ''
            try:
                if REGEXES[item].search(record) and REGEXES[item].search(record).group(1).lower() not in \
                        ['-', 'not applicable', 'undetermined', 'unknown', 'unspecified']:
                    self.a[item] = REGEXES[item].search(record).group(1)
                    if item == 'isni':
                        self.a[item] = 'http://isni.org/isni/' + self.a[item]
                    elif item == 'viaf':
                        self.a[item] = 'http://viaf.org/viaf/' + self.a[item]
            except:
                print('\nError [c001]: {}\n'.format(str(sys.exc_info())))

        self.name = clean_authorities(
            ', '.join(self.a[item] for item in self.a if 'dates' not in item and item not in ['isni', 'viaf'] and self.a[item] != ''))
        self.dates = clean_authorities(
            ', '.join(self.a[item] for item in self.a if 'dates' in item and self.a[item] != ''))
        self.isni = ', '.join(self.a[item] for item in self.a if 'isni' in item and self.a[item] != '')
        self.viaf = ', '.join(self.a[item] for item in self.a if 'viaf' in item and self.a[item] != '')

    def __str__(self):
        return clean_authorities(', '.join(self.a[item] for item in self.a if self.a[item] != ''))


class Corporation(Authority):
    def __init__(self, record):
        a = OrderedDict([
            ('c_name', ''),
            ('c_qualifiers', ''),
            ('c_jurisdiction', ''),
            ('c_dates', ''),
            ('isni', ''),
            ('viaf', ''),
        ])
        Authority.__init__(self, record, a, 'corporation')


class Family(Authority):
    def __init__(self, record):
        a = OrderedDict([
            ('f_surname', ''),
            ('f_epithet', ''),
            ('f_dates', ''),
            ('isni', ''),
            ('viaf', ''),
        ])
        Authority.__init__(self, record, a, 'family')


class Person(Authority):
    def __init__(self, record):
        a = OrderedDict([
            ('p_surname', ''),
            ('p_forename', ''),
            ('p_title', ''),
            ('p_epithet', ''),
            ('p_dates', ''),
            ('isni', ''),
            ('viaf', ''),
        ])
        Authority.__init__(self, record, a, 'person')


class Place(Authority):
    def __init__(self, record):
        a = OrderedDict([
            ('pl_name', ''),
            ('pl_localUnit', ''),
            ('pl_widerUnit', ''),
            ('pl_country', ''),
        ])
        Authority.__init__(self, record, a, 'place')


class Subject(Authority):
    def __init__(self, record):
        a = OrderedDict([
            ('s_text', ''),
        ])
        try:
            atype = REGEXES['s_type'].search(record).group(1).lower()
        except:
            atype = 'general term'
        Authority.__init__(self, record, a, atype)


class ArchiveDescription:
    cols = Output()

    def __init__(self, record, authorities):
        self.text = clean(record)
        # Rename additional titles to prevent conflicting element names
        # <Title> element inside <AdditionalTitle> is now named <ATitle>
        # <TitleType> element is now named <TTitleType>
        # <TitleType> element inside <AdditionalTitle> is now named <ATTitleType>
        self.text = re.sub(r'<AdditionalTitle>([^<>]*?)<Title>([^<>]*?)</Title>',
                         r'<AdditionalTitle>\1<ATitle>\2</ATitle>', self.text)
        self.text = re.sub(r'<TitleType>([^<>]*?)</TitleType>([^<>]*?)</AdditionalTitle>',
                         r'<ATTitleType>\1</ATTitleType>\2<AdditionalTitle>', self.text)
        self.text = self.text.replace('<TitleType', '<TTitleType').replace('</TitleType', '</TTitleType')
        self.output = Output()
        self.subjects = set()
        self.names = set()
        self.titles = set()
        self.authorities = authorities
        for item in self.output.values:
            self.output.values[item] = set()

        try: self.ID = str(record.split(',')[1])
        except: self.ID = ''
        self.output.values['_ID'].add(self.ID)
        self.type = record_type(self.ID)

        for k, v in self.cols.values.items():
            if v[1].startswith('||'):
                for c in v[1].strip('|').split('|'):
                    try:
                        self.output.values[k].add(
                            quick_clean(self.text.split('<' + c, 1)[1].split('>', 1)[1].split('</' + c, 1)[0]))
                    except: pass
            else:
                try: sub = getattr(self, k)
                except: pass
                else:
                    try:
                        if callable(sub):
                            self.output.values[k].add(sub())
                    except:
                        print('\nError [cad001]: {}\n'.format(str(sys.exc_info())))

        # IAMS <ExternalIdentifiers> element contains identifiers for fields VF, II, LC and OI
        for match in REGEXES['OI'].findall(self.text):
            if match[0] and match[0] !='':
                if 'VIAF' in match[1]:
                    self.output.values['VF'].add('http://viaf.org/viaf/' + match[0])
                elif 'ISNI' in match[1]:
                    self.output.values['II'].add('http://isni.org/isni/' + match[0])
                elif 'LCCN' in match[1]:
                    self.output.values['LC'].add(match[0])
                else:
                    if match[1] and match[1]!='':
                        self.output.values['OI'].add('{} [{}]'.format(match[0], match[1]))
                    else:
                        self.output.values['OI'].add(match[0])

        # IAMS <DigitalFormatName> element is added to field DS
        try: f = quick_clean(self.text.split('<DigitalFormatName>', 1)[1].split('</DigitalFormatName>', 1)[0])
        except: f = ''
        if f and f != '':
            self.output.values['DS'].add('Digital file format: {}.'.format(f))

        # <AdditionalTitle> is repeatable, so requires a regular expression
        # Also add title to self.titles
        for match in REGEXES['TV'].findall(self.text):
            self.output.values['TV'].add(match)
            self.titles.add(match)
        for match in REGEXES['b_title'].findall(self.text):
            self.titles.add(match)

        # Languages is repeatable, so requires a regular expression
        for match in REGEXES['LA'].findall(self.text):
            if match.lower() not in ['multiple languages', 'not applicable', 'undetermined', 'unknown', 'unspecified']:
                self.output.values['LA'].add(match)

        for match in REGEXES['S_LANGUAGES'].findall(self.text):
            if match.lower() not in ['mul', 'und', 'zxx']:
                self.output.values['S_LANGUAGES'].add(match)

        # Subjects requires authority lookup
        for match in REGEXES['SU'].findall(self.text):
            if match in self.authorities and str(self.authorities[match]) != '':
                self.subjects.add(self.authorities[match])
                self.output.values['SU'].add(str(self.authorities[match]))
                if record_type(match) == 'Place':
                    if len(self.output.values['G1']) == 0:
                        self.output.values['G1'].add(str(self.authorities[match]))
                    elif len(self.output.values['G2']) == 0:
                        self.output.values['G2'].add(str(self.authorities[match]))

        # Names requires authority lookup
        for match in REGEXES['AN'].findall(self.text):
            if match[0] in self.authorities and str(self.authorities[match[0]]) != '':
                if match[1].lower() == 'subject':
                    # Store a the name object
                    self.subjects.add(self.authorities[match[0]])
                    self.output.values['SU'].add(str(self.authorities[match[0]]))
                elif match[1] != '':
                    # Store a tuple containing the name object and the relationship
                    self.names.add((self.authorities[match[0]], match[1].lower()))
                    self.output.values['AN'].add(str(self.authorities[match[0]]) + ' [' + match[1].lower() + ']')
                    if match[1].lower() in ['author', 'creator'] and len(self.output.values['AA']) == 0 \
                            and self.authorities[match[0]].name != '':
                        self.output.values['AA'].add(self.authorities[match[0]].name)
                        self.output.values['AD'].add(self.authorities[match[0]].dates)
                        self.output.values['AT'].add(self.authorities[match[0]].atype)
                        self.output.values['AR'].add(match[1].lower())
                        # Add ISNI and VIAF from first author
                        self.output.values['II'].add(self.authorities[match[0]].isni)
                        self.output.values['VF'].add(self.authorities[match[0]].viaf)

    def __str__(self):
        return self.text

    def RT(self):
        try:
            return quick_clean(record_type(self.ID) + '. ' +
                               self.text.split('<MaterialType', 1)[1].split('>', 1)[1].split('</MaterialType>', 1)[0])
        except: return ''

    def SM(self):
        try:
            ref = quick_clean(self.text.split('<Reference', 1)[1].split('>', 1)[1].split('</Reference>', 1)[0])
        except: ref = ''
        try:
            collection = quick_clean(
                self.text.split('<CollectionArea', 1)[1].split('>', 1)[1].split('</CollectionArea>', 1)[0])
        except: collection = ''
        return quick_clean(collection + '. ' + ref)

    def SX(self):
        try: return STATUS[self.text.split(',')[3]]
        except: return ''


# ====================
#      Functions
# ====================


def create_table(conn, cursor, table_name, debug=False):
    """Function to create a table within the database"""
    if table_name is None or table_name not in TABLE_DEFINITIONS: exit_prompt('Table name not recognised')

    print('\n\nCreating {} table'.format(table_name))
    print('----------------------------------------')
    print(str(datetime.datetime.now()))

    cursor.execute('DROP TABLE IF EXISTS {};'.format(table_name))

    if table_name == 'records':
        fields = Output()
        sql_command = 'CREATE TABLE records (id INTEGER PRIMARY KEY, RecordId NCHAR(13), {Fields});'.format(
            Fields=', '.join(item + ' NTEXT' for item in fields.values).replace('S_DATE1 NTEXT', 'S_DATE1 INTEGER').replace('S_DATE2 NTEXT', 'S_DATE2 INTEGER'))
    else:
        sql_command = 'CREATE TABLE {} (id INTEGER PRIMARY KEY, RecordId NCHAR(13), {});'.format(
            table_name, ', '.join('{} {}'.format(key, value) for (key, value) in TABLE_DEFINITIONS[table_name]))
    if debug:
        print(sql_command)
    cursor.execute(sql_command)
    conn.commit()
    gc.collect()


def build_index(conn, cursor, table_name, index_name):
    """Function to build an index on a table within the database"""
    if table_name is None or index_name is None: exit_prompt('Error building index {} on table {}'.format(index_name, table_name))
    print('Building index {}'.format(index_name))
    cursor.execute("""DROP INDEX IF EXISTS {};""".format(index_name))
    cursor.execute("""CREATE INDEX {} ON {} (RecordId ASC)""".format(index_name, table_name))
    conn.commit()
    gc.collect()
    return


def dump_table(cursor, table_name):
    """Function to dump a database table into a text file"""
    if table_name is None: exit_prompt('Table name not recognised')
    record_count = 0
    try:
        cursor.execute('SELECT * FROM {};'.format(table_name))
    except:
        print('{} table does not exist'.format(table_name))
    else:
        print('Creating dump of {} table'.format(table_name))
        file = open('{}.txt'.format(table_name), mode='w', encoding='utf-8', errors='replace')
        row = cursor.fetchone()
        while row:
            record_count += 1
            file.write('{}\n'.format(str(row)))
            row = cursor.fetchone()
        file.close()
        gc.collect()
        print('{} records in {} table'.format(str(record_count), table_name))
    return record_count


def clean(string):
    """Function to clean punctuation and normalize Unicode"""
    if string is None or not string or string == 'None':
        return ''
    string = string.strip()
    string = re.sub(
        u'[\u0022\u055A\u05F4\u2018\u2019\u201A\u201B\u201C\u201D\u201E\u201F\u275B\u275C\u275D\u275E\uFF07\u0060]',
        '\'',
        string)
    string = re.sub(
        u'[\u0000-\u0009\u000A-\u000f\u0010-\u0019\u001A-\u001F\u0080-\u0089\u008A-\u008F\u0090-\u0099\u009A-\u009F\u2028\u2029]+',
        '', string)
    string = re.sub(u'[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]+', ' ', string)
    string = re.sub(u'[\u002D\u2010-\u2015\u2E3A\u2E3B\uFE58\uFE63\uFF0D]+', '-', string)
    string = string.replace('&gt;', '>').replace('&lt;', '<').replace('&amp;', '&')
    string = re.sub(r'\s+', ' ', string).strip()
    string = re.sub(r'<[^/>]+/>', '', string, flags=re.IGNORECASE)
    string = re.sub(r'[.\s]*</p>\s*', '. ', re.sub(r'[;\s.]+;', ';', string.replace('</item>', '; ')),
                    flags=re.IGNORECASE)
    string = re.sub(
        r'[<\[]/*(b|br|emph|i|italic|italics|item|li|list|ol|p|sup|superscript|sub|subscript|ul)(\s+[^>\]]+)?\s*/*[>\]]',
        ' ', string, flags=re.IGNORECASE)
    string = re.sub(r';\s+\.', '.', string).strip()
    string = re.sub(r'\s+', ' ', string).strip()
    string = unicodedata.normalize('NFC', string)
    return string


def quick_clean(string, hyphens=True):
    """If a string has been cleaned with clean() once, quick_clean() is sufficient for subsequent cleaning
        If hyphens=True, trailing/leading hyphens are preserved"""
    if string is None or not string or string == 'None':
        return ''
    string = re.sub(r';\s+\.', '.', string).strip()
    string = re.sub(r'\s+', ' ', string).strip()
    l = '?$.,:;/\])} ' if hyphens else '?$.,:;/\-])} '
    r = '.,:;/\[({ ' if hyphens else '.,:;/\-[({ '
    string = re.sub(r'\s+', ' ', string.strip().lstrip(l).rstrip(r)).strip()
    string = string.replace('( ', '(').replace(' )', ')')
    string = string.replace(' ,', ',').replace(',,', ',').replace(',.', '.').replace('.,', ',')
    string = string.replace('. [', ' [').replace(' : (', ' (')
    string = string.replace('= =', '=').replace('= :', '=').replace('+,', '+')
    return string


def clean_authorities(string):
    """Function to clean the string form of an authority record"""
    string = re.sub(r', Family(?:,|$)', ' family', string)
    string = re.sub(r'(, |^)fl (?:[0-9])', r'\1active ', string)
    string = re.sub(r'(, |^)b ([0-9]{4})', r'\1\2-', string)
    string = re.sub(r'(, |^)d (?:[0-9])', r'\1-', string)
    string = re.sub(r' cent$', ' century', string)
    string = re.sub(r'(, |- |^)c\s*([0-9]+)', r'\1approximately \2', string)
    string = string.replace(' - ', '-')
    return quick_clean(string)


def clean_msg(string):
    """Function to clean punctuation and normalize Unicode in an Outlook .msg file"""
    if string is None or not string or string == '': return ''
    string = string.replace('"', '\\"').replace('\n', '')
    string = re.sub(r'<[^>]+>', '', re.sub(r'[\u0000\ufffd]', '', string)).replace('&nbsp;', '')
    string = unicodedata.normalize('NFC', string)
    return string


def is_IAMS_id(string):
    """Function to test whether a string is a valid IAMS record ID"""
    if string is None or not string or string == 'None':
        return False
    if RE_IAMS_ID.fullmatch(string) and string.split('-')[0] in TYPES:
        return True
    return False


def record_type(string):
    """Function to return the record type from an IAMS record ID"""
    if is_IAMS_id(string) and string.split('-')[0] in TYPES:
        return TYPES[string.split('-')[0]]
    return


def add_string(string, base, separator):
    """Function to append a string to another string"""
    if string != '' and string not in base:
        if base != '': base = base + separator
        base += string
    return base


def get_boolean(prompt):
    """Function to prompt the user for a Boolean value"""
    while True:
        try:
            return {'Y': True, 'N': False}[input(prompt).upper()]
        except KeyError:
            print('Sorry, your choice was not recognised. Please enter Y or N:')


def csv_row(row):
    """Function to clean a row of CSV data before writing it to an output file"""
    if row is None or not row: return ''
    return '"' + '","'.join(re.sub(r'^None$|\[\s*\]', '', str(s)).replace(',  [', ' [') for s in row) + '"\n'


def run_sql(cursor, sql_command, ofile, debug=False):
    """Function to run an SQL command and write the results to a CSV file"""
    j = 0
    if debug:
        print(str(cursor.execute("""EXPLAIN QUERY PLAN {}""".format(sql_command)).fetchall()))
    cursor.execute(sql_command)
    row = cursor.fetchone()
    while row:
        ofile.write(csv_row(row))
        j += 1
        print('\r{} records written to file'.format(str(j)), end='\r')
        try: row = cursor.fetchone()
        except: return j


def check_file_location(file_path, function, file_ext='', exists=False):
    """Function to check whether a file exists and has the correct file extension."""
    folder, file, ext = '', '', ''
    if file_path == '':
        exit_prompt('Error: Could not parse path to {} file'.format(function))
    try:
        file, ext = os.path.splitext(os.path.basename(file_path))
        folder = os.path.dirname(file_path)
    except:
        exit_prompt('Error: Could not parse path to {} file'.format(function))
    if file_ext != '' and ext != file_ext:
        exit_prompt('Error: The specified file should have the extension {}'.format(file_ext))
    if exists and not os.path.isfile(os.path.join(folder, file + ext)):
        exit_prompt('Error: The specified {} file cannot be found'.format(function))
    return folder, file, ext


def exit_prompt(message=''):
    """Function to exit the program after prompting the use to press Enter"""
    if message != '':
        print(str(message))
    input('\nPress [Enter] to exit...')
    sys.exit()


# ====================


class Converter(object):
    """A class for converting IAMS data.

    :param debug: Display additional output to assist with debugging.
    """

    def __init__(self, debug=False):
        self.debug = debug

    def show_header(self):
        if self.header:
            print(self.header)


class SQL2RF(Converter):

    def __init__(self, debug=False):
        self.search_criteria = {
            'l1': set(),
            'txt': set(),
            'd1': '',
            'd2': '',
        }
        self.output_fields = Output()
        self.search_string = ''
        self.search_list = ''
        self.ids = set()
        self.header = '========================================\n' \
                      'sql2rf\n' \
                      'IAMS data extraction for Researcher Format\n' \
                      '========================================\n' \
                      'This utility searches an SQL database of IAMS records\n' \
                      'created using the utility snapshot2sql\n' \
                      'and converts matching records to Researcher Format\n'
        Converter.__init__(self, debug)

    def iams2rf_sql2rf(self, db_path, request_path, output_folder):
        """Search for records within an SQL database created using snapshot2sql
        and convert to Researcher Format

        :param db_path: Path to the SQL database.
        :param request_path: Path to Outlook message containing details of the request.
        :param output_folder: Folder to save Researcher Format output files.
        """

        self.show_header()

        # Check file locations
        db_folder, db_file, db_ext = check_file_location(db_path, 'SQL database', '.db', True)
        if request_path != '':
            request_folder, request_file, request_ext = check_file_location(request_path, 'request message', '.msg', True)
        if output_folder != '':
            try:
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)
            except os.error:
                exit_prompt('Error: Could not create folder for output files')


        # --------------------
        # Parameters seem OK => start script
        # --------------------

        # Display confirmation information about the transformation
        print('SQL database: {}'.format(db_file + db_ext))
        if request_path != '':
            print('Request message: {}'.format(request_file + request_ext))
        if output_folder != '':
            print('Output folder: {}'.format(output_folder))
        if self.debug:
            print('Debug mode')

        # --------------------
        # Define possible output fields
        # --------------------

        # If request message has been specified, use this to determine transformation parameters
        if request_path != '':

            # Determine output fields to be included based on contents of request message file
            msgfile = open(os.path.join(request_folder, request_file + request_ext), mode='r', encoding='utf-8', errors='replace')
            for filelineno, line in enumerate(msgfile):
                line = clean_msg(line)
                if 'Coded parameters for your transformation' in line: break
            for filelineno, line in enumerate(msgfile):
                line = clean_msg(line)
                if 'End of coded parameters' in line: break
                if '=' in line:
                    line = clean_msg(line)
                    parameter = line.split('=', 1)[0].strip()
                    values = re.sub(r'^ ', '', line.split('=', 1)[-1])
                    if parameter != '' and values != '':
                        if parameter == 'o':
                            for f in re.sub(r'[^a-zA-Z0-9|]', '', values).split('|'):
                                f = f.replace('ID', '_ID').replace('IS', '_IS').replace('8F', '_8F')
                                if f in self.output_fields.values:
                                    self.output_fields.values[f][0] = self.output_fields.values[f][0].replace('N==', 'Y==')
                        elif parameter == 'v' and (re.sub(r'[^rtnscRTNSC|]', '', values) == values):
                            values = values.lower()
                            file_records = 'r' in values
                            file_titles = 't' in values
                            file_names = 'n' in values
                            file_topics = 's' in values

                        elif parameter in ['l1', 'txt']:
                            # Languages codes and search strings
                            for v in re.sub(r'\$[a-z0-9]', ' ', re.sub(r'([^\x00-\x7F]|,)', '_', values)).split('|'):
                                self.search_criteria[parameter].add(v)
                        elif parameter in ['d1', 'd2']:
                            # Date range
                            if len(re.sub(r'[^0-9]', '', values)) >= 4:
                                self.search_criteria[parameter] = re.sub(r'[^0-9]', '', values)[:4]

            if len(self.search_criteria['l1']) > 0:
                self.search_string = '( ' + \
                                     ' OR '.join((' S_LANGUAGES LIKE "%{}%" '.format(s))
                                                 for s in sorted(self.search_criteria['l1'])) \
                                     + ' )'
            if len(self.search_criteria['txt']) > 0:
                self.search_string = add_string('( {} )'.format(
                    ' OR '.join('{} LIKE "%{}%"'.format(f, s)
                                for f in ['AA', 'AN', 'TT', 'DS', 'SM', 'SU', 'NN', 'PV', 'RF']
                                for s in sorted(self.search_criteria['txt']))), self.search_string, ' AND ')
            self.search_string = add_string(self.search_criteria['d1'], self.search_string, ' AND S_DATE2 >= ')
            self.search_string = add_string(self.search_criteria['d2'], self.search_string, ' AND S_DATE1 <= ')
            if self.search_string != '':
                self.search_string = 'WHERE ' + self.search_string
            if self.debug:
                try:
                    print(self.search_string)
                    print(str(self.search_criteria))
                except ValueError: pass

        # If request message has not been specified, user must provide transformation parameters
        else:
            # User selects columns to include
            print('\n')
            print('----------------------------------------')
            print('Select one of the following options: \n')
            print('   D     Default columns \n')
            print('   A     All columns \n')
            print('   S     Select columns to include \n\n')
            selection = input('Choose an option:').upper()
            while selection not in ['D', 'A', 'S']:
                selection = input('Sorry, your choice was not recognised. Please enter D, A, or S:').upper()

            # Default column selection
            if selection == 'D':

                # Columns to include
                for f in ['_ID', 'RT', 'BN', 'IB', 'AA', 'AD', 'AT', 'AR', 'AN', 'TT', 'TV', 'SE', 'SN', 'PC',
                          'PP', 'PB', 'PD', 'ED', 'DS', 'DW', 'SM', 'SU', 'GE', 'LA', 'NN', 'AK', 'PV', 'RF']:
                    self.output_fields.values[f][0] = self.output_fields.values[f][0].replace('N==', 'Y==')

            # All columns
            elif selection == 'A':

                # Columns to include
                # Select all columns
                for f in self.output_fields.values:
                    self.output_fields.values[f][0] = self.output_fields.values[f][0].replace('N==', 'Y==')
                # Remove context-specific columns
                for f in ['ES', '_8F', 'BU', 'CG', 'CL', 'EL', 'FA', 'G1', 'G2', 'HA', 'HF', 'HL',
                          'IO', 'ND', 'NL', 'P1', 'P2', 'PJ', 'SD', 'SO', 'SX']:
                    self.output_fields.values[f][0] = self.output_fields.values[f][0].replace('Y==', 'N==')

            # Choose columns one-by-one
            else:
                print('\nChoose the optional columns: \n')

                for f in self.output_fields.values:
                    skip = False
                    # Skip some choices depending upon the values of parameters already set
                    if f in ['AK', 'PV']:
                        skip = True
                        self.output_fields.values[f][0] = self.output_fields.values[f][0].replace('N==', 'Y==')
                    if f in ['_8F', 'NL', 'P1', 'P2', 'SD', 'SO', 'SX'] \
                            or (f in ['AD', 'AT', 'AR', 'II', 'VF'] and self.output_fields.values['AA'][0].startswith('N==')) \
                            or (f == 'SN' and self.output_fields.values['SE'][0].startswith('N==')) \
                            or (f == 'PU' and self.output_fields.values['PD'][0].startswith('N==')):
                        skip = True
                    if not skip:
                        opt_text = self.output_fields.values[f][0].replace('N==', '').replace('Y==', '')
                        # Add additional explanatory text to column heading when presenting user with choices
                        if f == 'AA':
                            opt_text += ' (of first author)'
                        elif f in ['FA', 'FC', 'FF', 'HA', 'HF', 'HL', 'IL', '_IS', 'PG', 'TK']:
                            opt_text += ' (relevant to serials only)'
                        elif f in ['BU', 'CG', 'CL', 'IO', 'ND']:
                            opt_text += ' (relevant to newspapers only)'
                        elif f in ['IM', 'MF', 'MG', 'MA']:
                            opt_text += ' (relevant to music only)'
                        elif f in ['SC', 'JK', 'CD', 'CA']:
                            opt_text += ' (relevant to cartographic materials only)'
                        opt_input = get_boolean('Include {0}? (Y/N):'.format(opt_text))
                        if opt_input:
                            self.output_fields.values[f][0] = self.output_fields.values[f][0].replace('N==', 'Y==')
                        else:
                            self.output_fields.values[f][0] = self.output_fields.values[f][0].replace('Y==', 'N==')

            # User selects output files to include
            print('\n----------------------------------------')
            print('Select output files to include:\n')
            file_records = get_boolean('Include the Records file? (Y/N):')
            file_titles = get_boolean('Include the Titles file? (Y/N):')
            file_names = get_boolean('Include the Names file? (Y/N):')
            file_topics = get_boolean('Include the Topics file? (Y/N):')

        # --------------------
        # Build column headings and create output files
        # --------------------

        records_header = ''
        if file_records:
            for f in self.output_fields.values:
                if self.output_fields.values[f][0].startswith('Y=='):
                    records_header += ',"' + self.output_fields.values[f][0].replace('Y==', '') + '"'
            records_header = records_header.strip(',') + '\n'
            records = open(os.path.join(output_folder,'records_IAMS.csv'), mode='w', encoding='utf-8', errors='replace')
            records.write(records_header)

        names_header = '"Name","Dates associated with name","Type of name","Role","Other names"'
        if file_names:
            for f in self.output_fields.values:
                if self.output_fields.values[f][0].startswith('Y==') and f not in ['AA', 'AD', 'AT', 'AR', 'II', 'VF', 'AN']:
                    names_header += ',"' + self.output_fields.values[f][0].replace('Y==', '') + '"'
            names_header = names_header.strip(',') + '\n'
            names = open(os.path.join(output_folder,'names_IAMS.csv'), mode='w', encoding='utf-8', errors='replace')
            names.write(names_header)

        titles_header = '"Title","Other titles"'
        if file_titles:
            for f in self.output_fields.values:
                if self.output_fields.values[f][0].startswith('Y==') and f not in ['TT', 'TV', 'TU', 'TK']:
                    titles_header += ',"' + self.output_fields.values[f][0].replace('Y==', '') + '"'
            titles_header = titles_header.strip(',') + '\n'
            titles = open(os.path.join(output_folder,'titles_IAMS.csv'), mode='w', encoding='utf-8', errors='replace')
            titles.write(titles_header)

        topics_header = '"Topic","Type of topic"'
        if file_topics:
            for f in self.output_fields.values:
                if self.output_fields.values[f][0].startswith('Y==') and f != 'SU':
                    topics_header += ',"' + self.output_fields.values[f][0].replace('Y==', '') + '"'
            topics_header = topics_header.strip(',') + '\n'
            topics = open(os.path.join(output_folder,'topics_IAMS.csv'), mode='w', encoding='utf-8', errors='replace')
            topics.write(topics_header)

        # --------------------
        # Connect to local database
        # --------------------

        print('\nConnecting to local database ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))

        conn = sqlite3.connect(os.path.join(db_folder, db_file + db_ext))
        cursor = conn.cursor()

        print('\nSearching for matching records ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))

        # Get a list of matching records
        i = 0
        format_str = """
SELECT RecordId FROM records {where}
ORDER BY RecordId ASC;"""
        try: sql_command = format_str.format(where=self.search_string)
        except:
            exit_prompt('Error creating XSL command to search for matching records: {}'.format(str(sys.exc_info())))
        else:
            if self.debug:
                print(str(cursor.execute("""EXPLAIN QUERY PLAN {}""".format(sql_command)).fetchall()))
            cursor.execute(sql_command)
            row = cursor.fetchone()
            while row:
                self.ids.add(row[0])
                i += 1
                print('\r{} matching records'.format(str(i)), end='\r')
                try: row = cursor.fetchone()
                except: break

        # Remove records not to be exported
        if os.path.isfile(os.path.join(db_folder, 'List of IDs not to be exported.txt')):
            print('\n\nRemoving records that should not be exported ...')
            eap = set()
            ifile = open(os.path.join(db_folder, 'List of IDs not to be exported.txt'), mode='r', encoding='utf-8', errors='replace')
            for filelineno, line in enumerate(ifile):
                line = line.strip()
                if is_IAMS_id(line): eap.add(line)
            ifile.close()
            self.ids = self.ids - eap
            del eap
        gc.collect()
        self.search_list = '\'' + str('\', \''.join(self.ids)) + '\''

        # Records
        if file_records:

            print('\n\nWriting records file ...')
            format_str = """
SELECT {search_fields} FROM records
WHERE RecordId IN ({search_list}) ORDER BY RecordId ASC;"""
            sql_command = format_str.format(
                search_fields=', '.join(str(f) for f in self.output_fields.values
                                        if self.output_fields.values[f][0].startswith('Y==')),
                search_list=self.search_list)
            run_sql(cursor, sql_command, records, self.debug)

        # Titles
        if file_titles:
            print('\n\nWriting titles file ...')

            format_str = """
SELECT t1.Title,
( SELECT GROUP_CONCAT(t2.Title, ' ; ')
FROM  titles t2
WHERE ( t2.RecordId = t1.RecordId AND t1.Title <> t2.Title )
ORDER BY t2.Title ASC
) AS otherTitles,
{search_fields}
FROM titles t1
INNER JOIN records ON records.RecordId = t1.RecordId
WHERE t1.RecordId IN ({search_list}) ORDER BY t1.Title ASC ;"""
            sql_command = format_str.format(
                search_fields=', '.join(('records.' + str(f)) for f in self.output_fields.values
                                       if self.output_fields.values[f][0].startswith('Y==')
                                       and f not in ['TK', 'TT', 'TU', 'TV']),
                search_list=self.search_list)
            run_sql(cursor, sql_command, topics, self.debug)

            # Old format_str
            # format_str =
            """
SELECT TT, TV, {search_fields} FROM records
WHERE ( RecordId IN ({search_list}) AND TT <> '' )
ORDER BY RecordId ASC;"""

        # Names
        if file_names:

            print('\n\nWriting names file ...')
            format_str = """
SELECT n1.Name, n1.NameDates, n1.NameType, n1.NameRole, n1.NameISNI, n1.NameVIAF,
( SELECT GROUP_CONCAT(n2.Name || ', ' || n2.NameDates || ' [' || n1.NameRole || '], ' || n2.NameISNI || ', ' || n2.NameVIAF, ' ; ')
FROM  names n2
WHERE ( n2.RecordId = n1.RecordId AND n1.Name <> n2.Name AND n1.NameDates <> n2.NameDates )
ORDER BY n2.Name ASC
) AS otherNames,
{search_fields}
FROM names n1
INNER JOIN records ON records.RecordId = n1.RecordId
WHERE n1.RecordId IN ({search_list}) ORDER BY n1.Name ASC ;"""
            sql_command = format_str.format(
                search_fields=', '.join(('records.' + str(f)) for f in self.output_fields.values
                                       if self.output_fields.values[f][0].startswith('Y==')
                                       and f not in ['AN', 'AA', 'AD', 'AT', 'AR', 'II', 'VF']),
                search_list=self.search_list)
            run_sql(cursor, sql_command, names, self.debug)

        # Subjects
        if file_topics:

            print('\n\nWriting subjects file ...')
            format_str = """
SELECT subjects.Topic, subjects.TopicType, {search_fields}
FROM subjects INNER JOIN records ON subjects.RecordId = records.RecordId
WHERE subjects.RecordId IN ({search_list}) ORDER BY subjects.Topic ASC ;"""
            sql_command = format_str.format(
                search_fields=', '.join(('records.' + str(f)) for f in self.output_fields.values
                                       if self.output_fields.values[f][0].startswith('Y==')
                                       and f not in ['SU']),
                search_list=self.search_list)
            run_sql(cursor, sql_command, topics, self.debug)

        # Close files
        for file in [records, names, titles, topics]:
            try: file.close()
            except: pass

        # Close connection to local database
        conn.close()


class IAMS2SQL(Converter):

    def __init__(self, debug=False):
        self.authorities = {}
        self.header = '========================================\n' + \
                      'snapshot2sql\n' + \
                      'IAMS data preparation for Researcher Format\n' + \
                      '========================================\n' + \
                      'This utility converts the IAMS Published Snapshot to an SQL database\n'
        Converter.__init__(self, debug)

    def iams2rf_snapshot2sql(self, iams_snapshot_path, db_path):
        """Convert the IAMS Published Snapshot to an SQL database

        :param iams_snapshot_path: Path to the IAMS Published Snapshot.
        :param db_path: Path to save the SQL database.
        """

        self.show_header()

        # Check file locations
        iams_folder, iams_file, iams_ext = check_file_location(iams_snapshot_path, 'IAMS Database Snapshot', '.csv', True)
        db_folder, db_file, db_ext = check_file_location(db_path, 'SQL database', '.db')

        # --------------------
        # Parameters seem OK => start script
        # --------------------

        # Display confirmation information about the transformation
        print('IAMS Database Snapshot: {}'.format(str(os.path.join(iams_folder, iams_file + iams_ext))))
        print('SQL database: {}'.format(str(os.path.join(db_folder, db_file + db_ext))))
        if self.debug:
            print('Debug mode')

        # --------------------
        # Build indexes
        # --------------------

        print('\nBuilding indexes ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))

        rfile = open(os.path.join(iams_folder, iams_file + iams_ext), mode='r', encoding='utf-16-le', errors='replace')
        i = 0
        rec, rid = '', ''

        # Don't process lines before authorities have been reached
        if self.debug:
            print('Enumerating IAMS Database Snapshot ... Authorities not yet reached')
        for filelineno, line in enumerate(rfile):
            if '},045-' in line:
                rec = line
                if self.debug: print('Reached the authorities')
                break
        gc.collect()

        for filelineno, line in enumerate(rfile):
            line = line.strip()
            if line.startswith('{') and rec:
                i += 1
                print('\r{} records indexed'.format(str(i)), end='\r')
                rec = clean(rec)
                rid = rec.split(',')[1]
                if record_type(rid) == 'Corporation':
                    self.authorities[rid] = Corporation(rec)
                elif record_type(rid) == 'Family':
                    self.authorities[rid] = Family(rec)
                elif record_type(rid) == 'Person':
                    self.authorities[rid] = Person(rec)
                elif record_type(rid) == 'Place':
                    self.authorities[rid] = Place(rec)
                elif record_type(rid) == 'Subject':
                    self.authorities[rid] = Subject(rec)
                rec = line
            else:
                rec += line
        # Ensure last record in the file is processed
        if rec:
            i += 1
            print('\r{} records indexed'.format(str(i)), end='\r')
            rec = clean(rec)
            rid = rec.split(',')[1]
            if record_type(rid) == 'Corporation':
                self.authorities[rid] = Corporation(rec)
            elif record_type(rid) == 'Family':
                self.authorities[rid] = Family(rec)
            elif record_type(rid) == 'Person':
                self.authorities[rid] = Person(rec)
            elif record_type(rid) == 'Place':
                self.authorities[rid] = Place(rec)
            elif record_type(rid) == 'Subject':
                self.authorities[rid] = Subject(rec)

        rfile.close()
        gc.collect()

        # --------------------
        # Connect to local database
        # --------------------

        print('\n\nConnecting to local database ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))

        if self.debug:
            print('sqlite connection: {}'.format(str(os.path.join(db_folder, db_file + db_ext))))
        conn = sqlite3.connect(os.path.join(db_folder, db_file + db_ext))
        cursor = conn.cursor()

        # Create tables
        fields = Output()
        for table_name in ['records', 'names', 'subjects', 'titles']:
            create_table(conn, cursor, table_name, debug=self.debug)

        # Add records to database
        # ====================================================================================================

        print('\n\nAdding records to the database ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))

        # --------------------
        # Open CSV file
        # --------------------

        print('\n\nOpening CSV file from IAMS snapshot ...')
        rfile = open(os.path.join(iams_folder, iams_file + iams_ext), mode='r', encoding='utf-16-le', errors='replace')
        rec = ''
        i = 0
        for filelineno, line in enumerate(rfile):
            line = line.strip()
            if line.startswith('{') and rec:
                i += 1
                print('\r{} records processed'.format(str(i)), end='\r')
                record = ArchiveDescription(rec, self.authorities)
                # Need to amend something here

                if record.type in ['Corporation', 'Family', 'Person', 'Place', 'Subject']:
                    break

                if is_IAMS_id(record.ID):
                    try:
                        format_str = """
INSERT INTO records (id, RecordId, {Fields})
VALUES (NULL, "{RecordId}", {Values}); """
                        sql_command = format_str.format(RecordId=record.ID,
                                                        Fields=', '.join(item for item in fields.values),
                                                        Values='"' + '", "'.join(str(p) for p in [
                                                            ' ; '.join(
                                                                str(q) for q in sorted(record.output.values[item]))
                                                            for item in fields.values]) + '"')
                        cursor.execute(sql_command)
                    except:
                        print('\nError [at002]: {}\n'.format(str(sys.exc_info())))

                    # Save names
                    for n in record.names:
                        try:
                            format_str = """
INSERT INTO names (id, RecordId, Name, NameDates, NameType, NameRole, NameISNI, NameVIAF)
VALUES (NULL, "{RecordId}", "{Name}", "{NameDates}", "{NameType}", "{NameRole}", "{NameISNI}", "{NameVIAF}"); """
                            sql_command = format_str.format(RecordId=record.ID,
                                                            Name=n[0].name,
                                                            NameDates=n[0].dates,
                                                            NameType=n[0].atype,
                                                            NameRole=n[1],
                                                            NameISNI=n[0].isni,
                                                            NameVIAF=n[0].viaf)
                            cursor.execute(sql_command)
                        except:
                            print('\nError [at003]: {}\n'.format(str(sys.exc_info())))

                    # Save subjects
                    for s in record.subjects:
                        try:
                            format_str = """
INSERT INTO subjects (id, RecordId, Topic, TopicType)
VALUES (NULL, "{RecordId}", "{Topic}", "{TopicType}"); """
                            sql_command = format_str.format(RecordId=record.ID,
                                                            Topic=str(s),
                                                            TopicType=s.atype)
                            cursor.execute(sql_command)
                        except:
                            print('\nError [at004]: {}\n'.format(str(sys.exc_info())))


                    # Save titles
                    for t in record.titles:
                        try:
                            format_str = """
INSERT INTO titles (id, RecordId, Title)
VALUES (NULL, "{RecordId}", "{Title}"); """
                            sql_command = format_str.format(RecordId=record.ID,
                                                            Title=str(t))
                            cursor.execute(sql_command)
                        except:
                            print('\nError [at005]: {}\n'.format(str(sys.exc_info())))

                rec = line
            else:
                rec += line

            # Save changes at every 1000th record
            if filelineno % 1000 == 0:
                conn.commit()

        conn.commit()

        # Build indexes
        # ====================================================================================================

        print('\n\nCreating indexes ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))
        for table_name in ['records', 'names', 'subjects', 'titles']:
            build_index(conn, cursor, table_name, 'IDX_{}'.format(table_name))

        # Text file dumps of tables
        # ====================================================================================================

        print('\n\nWriting tables to text files ...')
        print('----------------------------------------')
        print(str(datetime.datetime.now()))
        for table_name in ['records', 'names', 'subjects', 'titles']:
            dump_table(cursor, table_name)

        # Close connection to local database
        conn.close()
