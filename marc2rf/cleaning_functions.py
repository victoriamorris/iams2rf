# -*- coding: utf8 -*-

"""Data cleaning functions used in the Researcher Format transformation."""

# Import required modules
import os
import sys
import unicodedata

import marc2rf.publisher as publisher
import marc2rf.multiregex as mrx
import regex as re
from marc2rf.lookup import *

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'

# ====================
#     Constants
# ====================

BRACKETS = [('[', ']'), ('(', ')'), ('{', '}')]

# ====================
#  Regular expressions
# ====================

RE_IAMS_ID = re.compile('0[34][0-9]-[0-9]{9}')
RE_ISBN10 = re.compile(r'ISBN\x20(?=.{13}$)\d{1,5}([- ])\d{1,7}'r'\1\d{1,6}\1(\d|X)$|[- 0-9X]{10,16}')
RE_ISBN13 = re.compile(r'97[89]{1}(?:-?\d){10,16}|97[89]{1}[- 0-9]{10,16}')
RE_YEAR = re.compile('(?<![0-9])(1[0-9][0-9]{2}|2[0-9]{3})(?![0-9])')
RE_YEAR_POST_1500 = re.compile('(?<![0-9])(1[5-9][0-9]{2}|2[0-9]{3})(?![0-9])')       # Only matches dates after 1500 to avoid confusion with volume numbers
RE_DATE_1 = re.compile('(janu?a?r?y?|ocak|stycz|febr?u?a?r?y?|marc?h?|apri?l?|ma[iy]|june?|ioun|july?|augu?s?t?|agosto|sept?e?m?b?e?r?|o[ck]to?b?e?r?|nove?m?b?e?r?|listopad|de[czk]e?m?b?e?r?|grudz)\.*\s*([1-9][0-9]?([\-/][1-9]?[0-9]?)?)[^0-9]', flags=re.IGNORECASE)
RE_DATE_2 = re.compile('(?<![0-9])([1-9][0-9]?([\-/][1-9]?[0-9]?)?)(th|st|nd)?\.*\s*(janu?a?r?y?|ocak|stycz|febr?u?a?r?y?|marc?h?|apri?l?|ma[iy]|june?|ioun|july?|augu?s?t?|agosto|sept?e?m?b?e?r?|o[ck]to?b?e?r?|nove?m?b?e?r?|listopad|de[czk]e?m?b?e?r?|grudz)', flags=re.IGNORECASE)
RE_VOLUME = re.compile('(?<![a-z])(vo?l?|tomo|d|god|izd|rhif|rok|jahrg|jahrgang)[.:]*\s*([1-9xiv][0-9lxiv]*)[^0-9lxiv]', flags=re.IGNORECASE)
RE_ISSUE = re.compile('(?<![a-z])(issue|br)[.:]*\s*([1-9lxiv][0-9lxiv]*)[^0-9lxiv]', flags=re.IGNORECASE)
RE_NUMBER = re.compile('(?<![a-z])(no|nr|numb?e?r?|pa?r?t)[.:]*\s*([1-9lxi][0-9lxiv]*)[^0-9lxiv]', flags=re.IGNORECASE)
RE_SERIES_NUMBER = re.compile('\s+(ba?n?d|fasc|he?fte?|jahrga?n?g?|knji?g?a?|n[or](?![a-z])|number|pa?r?t|sva?z?e?k?|volu?m?e?)s?[.\s]*[0-9a-zA-Z,\-.\s/]+$', flags=re.IGNORECASE)
RE_NUMERAL = re.compile('[1-9]+[0-9]*|[cdilmvx]+')

# ====================
#      Functions
# ====================

# Functions for cleaning strings


def clean(string, hyphens=True, space=True):
    """Function to clean punctuation and normalize Unicode."""
    string = re.sub(u'[\u0022\u055A\u05F4\u2018\u2019\u201A\u201B\u201C\u201D\u201E\u201F\u275B\u275C\u275D\u275E\uFF07]', '\'', string)
    string = re.sub(
        u'[\u0000-\u0009\u000A-\u000f\u0010-\u0019\u001A-\u001F\u0080-\u0089\u008A-\u008F\u0090-\u0099\u009A-\u009F\u2028\u2029]+',
        '', string)
    if space:
        string = re.sub(r'\s+', ' ', string)
        string = re.sub(u'[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]+', ' ', string)
        string = re.sub(r'\(adap(ta)?tions\)', '', string, flags=re.IGNORECASE)
        string = re.sub(r'(?<![a-z])pseud\.*(?![a-z])', 'pseudonym', string, flags=re.IGNORECASE)
        string = quick_clean(string, hyphens)
        string = re.sub(r'</*(b|br|emph|i|li|ol|p|sup|sub|ul)\s*/*>', ' ', string, flags=re.IGNORECASE)
        string = re.sub(r'\s+', ' ', string)
        string = check_brackets(string)
    string = unicodedata.normalize('NFC', string).strip()
    return string


def quick_clean(string, hyphens=True):
    """Quick clean.

    If a string has been cleaned with clean() once, quick_clean() is sufficient for subsequent cleaning.
    Subfields are cleaned when retrieved using field.get_subfields(), so only require quick_clean().
    If hyphens=True, trailing/leading hyphens are preserved."""
    l = '?$.,:;/\])} ' if hyphens else '?$.,:;/\-])} '
    r = '.,:;/\[({ ' if hyphens else '.,:;/\-[({ '
    string = re.sub(r'\s+', ' ', string.strip().lstrip(l).rstrip(r)).strip()
    string = string.replace('( ', '(').replace(' )', ')')
    string = string.replace(' ,', ',').replace(',,', ',').replace(',.', '.').replace('.,', ',')
    string = string.replace('. [', ' [').replace(' : (', ' (').replace('= =', '=').replace('= :', '=').replace('+,', '+')
    return string


def clean_msg(string):
    """Function to clean punctuation and normalize Unicode in an Outlook .msg file"""
    if string is None or not string or string == '': return ''
    string = string.replace('"', '\\"').replace('\n', '')
    string = re.sub(r'<[^>]+>', '', re.sub(r'[\u0000\ufffd]', '', string)).replace('&nbsp;', '')
    string = unicodedata.normalize('NFC', string)
    return string


def clean_search_string(string):
    """Function to clean a search string"""
    if string is None or not string or string == '': return ''
    string = string.strip('!"£%^&*()_-+={}[]::@~#<,>.?/|\`¬ ').strip("' ")
    return string


def escape_regex_chars(string):
    """Function to escape regex characters in a string"""
    string = string.replace('\\', '\\' + '\\').replace('^', '\^').replace('$', '\$').replace('.', '\.')
    string = string.replace('|', '\|').replace('?', '\?').replace('*', '\*').replace('+', '\+')
    string = string.replace('(', '\(').replace(')', '\)').replace('[', '\[').replace(']', '\]')
    string = string.replace('{', '\{').replace('}', '\}')
    return string


def check_brackets(string):
    """Function to check for inconsistent brackets"""
    string = quick_clean(string)
    for (oB, cB) in [('[', ']'), ('<', '>')]:
        while string.startswith(oB) and string.endswith(cB):
            string = quick_clean(string[1:-1])
    for (oB, cB) in BRACKETS:
        while string.startswith(oB) and cB not in string:
            string = quick_clean(string[1:])
        while string.endswith(cB) and oB not in string:
            string = quick_clean(string[:-1])
    for (oB, cB) in BRACKETS:
        while string.count(oB) > string.count(cB):
            string = string + cB
        while string.count(cB) > string.count(oB):
            string = oB + string
        string = quick_clean(string)
    for (oB, cB) in BRACKETS:
        if oB in string and cB in string and string.index(cB) < string.index(oB):
            string = quick_clean(string.replace(cB, '').replace(oB, ''))
    string = re.sub(r'\s*[)}>]\s*', ') ', re.sub(r'\s*[({<]\s*', ' (', string)).strip()
    string = re.sub(r'\s*\]\s*', '] ', re.sub(r'\s*\[\s*', ' [', string)).strip()
    string = string.replace('()', '').replace('{}', '').replace('[]', '').replace('<>', '')
    string = unicodedata.normalize('NFC', string).strip()
    return string


def remove_brackets(string):
    """Function to remove brackets surrounding a string"""
    string = string.strip()
    for (oB, cB) in BRACKETS:
        while len(string) >= 3 and string.startswith(oB) and string.endswith(cB):
            string = string[1:-1]
            string = quick_clean(string)
    string = quick_clean(string)
    return string


def remove_quotes(string):
    """Function to remove quotation marks surrounding a string"""
    string = string.strip()
    while len(string) >= 3 and string.startswith('\'') and string.endswith('\''):
        string = string[1:-1]
        string = quick_clean(string)
    string = quick_clean(string)
    return string


# Functions for parsing strings


def last_word(string):
    """Function to get the last word of a sentence in lower case"""
    try: return (string.rsplit(None, 1)[-1]).lower()
    except: return None


def last_but_one_word(string):
    """Function to get the last but one word of a sentence in lower case"""
    try: return (string.rsplit(None, 2)[-2]).lower()
    except: return None


def change_case(string):
    """Function to replace UPPERCASE words in a string with Sentence Case"""
    if string == '': return ''
    words = string.split()
    for i in range(len(words)):
        upper = True
        for c in words[i]:
            if unicodedata.category(c) == 'Ll': upper = False
        if upper and words[i] not in 'IIIIVIIIIXIII':
            words[i] = words[i].title()
    string = ' '.join(words)
    return string


def change_case_specific(string, words, case='lower'):
    """Function to convert specific words in a string to a specific case"""
    if string == '': return ''
    if not words or len(words) == 0: return string
    if case not in ['lower', 'upper', 'caps']: return string
    try:
        rx = re.compile('\b(' + '|'.join(w.lower() for w in words) + ')\b', flags=re.IGNORECASE)
        if case == 'lower': string = rx.sub(lambda m: m.group(1).lower(), string)
        elif case == 'upper': string = rx.sub(lambda m: m.group(1).upper(), string)
        elif case == 'caps': string = rx.sub(lambda m: m.group(1).capitalise(), string)
        return quick_clean(string)
    except: return string

# Functions to test whether a string has a particular structure


def is_number(string):
    """Function to test whether a string is a float number"""
    try:
        float(string)
        return True
    except ValueError: return False


def is_IAMS_id(string):
    """Function to test whether a string is a valid IAMS record ID"""
    if string is None or not string or string == 'None':
        return False
    if RE_IAMS_ID.fullmatch(string) and string.split('-')[0] in TYPES:
        return True
    return False


def is_year(string):
    """Function to test whether a string looks like a year"""
    string = string.lstrip('[{(').rstrip(']})').strip()
    if is_number(string):
        if len(string) == 4 and string[0] in ['1', '2']: return True
        else: return False
    else: return False


def is_lower_case_letter(string):
    """Function to test whether a string is a single lower-case letter"""
    if not string or string is None: return False
    if len(string) == 1 and string.islower(): return True
    return False


def normalize_dewey(string, escapes=False):
    """Function to normalize a string so that it has the correct structure
     for a Dewey classification, i.e. 3 digits, optionally followed by a
     decimal point then more digits."""
    if string is None or not string: return ''
    string = re.sub(r'[^0-9.\[\]\-]', '', string).rstrip('.')
    if string == '': return ''
    integer, decimal = (string + '.').split('.', 1)
    if decimal != '':
        if escapes:
            decimal = '/?\.' + re.sub(r'([0-9\]])(?![\]\-])', r'\1/?', decimal.rstrip('0')).rstrip('/?')
        else:
            integer = ('000' + integer)[-3:]
            decimal = '.' + decimal.replace('.', '').rstrip('0')
    string = '{}{}'.format(integer, decimal).rstrip('.')
    return string


# ISBN FUNCTIONS

def isbn_10_check_digit(nine_digits):
    """Function to get the check digit for a 10-digit ISBN"""
    if len(nine_digits) != 9: return None
    try: int(nine_digits)
    except: return None
    remainder = int(sum((i + 2) * int(x) for i, x in enumerate(reversed(nine_digits))) % 11)
    if remainder == 0: tenth_digit = 0
    else: tenth_digit = 11 - remainder
    if tenth_digit == 10: tenth_digit = 'X'
    return str(tenth_digit)


def isbn_13_check_digit(twelve_digits):
    """Function to get the check digit for a 13-digit ISBN"""
    if len(twelve_digits) != 12: return None
    try: int(twelve_digits)
    except: return None
    thirteenth_digit = 10 - int(sum((i % 2 * 2 + 1) * int(x) for i, x in enumerate(twelve_digits)) % 10)
    if thirteenth_digit == 10: thirteenth_digit = '0'
    return str(thirteenth_digit)


def isbn_10_check_structure(isbn10):
    """Function to check the structure of a 10-digit ISBN"""
    return True if re.match(RE_ISBN10, isbn10) else False


def isbn_13_check_structure(isbn13):
    """Function to check the structure of a 13-digit ISBN"""
    return True if re.match(RE_ISBN13, isbn13) else False


def is_isbn_10(isbn10):
    """Function to validate a 10-digit ISBN"""
    isbn10 = re.sub(r'[^0-9X]', '', isbn10.replace('x', 'X'))
    if len(isbn10) != 10: return False
    return False if isbn_10_check_digit(isbn10[:-1]) != isbn10[-1] else True


def is_isbn_13(isbn13):
    """Function to validate a 13-digit ISBN"""
    isbn13 = re.sub(r'[^0-9X]', '', isbn13.replace('x', 'X'))
    if len(isbn13) != 13: return False
    if isbn13[0:3] not in ('978', '979'): return False
    return False if isbn_13_check_digit(isbn13[:-1]) != isbn13[-1] else True


def isbn_convert(isbn10):
    """Function to convert a 10-digit ISBN to a 13-digit ISBN"""
    if not is_isbn_10(isbn10): return None    
    return '978' + isbn10[:-1] + isbn_13_check_digit('978' + isbn10[:-1])

# Miscellaneous other useful functions


def add_string(string, base, separator, brackets=False):
    """Function to append a string to another string"""
    if string != '' and string not in base:
        if base != '': base = base + separator
        if brackets: string = '( {} )'.format(string)
        base += string
    return base


def get_boolean(prompt):
    """Function to prompt the user for a Boolean value"""
    while True:
        try: return {'Y': True, 'N': False}[input(prompt).upper()]
        except KeyError:
            print('Sorry, your choice was not recognised. Please enter Y or N:')


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


# START OF MORE SPECIFIC FUNCTIONS


def repair_accents_in_place_names(string):
    """Function to repair missing accents in place names"""
    string = mrx.PlaceNamesAccents().sub(string)
    string = quick_clean(string)
    return string

# FUNCTIONS TO GET PARTS OF NAMES, TOPICS ETC


def clean_for_date_parse(string):
    if string == '': return ''
    string = string.lower()
    # remove leading zeros from numbers
    string = re.sub(r'\s+0+(?=[1-9])', ' ', string)
    # replace ' as year abbreviation
    string = re.sub(r"'([0-9]{2})(?![0-9])", r"19\1", string)

    # replace numerically formatted dates
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?1[.]([12][0-9]{3})(?![0-9])', r' \1 january \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?2[.]([12][0-9]{3})(?![0-9])', r' \1 february \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?3[.]([12][0-9]{3})(?![0-9])', r' \1 march \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?4[.]([12][0-9]{3})(?![0-9])', r' \1 april \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?5[.]([12][0-9]{3})(?![0-9])', r' \1 may \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?6[.]([12][0-9]{3})(?![0-9])', r' \1 june \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?7[.]([12][0-9]{3})(?![0-9])', r' \1 july \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?8[.]([12][0-9]{3})(?![0-9])', r' \1 august \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]0?9[.]([12][0-9]{3})(?![0-9])', r' \1 september \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]10[.]([12][0-9]{3})(?![0-9])', r' \1 october \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]11[.]([12][0-9]{3})(?![0-9])', r' \1 november \2 ', string)
    string = re.sub(r'(?<![0-9])([0-9][0-9]?)[.]12[.]([12][0-9]{3})(?![0-9])', r' \1 december \2 ', string)

    string = quick_clean(string)
    return string


def get_year(string, defaultyear=''):
    if string == '': return ''
    year = ''
    string = clean_for_date_parse(string)

    if RE_YEAR_POST_1500.search(string) is not None:
        yearlist = RE_YEAR_POST_1500.findall(string)
        year = RE_YEAR_POST_1500.search(yearlist[-1]).group(1)
    if year == '' and defaultyear != '': return defaultyear
    return year


def get_month(string):
    if string == '': return ''
    string = clean_for_date_parse(string)

    if any(s in string for s in ['jan', 'enero', 'ionawr', 'ocak', 'stycz']):
        return 'January'
    if any(s in string for s in ['feb', 'chwe', 'fev']):
        return 'February'
    if any(s in string for s in ['mar', 'mawrth', u'mu\0072rz']):
        return 'March'
    if any(s in string for s in ['apr', 'avr', 'ebrill', 'kwie']):
        return 'April'
    if any(s in string for s in ['may', 'mai', 'mai︠a︡', 'maj']):
        return 'May'
    if any(s in string for s in ['jun', 'czerw', 'ii︠u︡ni︠a︡', 'ioun', 'juin', 'meh']):
        return 'June'
    if any(s in string for s in ['jul', 'gorff', 'juil']):
        return 'July'
    if any(s in string for s in ['aug', 'agosto', 'août', 'awst', 'sierp']):
        return 'August'
    if any(s in string for s in ['sep', 'medi', 'rugsėjo', 'wrzes']):
        return 'September'
    if any(s in string for s in ['oct', 'hydref', 'pazdzier']):
        return 'October'
    if any(s in string for s in ['nov', 'listopad', 'tachw']):
        return 'November'
    if ('dec' in string and 'dechreuodd' not in string) or any(s in string for s in ['dek', 'grudz', 'rhagfyr']):
        return 'December'
    if any(s in string for s in ['spring', 'gwanwyn']):
        return 'Spring'
    if any(s in string for s in ['summer', 'yr haf']):
        return 'Summer'
    if any(s in string for s in ['autumn']):
        return 'Autumn'
    if any(s in string for s in ['winter', 'gaeaf']):
        return 'Winter'
    return ''


def get_date(string):
    if string == '': return ''
    date = ''
    string = clean_for_date_parse(string)

    if RE_DATE_1.search(string) is not None:
        date = str(RE_DATE_1.search(string).group(2)).strip('/')
    elif RE_DATE_2.search(string) is not None:
        date = str(RE_DATE_2.search(string).group(1)).strip('/')
    date = date.replace('/', '-')
    return date


def get_volume(string):
    if string == '': return ''
    volume = ''
    string = clean_for_date_parse(string)
    if RE_VOLUME.search(string) is not None:
        volume = str(RE_VOLUME.search(string).group(2))
    return volume


def get_issue(string):
    if string == '': return ''
    issue = ''
    string = clean_for_date_parse(string)
    if RE_ISSUE.search(string) is not None:
        issue = str(RE_ISSUE.search(string).group(2))
    return issue


def get_number(string):
    if string == '': return ''
    number = ''
    string = clean_for_date_parse(string)
    if RE_NUMBER.search(string) is not None:
        number = str(RE_NUMBER.search(string).group(2))
    return number


def get_date_parts(string, defaultyear=''):
    if string is None or not string or string == '': return '', '', '', '', '', ''

    year = get_year(string, defaultyear)
    month = get_month(string)
    date = get_date(string)
    volume = get_volume(string)
    issue = get_issue(string)
    number = get_number(string)

    return year, month, date, volume, issue, number


def get_date_parts_as_string(string, defaultyear=''):
    year, month, date, volume, issue, number = get_date_parts(string, defaultyear)
    chronology = '%(year)s %(month)s %(date)s' % locals()
    chronology = quick_clean(chronology)
    enumeration = ''
    if volume != '': enumeration += ' volume ' + volume
    if issue != '': enumeration += ' issue ' + issue
    if number != '': enumeration += ' number ' + number
    enumeration = quick_clean(enumeration)
    if enumeration != '':
        chronology = '%(chronology)s (%(enumeration)s)' % locals()
    chronology = remove_brackets(chronology)
    return chronology


def get_date_range(string, default_start_year='', default_end_year=''):
    if string == '': return ''
    string = string.replace('-', ' FROM ; TO ')
    string = string.replace(' to ', ' TO ')
    string = string.replace(' from ', ' FROM ')
    string = re.sub(r'(began|commence|launch|start)[ed]*\s*(at|in|with)?', ' ; FROM ', string, flags=re.IGNORECASE)
    string = re.sub(r'(cease|end|finish)[ed]*\s*(with|in|on)?', ' ; TO ', string, flags=re.IGNORECASE)
    start, end, full = '', '', ''
    for sub_range in string.split(';'):
        if re.sub(r'[^0-9]', '', sub_range) != '':
            if ' FROM ' in sub_range:
                start = get_date_parts_as_string(sub_range, default_start_year)
            elif ' TO ' in sub_range:
                end = get_date_parts_as_string(sub_range, default_end_year)
            else:
                full = get_date_parts_as_string(sub_range, default_start_year)
    if start != '' or end != '':
        full = start + ' - ' + end
    full = full.strip()
    return full


def get_frequency(string):
    if string == '': return ''
    string = remove_brackets(string.lower())

    # Remove anything in brackets
    string = re.sub(r'\([^)]*\)', '', string).replace(' and ', ' & ')

    # Replace text with numbers
    string = re.sub(r'(?<![a-z])twenty-*four(?![a-z])', '24', string)
    string = re.sub(r'(?<![a-z])one(?![a-z])', '1', string)
    string = re.sub(r'(?<![a-z])two(?![a-z])', '2', string)
    string = re.sub(r'(?<![a-z])three(?![a-z])', '3', string)
    string = re.sub(r'(?<![a-z])four(?![a-z])', '4', string)
    string = re.sub(r'(?<![a-z])five(?![a-z])', '5', string)
    string = re.sub(r'(?<![a-z])six(?![a-z])', '6', string)
    string = re.sub(r'(?<![a-z])seven(?![a-z])', '7', string)
    string = re.sub(r'(?<![a-z])eight(?![a-z])', '8', string)
    string = re.sub(r'(?<![a-z])nine(?![a-z])', '9', string)
    string = re.sub(r'(?<![a-z])ten(?![a-z])', '10', string)
    string = re.sub(r'(?<![a-z])eleven(?![a-z])', '11', string)
    string = re.sub(r'(?<![a-z])twelve(?![a-z])', '12', string)

    # Test if the string starts with a number and ends with 'year'
    if ' ' in string and is_number(string.split(' ')[0]) and any(
                    s in string for s in ['year', 'annual', 'annum', 'annuel']):
        if string.split(' ')[0] == '2': return 'Semi-annual'
        if string.split(' ')[0] == '3': return 'Tri-annual'
        string = string.split(' ')[0] + ' times a year'
        string = quick_clean(string)
        return string

    if any(s in string for s in ['daily']):
        if any(s in string for s in ['monday to friday', 'except saturday & sunday']):
            return 'Daily (except Saturday & Sunday)'
        if any(s in string for s in ['monday to saturday', 'except sunday']):
            return 'Daily (except Sunday)'
        if any(s in string for s in ['except monday']):
            return 'Daily (except Monday)'
        return 'Daily'
    if any(s in string for s in ['3', 'three', 'tri-', 'triweek', 'thrice']) and 'week' in string:
        return 'Tri-weekly'
    if any(s in string for s in ['2', 'two', 'twice', 'semi']) and 'week' in string:
        return 'Semi-weekly'
    if any(s in string for s in ['biweekl', 'bi-weekl']):
        return 'Bi-weekly'
    if any(s in string for s in ['wythnosol', 'weekl', 'weeky']):
        return 'Weekly'
    if any(s in string for s in ['fortnight']):
        return 'Fortnightly'
    if any(s in string for s in ['3', 'three', 'tri-', 'trimonth', 'thrice']) and 'month' in string:
        return 'Tri-monthly'
    if any(s in string for s in ['2', 'two', 'twice', 'semi']) and 'month' in string:
        return 'Semi-monthly'
    if any(s in string for s in ['bi-month', 'bimonth', 'bi-a month']):
        return 'Bi-monthly'
    if any(s in string for s in ['monthly']):
        return 'Monthly'
    if any(s in string for s in ['quarter']):
        return 'Quarterly'
    if any(s in string for s in ['3', 'three', 'tri-', 'triannu', 'thrice']) and any(
                    s in string for s in ['year', 'annual', 'annum', 'annuel']):
        return 'Tri-annual'
    if any(s in string for s in ['2', 'two', 'twice', 'semi']) and any(
                    s in string for s in ['year', 'annual', 'annum', 'annuel']):
        return 'Semi-annual'
    if any(s in string for s in ['annual', 'annuel']):
        return 'Annual'
    if any(s in string for s in ['biennial', 'bi-ennial', 'biannual', 'bi-annual', 'bienal']):
        return 'Biennial'
    if any(s in string for s in ['triennial', 'tri-ennial']):
        return 'Triennial'
    if 'quadrennial' in string:
        return 'Quadrennial'
    if 'quinquennial' in string:
        return 'Quinquennial'
    if 'irregul' in string:
        return 'Irregular'
    if 'varies' in string:
        return 'Frequency varies'

    return ''


def get_holdings_years(string, startyear='', endyear=''):
    """Functions to get first and last year held"""
    start, end = '', ''
    string = string.lower()
    string = RE_VOLUME.sub('', string)
    string = RE_ISSUE.sub('', string)
    string = RE_NUMBER.sub('', string)

    # check if section contains a year number
    if RE_YEAR_POST_1500.search(string) is not None:
        yearlist = RE_YEAR_POST_1500.findall(string)
        y = []
        for item in yearlist:
            if not (startyear != '' and int(item) < int(startyear)) and not (endyear != '' and int(item) > int(endyear)):
                y.append(item)
        if len(y) > 0:
            start, end = sorted(y)[0], sorted(y)[-1]

        # remove everything other than numbers, hyphens, semi-colons and square brackets
        string = re.sub(r'[^0-9\-]', '', string).strip()

        if string[-1] in '-': end = 'Continuing'

    return start, end


def get_name_parts(field):
    """Function to get name parts from a name field"""
    name, dates, ntype, role, isni, viaf = '', '', '', set(), set(), set()
    if field.tag == '880' and '6' in field:
        field.tag = (field['6'][:3])

    # Name
    for subfield in field:
        code, content = subfield[0], clean(subfield[1], hyphens=False)
        if (field.tag[1] == '0' and code in ['a', 'b', 'c']) \
                or (field.tag[1] == '1' and code in ['a', 'b', 'c', 'f', 'g', 'n', 'p']):
            content = change_case(content)
            # Remove brackets
            if code in ['a', 'c']: content = remove_brackets(content)
            # Test if name part contains a relator term
            if code == 'c' and get_relators(content):
                role = role.union(get_relators(content))
            # Test if name part contains dates
            elif code in ['c', 'f', 'g', 'n', 'p'] and content.lower().startswith(('fl.', 'b.', 'd.')):
                content = clean_name_dates(content)
                if re.sub(r'[^0-9]', '', content) != '': dates = content
            else:
                if code == 'c': content = clean_words_associated_with_name(content)
                name = add_string(content, name, ', ').replace(', (', ' (')
        # Fuller form of name
        elif code == 'q':
            name = add_string(content, name, ' ' if field.tag[1] == '0' else ', ')
    # Check brackets
    name = remove_brackets(check_brackets(name))
    # Replace missing full stops after initials
    name = re.sub(r'([\s.\-][A-Z])([,\s]|$)', r'\1.\2', name)

    # Dates
    for subfield in field.get_subfields('d'):
        subfield = clean_name_dates(subfield)
        if RE_YEAR.search(subfield) is not None:
            dates = subfield

    # Type
    ntype = {
        '00': 'person',
        '10': 'organisation',
        '11': 'meeting/conference',
    }[field.tag[1:]]

    # Role
    for subfield in field.get_subfields('e'):
        if get_relators(subfield):
            role = role.union(get_relators(subfield))

    # ISNI
    # VIAF
    for subfield in field.get_subfields('8', '9', cleaning=False):
        subfield = re.sub(r'\s+', ' ', subfield.replace('|', '')).strip()
        if 'http://isni.org/isni/' in subfield:
            isni.add(subfield)
        elif 'http://viaf.org/viaf/' in subfield:
            viaf.add(subfield)

    role = ' ; '.join(sorted(str(s) for s in role if s != ''))
    isni = ' ; '.join(sorted(str(s) for s in isni if s != ''))
    viaf = ' ; '.join(sorted(str(s) for s in viaf if s != ''))

    return name, dates, ntype, role, isni, viaf


def get_series_parts(field):
    title, number = '', ''
    for subfield in field.get_subfields('a'):
        subfield = clean_490(subfield)
        if ';' in subfield:
            t_title, t_number = subfield.split(';', 1)
        elif not ('v' in field) and RE_SERIES_NUMBER.search(subfield):
            t_title = re.sub(RE_SERIES_NUMBER.search(subfield).group(0), '', subfield)
            t_number = RE_SERIES_NUMBER.search(subfield).group(0)
        else:
            t_title = subfield
            t_number = ''
        t_title = clean_490(t_title)
        t_number = clean_490_number(t_number)
        title = add_string(t_title, title, ' ')
        number = add_string(t_number, number, ' ')

    for subfield in field.get_subfields('v'):
        subfield = clean_490_number(subfield)
        if re.sub(r'[^0-9]', '', subfield) != '':
            number = add_string(subfield, number, ' ')

    return title, number


def get_relators(string):
    if string == '': return False
    rels = set()
    for substring in re.split(r'[,:;]| and ', string):
        substring = quick_clean(substring)
        if len(substring) > 0:
            if len(substring) == 3:
                l = relators.get(substring)
                if l: rels.add(l)
                continue
            substring = mrx.Relators().sub(substring)
            if substring != '':
                rels.add(substring)
    if len(rels) == 0: return False
    return rels


def get_topic_parts(field):
    """Function to get topic parts from a subject field"""
    term, ttype, genre = '', '', set()
    if field.tag == '880' and '6' in field:
        field.tag = (field['6'][:3])

    # Term
    for subfield in field:
        code, content = subfield[0], clean(subfield[1].strip('*'))
        if code in ['a', 'b', 'p', 'v', 'x', 'y', 'z']:
            content = remove_brackets(content)
            # Replace missing full stops after initials
            if field.tag == '600' and code == 'a':
                content = re.sub(r'([\s.\-][A-Z])([,\s]|$)', r'\1.\2', content)
            term = add_string(content, term, '--')
        elif code == ['c', 'd', 'e', 'n', 'q']:
            content = content.replace('B.C', 'B.C.').replace('A.D', 'A.D.').replace('..', '.')
            term = add_string(content, term, ', ' if code == 'd' else ' ')

    term = repair_accents_in_place_names(term)
    term = re.sub(r'[$][a-z\s]', '--', term).replace('----', '--')

    # Type
    if re.sub(r'[^0-9-]', '', term) == term:
        ttype = 'chronological term'
    else:
        try:
            ttype = {
                '00': 'person',
                '10': 'organisation',
                '11': 'meeting/conference',
                '30': 'title',
                '51': 'geographical term',
            }[field.tag[1:]]
        except:
            ttype = 'general term'

    # Genre
    for subfield in field.get_subfields('v'):
        subfield = clean_genre(subfield)
        genre.add(subfield)

    return term, ttype, genre


# FUNCTIONS TO EXPAND ABBREVIATIONS

def expand_abbreviations(string, plurals=True, case=True):
    if string == '': return ''

    # Expand abbreviations which span more than one word
    string = re.sub(r'(?<![a-z])n\.*\s*s(er)?\.*(?![a-z])', 'new series', string, flags=re.IGNORECASE)
    # French
    string = re.sub(r' et augm(ent)?(?![a-z])\.*', u' et augment\u00e9e', string, flags=re.IGNORECASE)
    string = re.sub(r' et corr(ig)?(?![a-z])\.*', u' et corrig\u00e9e', string, flags=re.IGNORECASE)
    string = re.sub(r'corr(ig)?(\u00e9e)?\.* et ', u'corrig\u00e9e et ', string, flags=re.IGNORECASE)
    string = re.sub(r'r(ev)?\.* et ', 'revue et ', string, flags=re.IGNORECASE)
    # German
    string = re.sub(r'(?<![a-z])n\.*\s*f\.*(?![a-z])', 'neue Folge', string, flags=re.IGNORECASE)
    # Italian
    string = re.sub(r'nuova ser(?![a-z])', 'nuova serie', string, flags=re.IGNORECASE)
    # Spanish
    string = re.sub(r'correg\.*\s*y\s*aum\.*', 'corregida y aumentada', string, flags=re.IGNORECASE)

    # Expand single-word abbreviations
    words = re.split('([\w\-]+\.*)', string)
    for i, word in enumerate(words):
        words[i] = mrx.Abbreviations().sub(words[i])
        if word != '' and case:
            if word.isupper():
                words[i] = words[i].upper()
            elif word[0].isupper():
                words[i] = words[i].capitalize()
        if words[i] in ['numbers', 'volumes', 'parts'] and not plurals:
            words[i] = words[i].rstrip('s')

    # Re-join string
    string = quick_clean(''.join(words))
    return string


def expand_place_abbreviations(string, ctrys=None):
    if string.lower() in 's.n. s. n. sn s n s.l. s. l. sl s l s.i. si s i nv n.v. n. v. n v blnpn': return ''
    if string.lower() == 'united states': return 'United States of America'
    string = string.replace('$42blnpn', '').replace('42blnpn', '')
    string = mrx.PlaceNamesUK().sub(string)
    string = mrx.PlaceNamesUS().sub(string)
    if countries:
        if 'Australia' in ctrys:
            string = mrx.PlaceNamesAustralia().sub(string)
        if 'Brazil' in ctrys:
            string = mrx.PlaceNamesBrazil().sub(string)
        if 'Canada' in ctrys:
            string = mrx.PlaceNamesCanada().sub(string)
        if 'New Zealand' in ctrys:
            string = mrx.PlaceNamesNewZealand().sub(string)
    string = mrx.PlaceNamesOther().sub(string)
    string = re.sub(r'[nN]ew[\-\s]*[yY]ork\s*(\(?,?\s*(NY|New York|City)\)?)?', 'New York', string)
    string = re.sub(r'(\bin the )?\bcounty of\b', '', string)
    string = re.sub(
        r'[,\s]+(U[\.\s]*K[\.\s]*|U[\.\s]*S[\.\s]*A*[\.\s]*|United\s*Kingdom|United\s*States(\s*of\s*America)?)\s*$',
        '', string, flags=re.IGNORECASE)
    string = string.replace('Saint Christopher - Nevis', 'Saint Kitts-Nevis')
    string = quick_clean(string).strip('?')
    return string

# FUNCTIONS FOR CLEANING SPECIFIC FIELDS


def clean_250(string):
    if string == '': return ''
    string = clean(string.rstrip('!-'))  # full clean IS required here
    if string.lower() in 'author all edition microfilm all editions microfilm':
        return ''

    # Remove characters that aren't needed
    string = quick_clean(re.sub(r'\s*[:;\]\/]+', ',', re.sub(r'[\[<>{}*]', '', string)))
    # Replace & with and
    string = quick_clean(re.sub(r'\s*&[aceimst\/]{1,3}\.*\s*', ' et cetera ', string, flags=re.IGNORECASE))
    if string.lower().endswith(' et cetera'): string = quick_clean(string[:-10])
    string = string.replace('&', ' and ')
    # Remove 'the' from start of string
    string = re.sub(r'^the\s+', '', string, flags=re.IGNORECASE)
    # Replace ellipsis with comma
    string = re.sub(r'\s*\.\.\.+\s*', ', ', string)
    # Add space between numbers and letters
    string = re.sub(r'[\s,]+i\.*\s*e[\.\s]+', ', that is ', string, flags=re.IGNORECASE)
    string = string.replace('U.K.', 'UK')
    string = re.sub(r'([a-z])\.*([0-9])', r'\1 \2', string)
    string = re.sub(r'([a-z])\.+([a-z])', r'\1. \2', string)

    # Known spelling mistakes
    string = re.sub(r'new[.,]+\s*ed', r'new ed', string, flags=re.IGNORECASE)
    string = string.replace('Reprograf. Nachdr. d.', 'Reprografischen Nachdruck der')
    string = re.sub(r'(?<![a-z])stere?o-*(typed)?[\.\s]+', 'stereotyped ', string, flags=re.IGNORECASE)
    string = string.replace('Unifrom', 'Uniform')

    string = quick_clean(string)

    string = re.sub(r'(?<=[0-9])(th| )ed(ition)?\.*', 'th edition', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![01])1(st[.,]*|\s*a[.,\s]+)\s*', '1st ', string, flags=re.IGNORECASE)
    string = re.sub(r'^([0-9]{0,2}1)[.][.,]*\s*', r'\1st ', string, flags=re.IGNORECASE)
    string = re.sub(r'2((nd|gn)[.,\s]+|[ad](?![a-z])[.,]*)\s*', '2nd ', string, flags=re.IGNORECASE)
    string = re.sub(r'^([0-9]{0,2}2)[.][.,]*\s*', r'\1nd ', string, flags=re.IGNORECASE)
    string = re.sub(r'3(rd[.,]|\s*(am|rda|te|[ad])(?![a-z])[.,]*)\s*', '3rd ', string, flags=re.IGNORECASE)
    string = re.sub(r'^([0-9]{0,2}3)[.][.,]*\s*', r'\1rd ', string, flags=re.IGNORECASE)
    string = re.sub(r'([4-9])(th[.,]|\s*the(?![a-z])[.,]*)\s*', r'\1th ', string, flags=re.IGNORECASE)
    string = re.sub(r'^([0-9]{0,2}4-9])[.][.,]*\s*', r'\1th ', string, flags=re.IGNORECASE)
    string = re.sub(r'^([0-9]{1,3})\.', r'\1th ', string, flags=re.IGNORECASE)
    string = re.sub(r'([0-9])(st|gn|nd|rd|th)(?![a-z])[.,]\s*', r'\1\2 ', string, flags=re.IGNORECASE)
    string = re.sub(r'([0-9])-?o?e(?![a-z])[.,]*\s*', r'\1e ', string, flags=re.IGNORECASE)
    string = expand_abbreviations(string)
    words = ('a', 'another', 'augmented', 'by', 'complete', 'critical',
             'edition', 'edited', 'editor', 'editors',
             'further', 'illustrated', 'imprinted', 'international', 'limited', 'notes',
             'prepared', 'printed', 'reprinted', 'revised', 'series', 'special', 'the', 'with')
    string = change_case_specific(string, words, case='lower')
    string = string.replace(' edited by ', ', edited by')

    if string.lower() in ['edition', 'u']: return ''
    string = remove_brackets(check_brackets(quick_clean(string.rstrip('!-'))))
    return string


def clean_26X(string):
    if string == '': return ''
    # Remove rubbish values
    if 'not identified' in string.lower(): return ''
    string = quick_clean(string)
    if string.lower() in 'b.i. b. i. bi b i b.m. b. m. bm b m s.n. s. n. sn s n s.l. s. l. sl s l s.i. si s i nv n.v. n. v. n v np n.p. n. p. n p anno domini imprinted in the year': return ''
    # Remove brackets
    if len(string) >= 3: remove_brackets(string)
    # Remove information about pagination and volume numbers
    string = quick_clean(re.sub(r'^((ff|p(p|ages?)|vol)[.\s]+[0-9,.\-\sivxfp\[\]]+\b\s*|[0-9,.\-\sivxfp\[\]]+\b\s*p(p|ages?))', '', string))
    # Remove information about size
    string = quick_clean(re.sub(r'\s*\b[0-9]+\s*cm[.\s]*$', '', string))
    # Remove known problems from the start of the string
    string = quick_clean(re.sub(r'^(at\s+|a paris)(?!press)', '', string))
    # Remove known problems from the end of the string
    string = quick_clean(re.sub(r'\s+(etc|anno)$', '', string))
    return string


# Retaining brackets?
def clean_300(field):
    description = ''
    for subfield in field:
        code, content = subfield[0], clean(subfield[1].lower())
        if code in ['a', 'b', 'c', 'e', 'f', 'g', '3'] and content not in ['cm', 'p. cm'] \
                and not (any(s in content for s in ['jaggard', 'london', 'macmillan'])):
            sub_desc = ''
            # Remove characters that aren't needed
            content = re.sub(r'[:;]+', ',', re.sub(r'[\[\]<>{}*]', '', content))
            # Replace & with and
            content = re.sub(r'\s*&[aceimsts\/]{1,3}\.*\s*', ' et cetera ', content)
            content = content.replace('&', ' and ')
            # Replace combining ring above (u030a) masculine ordinal indicator (u00ba) and superscript zero (u2070) with degree (u00b0)
            content = re.sub(u'[.\s]*[\u030a\u00ba\u2070\u00b0]+m*', u'\u00b0', content)
            content = re.sub(u'(?<=[0-9])\s*(mo|o|to|vo)(?![a-z])', u'\u00b0', content)
            # Add space between numbers and letters
            content = content.replace('i.e.', ', that is ')
            content = re.sub(r'([0-9])([a-z])', r'\1 \2', content)
            content = re.sub(r'([a-z])([0-9])', r'\1 \2', content)
            # Add space after full stops after letters
            content = re.sub(r'([a-z])\.', r'\1. ', content)
            # Check space around brackets, commas and hyphens
            content = re.sub(r'\s*-\s*', '-', re.sub(r'\s*,\s*', ', ', re.sub(r'\s*\)\s*', ') ', re.sub(r'\s*\(\s*', ' (', content))))
            content = content.replace('front.', 'frontispiece')
            content = re.sub(r'b(lack)?\.*\s*(and)?/*\s*w(hite)?\.*', 'black and white', content)
            # Split into words
            for item in re.split(r'\s+', content):
                oB, cB, cP = '', '', ''
                if item.endswith(','): cP = ','
                item = quick_clean(item)
                if item.startswith('('): oB = '('
                if item.endswith(')'): cB = ')'
                item = quick_clean(item.strip('()'))
                if item == 'v' and oB != '(':
                    item = 'volumes'
                elif item == 'p':
                    item = 'pages'
                elif item == 'cd':
                    item = 'CD'
                elif item == 'dvd':
                    item = 'DVD'
                elif item == 'ill':
                    item = 'illustrations'
                elif not re.fullmatch(RE_NUMERAL, item):
                    item = quick_clean(mrx.Abbreviations().sub(item))
                sub_desc += ' ' + oB + item + cB + cP
            # If pages appears before numeration, move it afterwards
            sub_desc = re.sub(r'^\s*pages ([0-9\-]+),*', r'\1 pages,', sub_desc)
            description = add_string(quick_clean(sub_desc), description, ', ')
    description = quick_clean(description)
    # Final cleaning
    description = re.sub(r'(?<![0-9])1 ([a-z]+[a-tv-z])s(?![a-z])', r'1 \1', description)
    description = description.replace(', and', ' and').replace(', (', ' (').replace(', of ', ' of ')
    description = description.replace('compact disc', 'CD')
    description = re.sub(r'4\s*((sup)?\s*3/(sub)?\s*4|\.75)', '4 3/4', description)
    description = description.replace('en colour', '(colour)')
    description = description.replace('some of which are in colour', '(some colour)')
    description = description.replace('general table', 'genealogical table')
    description = description.replace('loose leaf', 'loose-leaf')
    description = description.replace('min score', 'miniature score')
    description = re.sub(r'(?<![a-z])n s(?![a-z])', 'new series', description)
    description = description.replace('no pagination provided', '(unpaged)')
    description = description.replace('volume unpaged', 'volume (unpaged)')
    description = description.replace('wood engraved', 'wood engravings')
    description = re.sub(r'(?<![a-z])s sheet', 'single sheet', description)

    if len(description) < 5: return ''
    description = quick_clean(description)
    return description


def clean_310(string):
    if string == '': return ''
    string = string.replace(';', ',')
    string = quick_clean(string)
    terms = set()
    for item in string.split(','):
        frequency = get_frequency(item)
        if frequency != '':
            terms.add(frequency)
    return ' ; '.join(terms)


def clean_362(field, start_year='', end_year=''):
    date_range = ''
    for subfield in field.get_subfields('a', cleaning=False):
        date_range = add_string(subfield, date_range, ' ; ')

    date_range = get_date_range(date_range, start_year, end_year)
    return date_range


def clean_490(string):
    if string == '': return ''
    string = quick_clean(string.lstrip('$.,:;/\-[])} ').rstrip('.,:;/\-[]({ '))
    string = clean(re.sub(r'bibl\.*[\s]*pp?\.*[0-9\-]+$', '', string))
    if len(string) >= 3:
        string = remove_brackets(string)
    string = expand_abbreviations(string)
    return string


def clean_490_number(string):
    if string == '': return ''
    string = re.sub(r'(?<![a-z])u[.\s]*p[.\s]*(?![a-z])', 'UP', string, flags=re.IGNORECASE)
    string = expand_abbreviations(string, plurals=False, case=False)
    string = re.sub(r'^t[.\s]+', 'tome ', string, flags=re.IGNORECASE)
    string = re.sub(r'^n[.\s]+', 'number ', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])v[.\s]+', 'volume ', string, flags=re.IGNORECASE)
    # Remove space around hyphens
    string = re.sub(r'\s*-\s*', '-', string)
    # Space between letters and numbers
    string = re.sub(r'([0-9])([a-z])', r'\1 \2', string)
    string = re.sub(r'([a-z])([0-9])', r'\1 \2', string)
    # Replace full stop between parts of numbers with comma
    string = re.sub(r'([0-9])\.\s+([a-z])', r'\1, \2', string)
    string = quick_clean(string)
    return string


def clean_500(string):
    if string.lower() in ['formerly cip', 'formerley cip']: return ''
    string = expand_abbreviations(string, plurals=False, case=False)
    # string = string.replace(' ed.', ' edition ').replace('edition  ', 'edition ')
    string = re.sub(r'(?<![a-z])[sS]ig\.?(?![a-z])', 'signatures', string)
    string = quick_clean(string)
    return string


def clean_510(string):
    string = string.replace(' ed.', ' edition').replace('edition  ', 'edition ')
    string = string.replace(' bibl p', ' bibliography p')
    string = string.replace(' vol.', ' volume')
    string = re.sub(r'[cC][dD]-*[rR][oO][mM]', 'cd-rom', string)
    string = quick_clean(string)
    return string


def clean_852(string):
    if any(s in string.lower() for s in
           ['available', 'availalbe', 'british museum', 'catalog', 'classmark', 'customer service', 'discard',
            'holding', 'impression', 'lacking', 'libraries', 'not found', 'on order', 'pending', 'pressmark',
            'shelmark', 'shelfmrk', 'shelfmark', 'shelfmrak', 'smk', 'spacer', 'stained', 'superseded',
            'test', 'tightly bound', 'unavailable']):
        return ''
    string = string.replace('?', '').replace('for hard copy', '')
    if string.lower() in 'conf id conf index post id conf index eld digital store':
        return ''
    return string


def clean_genre(string):
    string = re.sub(r'[^a-z0-9\s]', '', string.lower())
    string = mrx.Genres().sub(string)
    return string


def clean_words_associated_with_name(string):
    if 'selection' in string.lower() or 'part ' in string.lower() or 'no.' in string.lower():
        return ''
    if ' ' in string and string.lower().split(' ', 1)[0] in ['and']:
        return ''
    if ' ' not in string: string = string.lower()
    # string = re.sub(r'(?<![a-z])[aA]uth?(o|eu)r','author',string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[bB]art\.*(?![a-z])', 'Baronet', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[bB]aron', 'Baron', string, flags=re.IGNORECASE)
    # string = re.sub(r'(?<![a-z])[cC]ivil engineer','civil engineer',string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[cC]urate', 'curate', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[eE]arl', 'Earl', string, flags=re.IGNORECASE)
    # string = re.sub(r'(?<![a-z])[eE]ditor','editor',string, flags=re.IGNORECASE)
    # string = re.sub(r'(?<![a-z])[eE]d\.*(?![a-z])','editor',string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[eE]xpression', '', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[hH]on\.*(?![a-z])', 'Honourable', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[hH]ungarian', 'Hungarian', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[iI]ssui?ng [bB]ody', 'issuing body', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[lL]ady', 'Lady', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[lL]ord', 'Lord', string, flags=re.IGNORECASE)
    # string = re.sub(r'(?<![a-z])[mM]athematics [tT]eacher','mathematics teacher',string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[mM]inister', 'minister', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[pP]reacher', 'preacher', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[mM]rs', 'Mrs', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[rR]ev\.*(?![a-z])', 'Reverend', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[rR]t\.*(?![a-z])', 'Right', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[sS]chool[\-\s]*master', 'schoolmaster', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[sS]ir', 'Sir', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[sS]tudent', 'student', string, flags=re.IGNORECASE)
    # string = re.sub(r'(?<![a-z])[tT]eacher','teacher',string, flags=re.IGNORECASE)
    # string = re.sub(r'(?<![a-z])[tT]r\.*(?![a-z])','translator',string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[vV]icar', 'vicar', string, flags=re.IGNORECASE)
    string = re.sub(r'(?<![a-z])[vV]iscount', 'Viscount', string, flags=re.IGNORECASE)
    # string = re.sub(r'(?<![a-z])[wW]riter','writer',string, flags=re.IGNORECASE)

    string = string.replace(".'", "'").replace('.]', ']').replace(' and ', ', ').replace('Mrs.', 'Mrs').replace(
        'Of ', 'of ').replace('The ', 'the ')

    string = quick_clean(string)
    return string


def clean_name_dates(string):
    string = string.strip().lstrip('.:,;/()[]! ').rstrip('.:,;/()[]! ').replace(';', ' ')
    # B.C. and A.D.
    string = re.sub(r'B\.?C\.?', 'BC', string)
    string = re.sub(r'A\.?D\.?', 'AD', string)
    # century
    string = re.sub(r'cent(?!u)', 'century', string)
    string = re.sub(r'([0-9]+) century', r'\1th century', string)
    string = re.sub(r'([0-9]+)th\. century', r'\1th century', string)
    string = re.sub(r'([0-9]+)-([0-9]+)th century', r'\1th century-\2th century', string)
    string = re.sub(r'\(([0-9]+)th\.?\)', r'(\1th century)', string)
    # approximately
    string = re.sub(r'ca?[.,]*\s*([0-9]+)', r'approximately \1', string, flags=re.IGNORECASE)
    string = re.sub(r'([0-9]{4})\s*\?', r'approximately \1', string, flags=re.IGNORECASE)
    # b. -> -
    # ne -> -
    string = re.sub(r'^(b|n.)[.,]*\s*([0-9]+)', r'\2-', string, flags=re.IGNORECASE)
    string = re.sub(r'^b[.,]*\s*(approximately [0-9]+)', r'\1-', string, flags=re.IGNORECASE)
    # d. -> -
    string = re.sub(r'd?d[.,]*\s*([0-9]+)', r'-\1', string)
    string = re.sub(r'^cd[.,]*\s*(approximately [0-9]+)', r'-\1', string, flags=re.IGNORECASE)
    string = re.sub(r'^D[.,]*\s*([0-9]+)', r'-\1', string, flags=re.IGNORECASE)
    # fl. -> active
    string = re.sub(r'^f[.,]*l?[.,]*\s*([0-9]+)', r'active \1', string, flags=re.IGNORECASE)
    # _ -> -
    # -- -> -
    string = string.replace('_', '-')
    string = string.replace('--', '-')
    # Remove r before numbers
    string = re.sub(r'r([0-9]{4})', r'\1', string)
    # bap. -. born approximately
    string = re.sub(r'bap\. ([0-9]{4})', r'approximately \1-', string)
    # Remove titles after dates
    string = re.sub(r'([0-9]+\??)\. .*$', r'\1', string)
    string = re.sub(r' [^ac][a-zA-Z]*$', '', string)
    string = unicodedata.normalize('NFC', string)
    return string


def clean_publication_places(string, ctys=None):
    publishers, states, places = set(), set(), set()
    if string == '': return publishers, states, places
    # Detect if string appears to contain publisher names
    # Space after 'for' to prevent detection of 'ford' e.g. Bradford
    publisher_flags = ['& co', 'book', 'for ', 'printed', 'private', 'published', 'shop', 'sold']
    if ',' not in string and any(s in string.lower() for s in publisher_flags):
        return clean_publisher_names(string)
    words_to_trim = ['one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten',
                     'first', 'second', 'third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth', 'ninth', 'tenth',
                     '1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th',
                     'erste', 'zweite', 'dritte', 'vierte', u'f\u00FCnfte', 'sechste', 'siebte', 'achte', 'neunte',
                     'zehnte',
                     'at', 'bde', 'by', u'd\u00EDl', 'for', 'heft', 'hefte', 'hft', 'hfte', 'honorary', 'imprinted',
                     'in', 'kn',
                     'page', 'pages', 'play', 'plays', 'pp', 'printed', 'privately', 'pt', 'reihe',
                     'sammlung', 'secretary', 'stuk', 'the', 'thl', 'vol',
                     u'\u03BC\u03B5\u03C1\u03B7', u'\u1F10\u03BD'
                     ]
    string = quick_clean(string)
    string = re.sub(r'\b([A-Z]) \[([a-z]+)\][.,\s]+', r'\1\2 ', string)
    string = re.sub(r'[\s\-.,]*[:;\[\]\\/(){}<>&*?|]+[\s\-.,]*', ';', string)
    string = re.sub(r'[\s\-.,]*(\bi[.\s]*e\b[.\s]*)[\s\-.,]*', ';', string, flags=re.IGNORECASE)
    if any(s in string.lower() for s in publisher_flags):
        string = re.sub(r'[\s\-.,]*,\s*(&|\bet\b|\band\b|\bund\b)[\s\-.,]*', ';', string, flags=re.IGNORECASE)
    else:
        string = re.sub(r'[\s\-.,]*(&|\bet\b|\band\b|\bund\b)[\s\-.,]*', ';', string, flags=re.IGNORECASE)
    string = re.sub(r'(\s*;\s*)+', ';', string)
    string = quick_clean(string, hyphens=False)
    for substring in string.split(';'):
        if ',' not in substring and any(s in substring.lower() for s in publisher_flags):
            subpublishers, substates, subplaces = clean_publisher_names(substring)
            for item in subpublishers:
                publishers.add(item)
            for item in substates:
                states.add(item)
            for item in subplaces:
                places.add(item)
            continue
        if substring and substring != '' and not is_number(substring):
            substring = clean_26X(quick_clean(substring, hyphens=False))
            substring = expand_place_abbreviations(substring, ctys)
            substring = quick_clean(substring, hyphens=False)

            if ' ' in substring:
                first = quick_clean(substring.split(None, 1)[0], hyphens=False)
                while ' ' in substring and (first.lower() in words_to_trim or re.sub(r'[0-9\-.,]', '', first) == '' or len(first) == 1 
                                            or first in countries.values() or first in ['S.A', 'U.S.A']):
                    if first in countries.values():
                        states.add(first)
                    elif first == 'S.A':
                        states.add('South Africa')
                    elif first in ['U.S.A']:
                        states.add('United States of America')
                    substring = quick_clean(substring.split(None, 1)[1], hyphens=False)
                    if ' ' not in substring: break
                    first = quick_clean(substring.split(None, 1)[0], hyphens=False)

            if ' ' in substring:
                last = quick_clean(substring.rsplit(None, 1)[1], hyphens=False)
                while ' ' in substring and (last.lower() in words_to_trim or re.sub(r'[0-9\-.,]', '', last) == '' or len(last) == 1
                                            or last in countries.values() or last in ['S.A', 'U.S.A']):
                    if last in countries.values():
                        states.add(last)
                    elif last == 'S.A':
                        states.add('South Africa')
                    elif last in ['U.S.A']:
                        states.add('United States of America')
                    substring = quick_clean(substring.rsplit(None, 1)[0], hyphens=False)
                    if ' ' not in substring: break
                    last = quick_clean(substring.rsplit(None, 1)[1], hyphens=False)

            # Take out full stops
            substring = substring.replace('.-', '-').replace('.', ' ')
            # Replace and with &
            substring = re.sub(r'[,.\-\s]*\b(and|et|und|&)\b[,.\-\s]*', ' & ', substring, flags=re.IGNORECASE)
            substring = quick_clean(substring, hyphens=False)

            substring = remove_quotes(substring)
            substring = quick_clean(substring, hyphens=False).strip('?')

            if substring.lower() in words_to_trim or len(substring) <= 3 or re.sub(r'[0-9\-.,]', '', substring) == '':
                substring = ''

            if substring != '':
                if substring in countries.values():
                    states.add(substring)
                elif substring == 'S A':
                    states.add('South Africa')
                elif substring in ['U S A']:
                    states.add('United States of America')
                elif substring in ['Great Britain', 'United Kingdom']:
                    states.add('United Kingdom')
                else:
                    places.add(substring)
                    substring = substring.title()
                    if substring in PLACES_US:
                        states.add('United States of America')
                    elif substring in PLACES_ENGLAND:
                        states.add('England')
                    elif substring in PLACES_IRELAND:
                        states.add('Ireland')
                    elif substring in PLACES_N_IRELAND:
                        states.add('Northern Ireland')
                    elif substring in PLACES_SCOTLAND:
                        states.add('Scotland')
                    elif substring in PLACES_WALES:
                        states.add('Wales')

    return publishers, states, places


def clean_publisher_names(string):
    publishers, states, places = set(), set(), set()
    if string == '': return publishers, states, places
    words_to_trim = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '&', '?', 'co', u'\u00E0', u'\u00E1',
                     'ac', u'academi\u00E6', 'act', 'aedibus', 'ah', 'also', 'all', 'an', 'and', 'appointment', 'apud',
                     'as', 'at', 'au', 'be', 'bei', 'beim', 'bey', 'books', 'bookds', 'bookseller', 'booksellers',
                     'bros', 'by', 'cased', 'catalogue', 'catalogues', 'cc', 'chez', 'chex', 'cie', 'city', 'co',
                     'collegio', 'company', 'cum', 'de', 'de\'', 'dem', 'der', 'designed', 'directs', 'didtributed',
                     'distribution', 'distributor', 'distributorp', 'distributors', 'distribuzione', 'each', 'ed',
                     u'\u00E9d', u'\u00E9diteur', 'edition', 'editions', 'editor', 'editore', 'editors', 'editrice',
                     'editura', 'edizioni', 'elsewhere', 'engraved', u'\u00E9rudition', 'esplanade', 'et', 'etiam',
                     'etonae', 'ex', 'exclusive', 'exclusively', 'excudebant', 'excudebat', 'execudit', 'excusum',
                     'filhos', 'filios', 'following', 'for', 'forlag', 'from', 'fur', u'f\u00FCr', 'gedruckt',
                     u'glasgu\u00E6', 'gmbh', 'graveur', 'had', 'her', 'herausgeber', 'him', 'his', 'impensa',
                     'impensis', 'impr', 'imprime', u'imprim\u00E9', 'imprimerie', 'imprint', 'imprinted', 'imprynt',
                     'in', 'inc', 'issued', 'istributor', 'izdanie', u'izdan\u012Be', u'izdatel\u02B9',
                     u'izdatel\u02B9stvo', u'izdava\u010Dko', 'izd-vo', 'kent', u'kiad\u00E1s', u'kiad\u00E1sa',
                     'kiadja', u'kiad\u00F3', 'king', u'k\u00F6nyvek', u'k\u00F6nyvkiad\u00F3' u'k\u00F6nyvkiad\u00F3',
                     'komm', 'kommission', 'law', 'librairies', 'librairie', 'librarie', u'libraire-\u00E9diteur',
                     'londini', u'limi\u1E6De\u1E0Da', 'majesty', 'may', 'musicsellers', 'na', 'nakladatelstvi', 'near',
                     'nella', 'newly', 'no', 'of', 'official', 'officina', 'oficyna', 'on', 'only', 'or', 'other',
                     'others', 'oxf', 'p', 'par', 'pl', 'por', 'pour', 'pr', u'preduze\u0107e', 'presso', 'print',
                     'printd', 'printed', 'printer', 'printers', 'printing', 'privately', 'privilegi', 'quem',
                     'reprinted', 're-printed', 'repirnted', 'rest', 'revised', 'sa', 'sculp', 'se', 'si', 'sohn',
                     'son', 'sons', 'son-in-law', 'stamperia', 'stereotype', 'stereotyped', 'stereotyper', 'sumtibus',
                     u'szerkeszt\u0151s\u00E9ge', 'the', 'them', 'tip', 'to', 'tpagr', u'tpagrut\u02BFiwn', 'tparan',
                     'translator', 'trova', 'trustees', 'typ', 'typograph', 'typographia', 'typographum', 'typis', 'u',
                     'uitgeverij', 'uk', 'vend', 'vendido', 'vendita', 'veneunt', 'verbindung', 'verl', 'verlag',
                     'verlegt', 'verlegung', 'von', 'vormals', 'with', 'within', 'without', 'withowt', 'written',
                     'wydawn', 'wydawnictwa', 'wydawnictwo', 'www', 'xi']
    # Can't add to the above: association, house, library, libraries, lit, mit, trust
    string = quick_clean(string)
    string = re.sub(r'\b([A-Z]) \[([a-z]+)\][.,\s]+', r'\1\2 ', string)
    string = re.sub(r'[\s\-.,]*[:;\[\]\\/(){}<>|]+[\s\-.,]*', ';', string)
    string = re.sub(r'[\s\-.,]*((\b[1-2][0-9]{3}\b|\.\.\.+|(&|\bet|\band|\bund)\s*c(o(mp(an)?y)?|(omp)?ie|orp)\b\.*|&c(ie)?\b\.*|\be[.\s]*t[.\s]*c\b[.\s]*|\bi[.\s]*e\b[.\s]*|\bl[.\s]*t[.\s]*d([.\s]*a)?\b[.\s]*|\bb[.\s]*[im]\b[.\s]*|\bn[.\s]*[pv]\b[.\s]*|\bs[.\s]*[inl]\b[.\s]*|\bp[.\s]*[ltv][.\s]*[cty]\b[.\s]*|\b(et\s*al|sic|viz)\b[.\s]*|(&|\bet|\band|\bund)\s*(bro(ther)?|son|the)s*\b[.\s]*|\b(also|distribution\s*(services*)?|exclusively|excudebant|incorporat(ing|ed)|issued|likewise|limited|lithographically|originally|serviced)\s*((by|for|in|into|par|pour|with)\b)?\s*(the\b)?|\bnot\s*avail(able)?\b|(\ban*)?\s*\b(book|division|imprint|part|publication)\s*(from|of)\b|(\bfor\s*)?\bsubscribers*\s*only\b|\btrading\s*as\b|(\bfor)?\s*(\bthe)?\s*\bprivate\s*(circulation|press)\b\s*(of\b)?|,([\s\-.,]*\b(an|and|at|by|chez|et|for|in|par|pour|the|typ|und|under|with)\b)+|\b(by|in|on|with)\s*(associatio?n|assignment|assistance|arrangement|authority|behalf|collaboration|conjunction|co-*operation|permission)\s*((from|of|with)\b)?\s*(the\b)?\s*((trustee|executor|guardian|proprietor)s*)?\s*(of\b)?\s*(the\b)?|(\b(for|from|of|with))?\s*(the)?\s*\b((trustee|executor|guardian|proprietor)s*)\b\s*(of\b)?\s*(the\b)?(late\b)?|\b((exclusive|joint|private)ly\s*)?((co|re)[\-\s]*)?((distribut|issu|print|produc|pub(lish|\.)?)ed\s*(&|\bet|\band|\bund)?\s*)+((exclusive|joint|private)ly\s*)?\s*((at\s*the\s*office\s*of|by|for|par|pour|with)\b)?\s*(the\b)?\s*(assistance\b)?\s*(of\b)?\s*(the\b)?|(\b(&|all|and|by|catholic|every|following|others?|principal|the|rest)[\s\-.,]*)*\b(administrator|author|distribute*or|.dit(eu|o)r|heir|perfumer|printer|proprietor|publisher|(book|law|music)[\s\-.,]*seller|stationer|successor)\'*s*(\s*(friend|syndicate)s?)?\b([\s\-.,]*(&|and|britain|city|country|county|great|in|kingdom|of|the|scotland|town|york)\b)*|\b(under|with)\s*the\s*((assistance|auspices?|co-*operation|direction|permission|sponsorship)\s*(&|\bet|\band|\bund)?\s*)+\s*(of\b)?\s*(the\b)?|((&|\bet|\band|\bund)\s*(are\s*to\s*be)|(&|\bet|\band|\bund)?\s*(are\s*to\s*be))\s*solde?\b|(&|\bet|\band|\bund)\s*subsidiar(y|ies)\b|\bar\s*ran\b|\bargraphwyd\s*(dros)?\b|\b(et\s*)?(a\s*lond.*?)?se\s*(trouve|vend)\b.*?(lond.*?)?che.*?propriet[aeious]*res?\b|\baux?\s*frais\s*(de\s*(l[eas]*\s*)?)?|\bimprim.e*.*?d.pens.*?ladite.*?academie\s*(par|pour)?\b|\bzu\s*finden\s*beym\b|\bdruck\s*der\b|\b(to|for)\s*((her|his|the\s*(king|queen)[.s\s]*most\s*excellent)\s*(majesty|royal\s*highness)([,\s]*pall-mall)?|the\s*society)([,\s]*(the\s*)?prince(\s*of\s*wales|sses))?\b|(&|\bet|\band|\bund)\s*(all|one|two|three|four|five|six|seven|eight|nine|ten|[0-9]+)?\s*others?\b|\ba favourite song in the enchanter\b|\ba scrapbook of pieces from m[.\s]*d\b[.\s]*|\bas\s*the\s*act\s*directs\b|\b(where|by\s*whom)\s*advertisements\s*are\s*taken\s*in\b|\bby\s*the\s*author.?s\s*appointment\b|\bmass\s*market\s*paperback\b|\bentered\s*at\s*stationer.?s hall\b|\b(also)?\s*(in|at)\s*(h(er|is)\s*majesty.?s|the|the\s*(king|queen)\'*s)\s*theatre\s*(in\b)?\s*(the\b)?\s*(hay-*market)?|\bin\s*(the)?\s*(u[.\s]*s[.\s]*a\b[.\s]*|north\s*(&|and)\s*south\s*america|united\s*states(\s*of\s*america)?|western\s*hemisphere)|\bat\s*(his|the)\s*(library|shop|(wholesale)?\s*warehouses?),?(\s*on\s*the\s*esplanade)?)[\s\-.,]*)+', ';', string, flags=re.IGNORECASE)
    string = re.sub(r'[\s\-.,]*\b(committee|office)s*\s*of\b[\s\-.,]*\'', ';\'', string, flags=re.IGNORECASE)
    string = re.sub(r'\'[\s\-.,]*\b(publi(cation|shing))s*[\s\-.,]*', '\';', string, flags=re.IGNORECASE)
    # Words to split after
    string = re.sub(r'[\s\-.,]*\b(agency|associatio?n|library|newspapers|organisation|organization|society|trust|university press)[\s\-.,]+(?!for|of)', r' \1;', string, flags=re.IGNORECASE)
    string = re.sub(r'[\s\-.,]*\b(press|publications|publishing)[\s\-.,]*(co)?[\s\-.,]+(?!and|house)', r' \1;', string, flags=re.IGNORECASE)
    # Words to split before
    string = re.sub(r'[\s\-.,]*\b(bloomsbury|british library|british school of|dover|j(ohn)? murray|methuen|penguin)', r';\1 ', string, flags=re.IGNORECASE)
    # Words to split between
    string = re.sub(r'[\s\-.,]*\b(books|london|westminster)[\s\-.,]*(?:&|et|and|und)?[\s\-.,]*(for the|london|westminster)\b', r' \1;\2 ', string, flags=re.IGNORECASE)
    string = re.sub(r'(\s*;\s*)+', ';', string)
    string = quick_clean(string)

    for substring in string.split(';'):
        if substring and substring != '' and not is_number(substring):
            substring = clean_26X(quick_clean(substring, hyphens=False))
            substring = publisher.Publishers().sub(substring).strip()
            if substring not in ['Books of Africa', 'Independent Publishers Group']:
                substring = re.sub(r'\s*\bu(ni)?(versity)?[.\s]*pr*(ess)?\b[.\s]*', ' University Press ', substring, flags=re.IGNORECASE)
                substring = quick_clean(substring, hyphens=False)

                if ' ' in substring:
                    first = quick_clean(substring.split(None, 1)[0], hyphens=False)
                    while ' ' in substring and (first.lower() in words_to_trim or is_lower_case_letter(first)):
                        substring = quick_clean(substring.split(None, 1)[1], hyphens=False)
                        if ' ' not in substring: break
                        first = quick_clean(substring.split(None, 1)[0], hyphens=False)

                if ' ' in substring:
                    last = quick_clean(substring.rsplit(None, 1)[1], hyphens=False)
                    while ' ' in substring and (last.lower() in words_to_trim or is_lower_case_letter(last)):
                        substring = quick_clean(substring.rsplit(None, 1)[0], hyphens=False)
                        if ' ' not in substring: break
                        last = quick_clean(substring.rsplit(None, 1)[1], hyphens=False)

                # Take out full stops
                substring = substring.replace('.-', '-').replace('(?!www).(?<!(co|uk))', ' ')
                # Replace and with &
                substring = re.sub(r'[,.\-\s]*\b(and|et|und|&)\b[,.\-\s]*', ' & ', substring, flags=re.IGNORECASE)
                substring = quick_clean(substring, hyphens=False)

                substring = remove_quotes(substring)
                substring = quick_clean(substring, hyphens=False).strip('?')

                if substring.lower() in words_to_trim or len(substring) <= 3 \
                        or substring in ['book', 'group', 'publications', 'publishing'] \
                        or re.sub(r'[0-9\-.,]', '', substring) == '':
                    substring = ''

                if substring != '':
                    if substring in countries.values():
                        states.add(substring)
                    elif substring == 'S A':
                        states.add('South Africa')
                    elif substring in ['U S A']:
                        states.add('United States of America')
                    elif substring in ['Great Britain', 'United Kingdom']:
                        states.add('United Kingdom')
                    elif substring.title() in PLACES:
                        substring = substring.title()
                        places.add(substring)
                        if substring in PLACES_US:
                            states.add('United States of America')
                        elif substring in PLACES_ENGLAND:
                            states.add('England')
                        elif substring in PLACES_IRELAND:
                            states.add('Ireland')
                        elif substring in PLACES_N_IRELAND:
                            states.add('Northern Ireland')
                        elif substring in PLACES_SCOTLAND:
                            states.add('Scotland')
                        elif substring in PLACES_WALES:
                            states.add('Wales')
                    else:
                        for item in substring.split(';'):
                            item = quick_clean(item)
                            if item != '':
                                publishers.add(item)
    return publishers, states, places
