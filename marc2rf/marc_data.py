# -*- coding: utf8 -*-

"""Classes for MARC records, fields and subfields used in the Researcher Format transformation.
Code uses elements of https://github.com/edsu/pymarc but with significant modifications."""

# Import required functions
from marc2rf.cleaning_functions import clean

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'


# ====================
#     Constants
# ====================

q880 = True

LEADER_LEN = 24
DIRECTORY_ENTRY_LEN = 12
SUBFIELD_INDICATOR = chr(0x1F)
END_OF_FIELD = chr(0x1E)
END_OF_RECORD = chr(0x1D)
ALEPH_CONTROL_FIELDS = ['DB ', 'FMT', 'SYS']

# ====================
#     Exceptions
# ====================


class RecordLengthError(Exception):
    def __str__(self): return 'Invalid record length in first 5 bytes of record'


class LeaderError(Exception):
    def __str__(self): return 'Error reading record leader'


class DirectoryError(Exception):
    def __str__(self): return 'Record directory is invalid'


class FieldsError(Exception):
    def __str__(self): return 'Error locating fields in record'


class BaseAddressLengthError(Exception):
    def __str__(self): return 'Base address exceeds size of record'


class BaseAddressError(Exception):
    def __str__(self): return 'Error locating base address of record'


# ====================
#       Classes
# ====================


class MARCReader(object):

    def __init__(self, marc_target):
        # print(str(marc_target))
        super(MARCReader, self).__init__()
        if hasattr(marc_target, 'read') and callable(marc_target.read):
            self.file_handle = marc_target

    def __iter__(self):
        return self

    def close(self):
        if self.file_handle:
            self.file_handle.close()
            self.file_handle = None

    def __next__(self):
        first5 = self.file_handle.read(5)
        if not first5: raise StopIteration
        if len(first5) < 5: raise RecordLengthError
        return Record(first5 + self.file_handle.read(int(first5) - 5))


class Record(object):
    def __init__(self, data='', leader=' ' * LEADER_LEN):
        self.leader = '{}22{}4500'.format(leader[0:10], leader[12:20])
        self.fields = list()
        self.pos = 0
        if len(data) > 0: self.decode_marc(data)

    def __str__(self):
        text_list = ['=LDR  {}'.format(self.leader)]
        text_list.extend([str(field) for field in self.fields])
        return '\n'.join(text_list) + '\n'

    def __getitem__(self, tag):
        fields = self.get_fields(tag)
        if len(fields) > 0: return fields[0]
        return None

    def __contains__(self, tag):
        fields = self.get_fields(tag)
        return len(fields) > 0

    def __iter__(self):
        self.__pos = 0
        return self

    def __next__(self):
        if self.__pos >= len(self.fields): raise StopIteration
        self.__pos += 1
        return self.fields[self.__pos - 1]

    def add_field(self, *fields):
        self.fields.extend(fields)

    def get_fields(self, *args):
        """
        Returns a list of all the fields in a record with a given tag.
            subjects = record.get_fields('600', '610', '650')
        If no tag is specified a list of all the fields will be returned.
        """
        if len(args) == 0: return self.fields
        return [f for f in self.fields if (f.tag in args or (q880 and f.tag == '880' and '6' in f and str(f['6'])[:3] in args))]

    def decode_marc(self, marc):
        # Extract record leader
        self.leader = marc[0:LEADER_LEN].decode('ascii')
        if len(self.leader) != LEADER_LEN: raise LeaderError

        # Extract the byte offset where the record data starts
        base_address = int(marc[12:17])
        if base_address <= 0: raise BaseAddressError
        if base_address >= len(marc): raise BaseAddressLengthError

        # Extract directory
        # base_address-1 is used since the directory ends with an END_OF_FIELD byte
        directory = marc[LEADER_LEN:base_address - 1].decode('ascii')

        # Determine the number of fields in record
        if len(directory) % DIRECTORY_ENTRY_LEN != 0:
            raise DirectoryError
        field_total = len(directory) / DIRECTORY_ENTRY_LEN

        # Add fields to record using directory offsets
        field_count = 0
        while field_count < field_total:
            entry_start = field_count * DIRECTORY_ENTRY_LEN
            entry_end = entry_start + DIRECTORY_ENTRY_LEN
            entry = directory[entry_start:entry_end]
            entry_tag = entry[0:3]
            entry_length = int(entry[3:7])
            entry_offset = int(entry[7:12])
            entry_data = marc[base_address + entry_offset:base_address + entry_offset + entry_length - 1]

            # Check if tag is a control field
            if str(entry_tag) < '010' and entry_tag.isdigit():
                field = Field(tag=entry_tag, data=entry_data.decode('utf-8'))
            elif str(entry_tag) in ALEPH_CONTROL_FIELDS:
                field = Field(tag=entry_tag, data=entry_data.decode('utf-8'))

            else:
                subfields = list()
                subs = entry_data.split(SUBFIELD_INDICATOR.encode('ascii'))
                # Missing indicators are recorded as blank spaces.
                # Extra indicators are ignored.

                subs[0] = subs[0].decode('ascii') + '  '
                first_indicator, second_indicator = subs[0][0], subs[0][1]

                for subfield in subs[1:]:
                    if len(subfield) == 0: continue
                    code, data = subfield[0:1].decode('ascii'), subfield[1:].decode('utf-8', 'strict')
                    subfields.append(code)
                    subfields.append(data)
                field = Field(
                    tag=entry_tag,
                    indicators=[first_indicator, second_indicator],
                    subfields=subfields,
                )
            self.add_field(field)
            field_count += 1

        if field_count == 0: raise FieldsError


class Field(object):

    def __init__(self, tag, indicators=None, subfields=None, data=''):
        if indicators is None: indicators = []
        if subfields is None: subfields = []
        indicators = [str(x) for x in indicators]

        # Normalize tag to three digits
        self.tag = '%03s' % tag

        # Check if tag is a control field
        if self.tag < '010' and self.tag.isdigit():
            self.data = str(data)
        elif self.tag in ALEPH_CONTROL_FIELDS:
            self.data = str(data)
        else:
            self.indicator1, self.indicator2 = self.indicators = indicators
            self.subfields = subfields

    def __iter__(self):
        self.__pos = 0
        return self

    def __str__(self):
        if self.is_control_field() or self.tag in ALEPH_CONTROL_FIELDS:
            text = '={}  {}'.format(self.tag, self.data.replace(' ', '\\'))
        else:
            text = '={}  '.format(self.tag)
            for indicator in self.indicators:
                if indicator in (' ', '\\'): text += '\\'
                else: text += indicator
            for subfield in self: text += '${}{}'.format(subfield[0], subfield[1])
        return text

    def __getitem__(self, subfield):
        """
        Retrieve the FIRST subfield with a given subfield code:
            field['a']
        """
        subfields = self.get_subfields(subfield)
        if len(subfields) > 0: return subfields[0]
        return None

    def __contains__(self, subfield):
        subfields = self.get_subfields(subfield)
        return len(subfields) > 0

    def __next__(self):
        if not hasattr(self, 'subfields'):
            raise StopIteration
        while self.__pos < len(self.subfields):
            subfield = (self.subfields[self.__pos], self.subfields[self.__pos + 1])
            self.__pos += 2
            return subfield
        raise StopIteration

    def get_subfields(self, *codes, cleaning=True):
        """Accepts one or more subfield codes and returns a list of subfield values
        Subfields are cleaned unless clean=False (may be useful for subfields containing URLs)
        """
        values = []
        for subfield in self:
            if len(codes) == 0 or subfield[0] in codes:
                if cleaning: values.append(clean(str(subfield[1])))
                else: values.append(str(subfield[1]))
        return values

    def is_control_field(self):
        if self.tag < '010' and self.tag.isdigit(): return True
        if self.tag in ALEPH_CONTROL_FIELDS: return True
        return False
