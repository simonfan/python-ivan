import csv
import os
import re
import xlrd
import string

# gets the cell_position from an string in the format 'F99'.
# currently only supports single-letter-column positioning
def cell_position(cell_string):
  match = re.search('([A-Z]|[a-z])([0-9]*)', cell_string);

  return {
    # get the position in alphabet of the character
    'column': string.uppercase.index(match.group(1)),
    'row': int(match.group(2)) - 1
  };

# reads the fields to be extracted
def read_extraction_fields():
  with open('extract.csv', 'rb') as csvfile:
    fields = [];
    for row in csv.reader(csvfile, delimiter=',', quotechar='"'):
      fields.append({
        'cell': cell_position(row[0]),
        'title': row[1]
      });

    return fields;

print read_extraction_fields();