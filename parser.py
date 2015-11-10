import csv
import os
import re
import xlrd
import string

# unicode hack encoding... HACK
# python csv module does not handle unicode well, didn't have time to read docs.
# anytime in the future: https://docs.python.org/2.7/library/csv.html#csv-examples
def hack_unicode(s):
  if isinstance(s, str):
    return s;
  elif isinstance(s, unicode):
    return s.encode('utf-8');
  else:
    return s;

# colors
class bcolors:
  HEADER = '\033[95m'
  OKBLUE = '\033[94m'
  OKGREEN = '\033[92m'
  WARNING = '\033[93m'
  FAIL = '\033[91m'
  ENDC = '\033[0m'
  BOLD = '\033[1m'
  UNDERLINE = '\033[4m'

# gets the cell_position from an string in the format 'F99'.
# currently only supports single-uppercase-letter-column positioning
def cell_position(cell_string):
  match = re.search('([A-Z])([0-9]*)', cell_string);

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
        'title': row[1] + '(' + row[0] + ')'
      });

    return fields;

# extracts the value from a field
# handles errors for cell reading
def extract_value_from_cell(sheet, field):

  # print 'reading ' + field['title'];

  try:
    return sheet.cell_value(
      field['cell']['row'],
      field['cell']['column']
    );
  except:
    # error_message = 'error reading ' + field['title'];
    # print error_message;
    return error_message;

# Function that extracts the required values from a given xlsx file
# takes the fields to extract as a second argument
def extract_values_from_file(file_name, fields):
  book = xlrd.open_workbook(file_name, on_demand = True);

  # get correct sheet
  sheet = book.sheet_by_index(2);

  # define array to hold results
  results = [];

  # loop through fields reading values
  for field in fields:
    cell_value = hack_unicode(extract_value_from_cell(sheet, field));
    results.append(cell_value);

  return results;

# extraction

# takes the dirname
def extract_values_from_dir(xlsx_files_dirname):

  # get fields to be extracted
  # in the format: [{ 'cell': { 'column': 9, 'row': 10 }, 'title': 'field title' }]
  fields = read_extraction_fields();

  # define variable to hold the data
  results_csv = [];

  # define result header row
  header = ['file'];
  for field in fields:
    header.append(hack_unicode(field['title']));

  results_csv.append(header);

  xlsx_files_dir = './' + xlsx_files_dirname;

  for xlsx_file_name in os.listdir(xlsx_files_dir):

    try:

      print 'reading values from ' + xlsx_file_name;

      values = extract_values_from_file(
        xlsx_files_dir + '/' + xlsx_file_name,
        fields
      );

      # insert file name into values
      values.insert(0, xlsx_file_name);

      results_csv.append(values);

      print bcolors.OKGREEN + 'successfully read file ' + xlsx_file_name + bcolors.ENDC;

    except:
      results_csv.append([
        xlsx_file_name,
        'error reading file',
      ]);
      print bcolors.WARNING + 'error at file ' + xlsx_file_name + bcolors.ENDC;

  # write data
  with open(xlsx_files_dirname + '-results.csv', 'w') as fp:
    a = csv.writer(fp, delimiter=',')
    a.writerows(results_csv);

# run it once
extract_values_from_dir('medicamentos_1');
extract_values_from_dir('medicamentos_2');
extract_values_from_dir('medicamentos_3');
extract_values_from_dir('medicamentos_4');