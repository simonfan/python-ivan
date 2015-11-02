import csv
import os
import xlrd

# Function that extracts the required values from a given xlsx file
def extract_values(file_name):
  book = xlrd.open_workbook(file_name)
  sheet = book.sheet_by_index(0)

  return {
    'f149': sheet.cell_value(148, 5),
    'g149': sheet.cell_value(148, 6)
  };

# define variable to hold the data
csv_data = [];

xlsx_files_dir = './files'

for xlsx_file_name in os.listdir(xlsx_files_dir):

  values = extract_values(xlsx_files_dir + '/' + xlsx_file_name);

  csv_data.append([
    values['f149'],
    values['g149']
  ])

print csv_data

# write data
with open('results.csv', 'w') as fp:
  a = csv.writer(fp, delimiter=',')
  a.writerows(csv_data)