import sys
import xlrd
import json

debug = None
try:
  if sys.argv[1] == '-d':
    debug = True
  else:
    debug = False
except IndexError:
  debug = False

result_array = []
book = xlrd.open_workbook("file.xls")


sheet_names = book.sheet_names()
sheet_total_pages = book.nsheets;


for i in range(0, sheet_total_pages):
  current_sheet = book.sheet_by_index(i)
  
  if  current_sheet.name == "PRINCIPAL":
    if debug:
      print (" **---** Ignoring sheet page [%d] **---**" % (i))
  else:
    if debug:
      print ("Parsing sheet page [%d]" % (i))
    
    num_cols = current_sheet.ncols # Number of columns
    for row_idx in range(6, current_sheet.nrows):    # Iterate through rows (Starting 6st col, not 0, to not read store name and cif)
      if debug:
        print ('Row: %s' % row_idx) # Prints row Number
      for col_idx in range(1, num_cols):  # Iterate through columns
        cell_obj = current_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        if debug:
          print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))  
        if cell_obj.value: # If cell value (string) is not empty
          email = cell_obj.value
          email = email.lower()
          result_array.append(email)

result_array = list(set(result_array)) # Removing duplicate entries
result_array.sort()

result_json = json.dumps(result_array)

file_ = open('result.json', 'w')
file_.write(result_json)
file_.close()