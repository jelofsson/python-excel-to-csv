# export.py 
# Export the sheets of Excel file to CSV files.
import sys
import csv
import xlrd

excel_filename  = sys.argv[1] if len(sys.argv) > 1 else 'excel.xlsx'
sheet_name	= sys.argv[2] if len(sys.argv) > 2 else False

workbook = xlrd.open_workbook(excel_filename)
for sheet in workbook.sheets():
    if sheet_name == False or sheet.name == sheet_name: 
        with open('{}.csv'.format(sheet.name), 'wb') as f:
            writer = csv.writer(f, delimiter=";")
            for row in range(sheet.nrows):
                out = []
                for cell in sheet.row_values(row):
                    if isinstance(cell, float):
                        out.append(cell)
                    else:
                        out.append(unicode(cell).encode('utf8'))
                writer.writerow(out)
