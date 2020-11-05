#! python
# fromExcelToCsv.py - get the excel data to the csv data
# XI 2020 Arnold Cytrowski

import os, csv, openpyxl

for filename in os.listdir('.'):
    if filename.endswith('.xlsx'):
        wb = openpyxl.load_workbook(filename)
        

        for sheet_name in wb.get_sheet_names():
            sheet = wb.get_sheet_by_name(sheet_name)

            csv_name = filename[:-5]
            csv_file = open(f'{csv_name}_{sheet_name}.csv', 'w', newline='')
            csv_writer = csv.writer(csv_file)

            for row_num in range(1, sheet.max_row + 1):
                row_data = []
                for col_num in range(1, sheet.max_col + 1):
                    row_data.append(sheet.cell(row = row_num, column = col_num).value)

                for row in row_data:
                    csv_writer.writerow(row_num)

            csv_file.close()

            