import xlrd
import sys
import os
import csv

if __name__ == "__main__":
  if len(sys.argv) > 1:
    xlsxfiles = sys.argv[1:]
    print("XLSX conversion tool")
    for xlsx in xlsxfiles:
      if os.path.isfile(xlsx):
        print(f"Converting {xlsx} to UTF-8 CSV {xlsx}.[sheet_name].csv")
        book = xlrd.open_workbook(xlsx)
        for sheetname in book.sheet_names():
          with open(f"{xlsx}.{sheetname}.csv", "w", encoding="utf-8") as csvsheet:
            print(f"{xlsx}-{sheetname}  ---> {xlsx}.{sheetname}.csv")
            c = csv.writer(csvsheet)
            sh = book.sheet_by_name(sheetname)
            for rn in range(sh.nrows):
              c.writerow(sh.row_values(rn))

