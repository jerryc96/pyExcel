from openpyxl import load_workbook
from sheet import *
import argparse

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('filepath')
    args = parser.parse_args()    
    filepath = args.filepath
    # load workbook, store the worksheet in a sheet object
    wb = load_workbook(filepath, read_only=True)
    sheet = Sheet(wb.active)
    # prompt user to enter a keyfield
    while True:
        # prompt the user to enter an ID field
        value = input("Select the ID you're trying to find: ")
        try:
            print(sheet.lookup_ID(value))
        except KeyError as k:
            print(k)
        break