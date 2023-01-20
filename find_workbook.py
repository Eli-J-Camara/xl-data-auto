import os
from openpyxl import load_workbook

def find():
    csv_title = 'Ace Party Rental'
    for file in os.listdir():
        if csv_title in file and 'csv' in file:
            file = file[0:len(file) - 4] + '.xlsx'
            return file