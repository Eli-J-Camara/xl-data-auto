"""source: https://www.blog.pythonlibrary.org/2021/09/25/converting-csv-to-excel-with-python/#:~:text=Converting%20a
%20CSV%20file%20to%20Excel,-You%20will%20soon&text=Your%20code%20uses%20Python's%20csv,to%20the%20input%20CSV%20file"""
import csv
import os
import openpyxl


def csv_to_xl(csv_file, excel_file):
    # Avoids UnicodeDecodeError
    csv_data = []
    with open(csv_file) as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in csv_data:
        sheet.append(row)
    workbook.save(excel_file)


def is_correct_csv(excel_file) -> bool:
    for file in os.listdir():
        if file == excel_file:
            return False
    return True
