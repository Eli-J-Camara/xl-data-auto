import os
from scripts.utils.csv_to_xl import csv_to_xl, is_correct_csv
from scripts.TimeSheet import TimeSheet

if __name__ == '__main__':
    # Changed variable below to match the file names that will be dragged and dropped into project.
    csv_title = 'Ace Party Rental'
    for csv_file in os.listdir():
        if csv_title in csv_file and 'csv' in csv_file:
            excel_file = csv_file[0:len(csv_file) - 4] + '.xlsx'
            if is_correct_csv(excel_file):
                csv_to_xl(csv_file, excel_file)
                timeSheet = TimeSheet()
                timeSheet.generate_work_sheet()
                timeSheet.populate_data()
                os.remove('./Ace Party Rental_2022-06-13_2022-06-17_timesheets.xlsx')
                break

