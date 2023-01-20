import os
from scripts.utils.csv_to_xl import csv_to_xl
from scripts.TimeSheet import TimeSheet

if __name__ == '__main__':
    # Changed variable below to match the file names that will be dragged and dropped into project.
    csv_title = 'Ace Party Rental'
    for file in os.listdir():
        if csv_title in file and 'csv' in file:
            if csv_to_xl(file, file[0:len(file) - 4] + '.xlsx'):
                timeSheet = TimeSheet()
                timeSheet.generate_work_sheet()
                timeSheet.populate_data()
                break

