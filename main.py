import os

from scripts.TimeSheet import TimeSheet
from scripts.utils.csv_to_xl import csv_to_xl


if __name__ == '__main__':
    timeSheet = TimeSheet()
    timeSheet.generate_work_sheet()
    csv_title = 'apr-weekly-timesheet'
    for file in os.listdir():
        if csv_title in file:
            csv_to_xl(file, file[0:len(file) - 4] + '.xlsx')


# year = datetime.datetime.today().strftime('%Y')
## This will be used for an edge case conditional that will create a new file after each month.
# new_file_name = f'Time Reporting Log ({current_month} {year}).xlsx'
