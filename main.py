import os
from scripts.utils.csv_to_xl import csv_to_xl


if __name__ == '__main__':
    csv_title = 'apr-weekly-timesheet'
    for file in os.listdir():
        if csv_title in file:
            csv_to_xl(file, file[0:len(file) - 4] + '.xlsx')


# We should find and alternative solution to this. I just did it for now in order
# to write the logic for the funciton.
from scripts.TimeSheet import TimeSheet

if __name__ == '__main__':
    timeSheet = TimeSheet()
    timeSheet.generate_work_sheet()
    timeSheet.populate_data()

# year = datetime.datetime.today().strftime('%Y')
## This will be used for an edge case conditional that will create a new file after each Year.
# new_file_name = f'Time Reporting Log ({current_month} {year}).xlsx'
