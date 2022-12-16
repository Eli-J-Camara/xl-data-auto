from TimeSheet import TimeSheet

if __name__ == '__main__':
    timeSheet = TimeSheet()
    timeSheet.generate_work_sheet()


# year = datetime.datetime.today().strftime('%Y')
## This will be used for an edge case conditional that will create a new file after each month.
# new_file_name = f'Time Reporting Log ({current_month} {year}).xlsx'
