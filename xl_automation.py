# REMEMBER YOU ARE USING PYTHON3!!!
from openpyxl import load_workbook
import datetime

# Reconfigure all this to be in a Class structure. (maybe?)


# Generating dynamic work sheet title.
def work_sheet_title():
    months = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August',
              9: 'September', 10: 'October', 11: 'November', 12: 'December'}
    month_num = datetime.datetime.today().strftime('%m')
    current_month = months[int(month_num)]
    day = datetime.datetime.today().strftime('%d')
    ws_name = f'Weekly Time Reporting {current_month[0:3]} {day}.'
    return ws_name


def work_sheet_header():
    # For now we will just get the past 5 days, once we have a UI we should give Chaz the option
    # to input any date range he wants.
    to_date = datetime.datetime.today().strftime('%m/%d')
    from_date = datetime.datetime.today() - datetime.timedelta(days=4)
    from_date_form = from_date.strftime('%m/%d')
    log_heading = f'Time Reporting Summary ({from_date_form} - {to_date}):'

    # Header date numberings.
    from_range = int(from_date_form[3:5])
    to_range = int(to_date[3:5])
    header_dates = [date for date in range(from_range, to_range + 1)]
    header_date_column = ['A', 'E', 'I', 'M', 'Q']
    header_date_row = ['8', '19', '27', '35', '43']
    header_cells = [[col + row for row in header_date_row] for col in header_date_column]
    return log_heading, header_cells, header_dates


# Creating and writing to the work sheet.
def generate_work_sheet():
    excel_template = 'TimeReportLog(Oct2022).xlsx'
    wb = load_workbook(filename=excel_template)
    template_ws = wb['Weekly Time Reporting Template']
    new_ws = wb.copy_worksheet(template_ws)
    new_ws.title = work_sheet_title()
    new_ws['I1'] = work_sheet_header()[0]

    cells_and_dates = work_sheet_header()
    for cells in cells_and_dates[1]:
        for i in range(5):
            index = cells_and_dates[1].index(cells)
            new_ws[cells[i]] = cells_and_dates[2][index]

    # Python doesn't like the file name to more than 31 characters!
    wb.save('TimeReportLog(Oct2022).xlsx')


generate_work_sheet()















# year = datetime.datetime.today().strftime('%Y')
## This will be used for an edge case conditional that will create a new file after each month.
# new_file_name = f'Time Reporting Log ({current_month} {year}).xlsx'