from openpyxl import load_workbook
import datetime


def ws_header():
    # For now, we will just get the past 5 days, once we have a UI we should give Chaz the option
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


class TimeSheet:

    def __init__(self):
        self.months = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August',
                       9: 'September', 10: 'October', 11: 'November', 12: 'December'}
        self.title = self.create_ws_title()
        self.header = ws_header()

    def create_ws_title(self) -> str:
        """Generates dynamic work sheet title."""
        month_num = datetime.datetime.today().strftime('%m')
        current_month = self.months[int(month_num)]
        day = datetime.datetime.today().strftime('%d')
        ws_name = f'Weekly Time Reporting {current_month[0:3]} {day}.'
        return ws_name

    def generate_work_sheet(self):
        excel_template = 'TimeReportLog(Oct2022).xlsx'
        wb = load_workbook(filename=excel_template)
        template_ws = wb['Weekly Time Reporting Template']
        new_ws = wb.copy_worksheet(template_ws)
        new_ws.title = self.title
        new_ws['I1'] = self.header[0]

        for cells in self.header[1]:
            for i in range(5):
                index = self.header[1].index(cells)
                new_ws[cells[i]] = self.header[2][index]

        # Python doesn't like the file name to more than 31 characters!
        wb.save('TimeReportLog(Oct2022).xlsx')
