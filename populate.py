from openpyxl import load_workbook
weekly_timesheet = 'apr-weekly-timesheet-dec-5th.xlsx'


def get_punch_times(column, am_pm) -> list:
    ws = load_workbook(weekly_timesheet)['Sheet']
    """Returns a list of dicts with early clock in or late clock out names, dates and times. Example uses:
    get_punch_times('E', 'pm'), get_punch_times('C', 'am')"""
    punch_times = []
    for cell in ws[column]:
        value = cell.value
        if not bool(value):
            continue
        if am_pm in value:
            left_of_colon = int(value[0:value.index(':')])
            if am_pm == 'pm':
                condition = left_of_colon > 5 and left_of_colon != 12
            else:
                condition = left_of_colon < 8
            if condition:
                punch_times.append({
                    'name': ws[f'A{cell.row}'].value,
                    'date': ws[f'B{cell.row}'].value,
                    'clock-out': value
                })
    return punch_times

    