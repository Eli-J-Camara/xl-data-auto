from openpyxl import load_workbook

weekly_timesheet = 'apr-weekly-timesheet-dec-5th.xlsx'
ws = load_workbook(weekly_timesheet)['Sheet']


def get_early_clock_ins():
    """returns a list of all early clock-in cells"""
    early_clock_ins = []
    for cell in ws['C']:
        value = cell.value
        if not bool(value):
            continue
        # Second part of bool makes sure number to left of colon is less than 8
        if 'am' in value and int(value[0:value.index(':')]) < 8:
            early_clock_ins.append(f'{cell.column_letter}{cell.row}')
    return early_clock_ins


print(get_early_clock_ins())
