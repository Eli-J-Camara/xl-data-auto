import datetime


def create_ws_title() -> str:
    months = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August',
              9: 'September', 10: 'October', 11: 'November', 12: 'December'}
    month_num = datetime.datetime.today().strftime('%m')
    current_month = months[int(month_num)]
    day = datetime.datetime.today().strftime('%d')
    ws_name = f'Weekly Time Reporting {current_month[0:3]} {day}.'
    return ws_name


class TimeSheet:

    def __init__(self):
        self.title = create_ws_title()
