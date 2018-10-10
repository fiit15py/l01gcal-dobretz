# -*- coding: utf-8 -*-
from __future__ import print_function
import xlrd
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools


def next_weekday(d, weekday):
    days_ahead = weekday - d.weekday()
    if days_ahead <= 0:  # Target day already happened this week
        days_ahead += 7
    return d + datetime.timedelta(days_ahead)


xls = xlrd.open_workbook('imi2018.xls')
sheet_names = xls.sheet_names()

print('Sheet Names', sheet_names)
sheet = xls.sheet_by_name(u'4 курс_ИТ')

row_to_day = {
    3: 'Понедельник',
    9: 'Вторник',
    15: 'Среда',
    21: 'Четверг',
    27: 'Пятница',
    33: 'Суббота',
}

row_to_weekDay = {
    3: 1,
    4: 1,
    5: 1,
    6: 1,
    7: 1,
    8: 1,
    9: 2,
    10: 2,
    11: 2,
    12: 2,
    13: 2,
    14: 2,
    15: 3,
    16: 3,
    17: 3,
    18: 3,
    19: 3,
    20: 3,
    21: 4,
    22: 4,
    23: 4,
    24: 4,
    25: 4,
    26: 4,
    27: 5,
    28: 5,
    29: 5,
    30: 5,
    31: 5,
    32: 5,
    33: 6,
    34: 6,
    35: 6,
    36: 6,
    37: 6,
    38: 6,
}

start_time = {
    1: 'T08:00:00+09:00',
    2: 'T09:50:00+09:00',
    3: 'T11:40:00+09:00',
    4: 'T14:00:00+09:00',
    5: 'T15:45:00+09:00',
    6: 'T17:30:00+09:00',
}

end_time = {
    1: 'T09:35:00+09:00',
    2: 'T11:25:00+09:00',
    3: 'T13:15:00+09:00',
    4: 'T15:35:00+09:00',
    5: 'T17:20:00+09:00',
    6: 'T19:05:00+09:00',
}

d = datetime.datetime.now()
day = {
    1: next_weekday(d, 0).strftime('%Y-%m-%d'),
    2: next_weekday(d, 1).strftime('%Y-%m-%d'),
    3: next_weekday(d, 2).strftime('%Y-%m-%d'),
    4: next_weekday(d, 3).strftime('%Y-%m-%d'),
    5: next_weekday(d, 4).strftime('%Y-%m-%d'),
    6: next_weekday(d, 5).strftime('%Y-%m-%d'),
    7: next_weekday(d, 6).strftime('%Y-%m-%d'),
}

cell_obj = sheet.cell(2, 8)
print(cell_obj.value)

store = file.Storage('token.json')
creds = store.get()
service = build('calendar', 'v3', http=creds.authorize(Http()))

for row_idx in range(3, sheet.nrows - 1):  # Iterate through rows
    cell_obj = sheet.cell(row_idx, 8)  # Get cell object by row, col
    cell_type = sheet.cell(row_idx, 9)
    cell_kab = sheet.cell(row_idx, 10)
    cell_kab_value = cell_kab.value
    if isinstance(cell_kab_value, float):
        cell_kab_value = int(cell_kab_value)
    if row_idx in row_to_day:
        print(row_to_day[row_idx])
    print(str((row_idx - 3) % 6 + 1) + ') ' + cell_obj.value + '(' + cell_type.value + ') ' + str(cell_kab_value))
    weekIndex = row_to_weekDay[row_idx]
    lessonIndex = (row_idx - 3) % 6 + 1
    if len(cell_obj.value) > 1:
        event = {
            'summary': cell_obj.value,
            'location': 'KFEN' + str(cell_kab_value),
            'description': cell_type.value,
            'start': {
                'dateTime': day[weekIndex] + start_time[lessonIndex],
                'timeZone': 'Asia/Yakutsk',
            },
            'end': {
                'dateTime': day[weekIndex] + end_time[lessonIndex],
                'timeZone': 'Asia/Yakutsk',
            },
            'recurrence': [
                'RRULE:FREQ=WEEKLY;COUNT=4'
            ],
            'attendees': [
                {'email': 'dgena4@gmail.com'},
            ],
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 24 * 60},
                    {'method': 'popup', 'minutes': 10},
                ],
            },
        }
        event = service.events().insert(calendarId='primary', body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))
