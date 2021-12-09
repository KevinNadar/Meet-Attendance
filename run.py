#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jul  1 11:34:09 2021

@author: kevin
"""

import os
import gspread
from datetime import date, datetime
from Calendar import FetchEvent
from Sheets import AttendanceRegister

now = datetime.now().ctime()

dt = date.today()\

f = open(f'/home/kevin/.config/spyder-py3/Projects/Meet-Attendance/Logs/{dt}', 'w')

f.write('\n' + now)

# try:

client = gspread.oauth()

instance = FetchEvent(dt)

mark = instance.events_to_attend(
    instance.latest_events, 'This Google Calendar event is going to be \
accounted in the Meet Attendance.'
)

spreadsheet = client.open('Meet Attendance 07/02/2021')

attendanceRegister = AttendanceRegister(client, spreadsheet, f)

ev_to_ws = attendanceRegister.event_to_worksheet(mark, dt)

attendanceRegister.register_attendance(ev_to_ws)

attendanceRegister.delete_sheet(spreadsheet, spreadsheet.worksheets()[-1])

print('Done!')
f.write('\nDone!\n')

# except Exception as e:
#     print(e)

f.close()
