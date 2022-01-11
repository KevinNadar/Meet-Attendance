#!/home/kevin/anaconda3/envs/attendance/bin python3.9
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 21:51:30 2021

@author: kevin
"""
import pandas as pd
from datetime import datetime as dt


class AttendanceRegister:

    def __init__(self, client, spreadsheet, textfile):
        self.spreadsheet = spreadsheet
        self.client = client
        self.textfile = textfile

    def event_to_worksheet(self, events, day):
        cal_event = list()
        sheets = filter(lambda x: dt.strptime(x.title[:10], '%m/%d/%Y').date() == day, self.spreadsheet.worksheets())
        for sheet in sheets:
            times = list(map(dt.strptime, sheet.col_values(4)[1:], ['%H:%M:%S'] * len(sheet.col_values(4)[1:])))
            before = dt.strptime(sheet.cell(times.index(max(times)) + 2, 2).value, '%m/%d/%Y %H:%M:%S')
            after = dt.strptime(sheet.cell(times.index(max(times)) + 2, 3).value, '%m/%d/%Y %H:%M:%S')
            difference = after - before
            for event in events:
                if dt.strptime(event['start']['dateTime'][:-6], '%Y-%m-%dT%H:%M:%S') <= before + difference / 2 <= dt.strptime(event['end']['dateTime'][:-6], '%Y-%m-%dT%H:%M:%S'):
                    cal_event.append((event, sheet))
                    print(f'{event["summary"].split()[0]} lecture data in {sheet.title} worksheet')
                    self.textfile.write(f'\n{event["summary"].split()[0]} lecture data in {sheet.title} worksheet')
        return cal_event

    def register_attendance(self, tup):
        for event_to_mark, worksheet in tup:
            now = dt.strptime(worksheet.cell(2, 2).value, '%m/%d/%Y %H:%M:%S')
            present_students = worksheet.col_values(1)[1:]
            students_workbook = self.client.open(event_to_mark['summary'].split(' ')[0] + ' Attendance Register')
# =============================================================================
#             TIP: Manual Override;
#             students_workbook = self.client.open('<name_of_workbook>')
# =============================================================================
            student_worksheet = students_workbook.worksheet(now.strftime('%B'))
            print(f'{event_to_mark["summary"].split()[0]} lecture data to be registered in {students_workbook.title} workbook')
            self.textfile.write(f'\n{event_to_mark["summary"].split()[0]} lecture data to be registered in {students_workbook.title} workbook')
            register = pd.DataFrame(
                data=student_worksheet.get_all_values()[1:],
                index=student_worksheet.col_values(2)[1:],
                columns=student_worksheet.row_values(1)
            )
            register[now.strftime('%A %d')] = 'Present'
            student_list = [name.upper() for name in present_students]
            i = 0
            while i < student_list.__len__():
                attendee = student_list[i]
                if attendee.startswith('KAVITHA') or attendee not in register.index:
                    student_list.remove(attendee)
                    print(f'Not marking {attendee} in {students_workbook.title}')
                    self.textfile.write(f'\nNot marking {attendee} in {students_workbook.title}')
                    continue
                i += 1
            if not student_list:
                self.textfile.write('\nNo student in classroom')
                raise ValueError('No student in classroom')
            register.loc[list(set(register.index) - set(student_list)), now.strftime('%A %d')] = 'Absent'
            register['Days Present'] = register.replace(['', 'Absent', 'Present'], [0, 0, 1]).sum(axis='columns')
            student_worksheet.update('A2', register.values.tolist())

    def delete_sheet(self, wb, ws):
        if wb.worksheets().__len__() > 7:
            wb.del_worksheet(ws)
            print(f'Deleted {ws.title} worksheet from {wb.title} workbook')
            self.textfile.write(f'\nDeleted {ws.title} worksheet from {wb.title} workbook')
