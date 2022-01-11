#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jun 22 17:14:28 2021

@author: kevin
"""

import os.path
from pprint import pprint
from datetime import datetime as dt
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.events.readonly']


def main(day, month, year):
    global events_result, creds  # OPTIMIZE: Added global

    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('/home/kevin/.config/spyder-py3/Projects/Meet-Attendance/token.json'):
        creds = Credentials.from_authorized_user_file(
            '/home/kevin/.config/spyder-py3/Projects/Meet-Attendance/token.json', SCOPES
        )
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '/home/kevin/.config/spyder-py3/Projects/Meet-Attendance/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('/home/kevin/.config/spyder-py3/Projects/Meet-Attendance/token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    timeMin = dt(year, month, day, 0, 0, 0, 0).strftime('%Y-%m-%dT%H:%M:%SZ')
    timeMax = dt(year, month, day, 23, 59, 59, 999999).strftime('%Y-%m-%dT%H:%M:%SZ')
    # print('Getting the upcoming 10 events')
    # now = dt.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    events_result = service.events().list(
        calendarId='primary', timeMin=timeMin, timeMax=timeMax, maxResults=10,
        singleEvents=True, orderBy='startTime').execute()
# =============================================================================
#     events = events_result.get('items', [])
#
#     if not events:
#         print('No upcoming events found.')
#     for event in events:
#         start = event['start'].get('dateTime', event['start'].get('date'))
#         print(start, event['summary'])
# =============================================================================

    return events_result['items']

class FetchEvent:
    """This class is used to get today's Google Calendar Events"""
    def __init__(self, date):
        self.day = date.day
        self.latest_events = main(date.day, date.month, date.year)

# =============================================================================
#     def todays_events(self, today):
#         """This method returns todays events"""
#         sieve = lambda start: dt.fromisoformat(
#             start['start']['dateTime']).date() == today
#         events = filter(sieve, self.latest_events['items'])
#         return list(events)
# =============================================================================

    def events_to_attend(self, filtered_events, description):
        regist = []
        for event in filtered_events:
            try:
                if event['description'] != description:
                    continue
                print(event['summary'])
            except KeyError:
                continue
            regist.append(event)
        return regist


if __name__ == '__main__':
    output = main(11, 1, 2022)
    pprint(output)
