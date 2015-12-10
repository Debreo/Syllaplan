from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime
import sys
from docx import Document
import re                                                                                                                                                                                                
import json
import yaml
from collections import OrderedDict
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Syllaplan'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    file = "Syllabus.docx"
    document = Document(file)
    wordDoc = document.tables
    dates = []
    assignments = []
    count = 0
    for table in wordDoc:
        for row in table.rows:
            for cell in row.cells:
                if "\n" in (cell.text).encode('utf-8') or "/" in (cell.text).encode('utf-8'):
                    if re.findall('\n?\w+\s\d+\/\d+\n?',cell.text.encode('utf-8')):

                        dates.append((cell.text.strip()))
                    else:
                        assignments.append(cell.text.replace("\n"," "))
    duedates = OrderedDict(zip(dates, assignments))
    for x,y in  duedates.items():
        duetext = x.encode('utf-8')+" "+ y.encode('utf-8')  
        created_event = service.events().quickAdd(calendarId='primary',
        text=duetext).execute()
    print (created_event)


if __name__ == '__main__':
    main()
