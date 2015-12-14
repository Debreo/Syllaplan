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
import argparse
import datetime
#try:
#    import argparse
#    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
#except ImportError:
#    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Syllaplan'
now = datetime.datetime.now()
#credentials = get_credentials()
#http = credentials.authorize(httplib2.Http())
#service = discovery.build('calendar', 'v3', http=http)

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

#def syllaparsed(parsed):
    #credentials = get_credentials()
    #http = credentials.authorize(httplib2.Http())
    #service = discovery.build('calendar', 'v3', http=http)   
    #dates = []
    #assignments = []
    #dates_hw = []
    #parsed = parsed.strip().replace("\n"," ")
    #if re.findall('\n?\w+\s\d+\/\d+\n?',parsed):
       # print (parsed.strip().replace("\n", " "))
    #   dates.append(parsed)
    #else:
    #    assignments.append(parsed)
    #print (parsed)
    #created_event = service.events().quickAdd(calendarId='primary',
    #text=parsed.strip().replace("\n"," ")).execute()
    #print (created_event)

def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    syllabus = argparse.ArgumentParser(prog="Syllaplan", description='Don\'t Cram Syllaplan!')
    syllabus.add_argument("-f",
                        dest="file",
                        type = str,
                        action="store",
                        required = True,
                        nargs = "+",
                        help=" The File to extract")
    args = syllabus.parse_args()  
    dates = []
    assignments = []
    for sylla in args.file:
        document = Document(sylla)
        wordDoc = document.tables 
        for table in wordDoc:
            for row in table.rows:
                for cell in row.cells:
                    if len(row.cells) == 3:
                        if "\n" in (cell.text).encode('utf-8') or "/" in (cell.text).encode('utf-8'):
                            if re.findall('\n?\w+\s\d+\/\d+\n?',cell.text.encode('utf-8')):
                                dates.append((cell.text.strip().encode('utf-8')))
                            else:
                                assignments.append(((cell.text).encode('utf-8')).replace("\n"," "))
                    elif len(row.cells) == 4:
                        if "/" in (cell.text).encode('utf-8'):
                            m = re.search('\d+\/\d+',cell.text.encode('utf-8'))
                            dates.append(m.group(0))
                            s = re.split('\d+\/\d+',cell.text.encode('utf-8'))
                            for items in s:
                                if "" != items:
                                    (assignments.append(items.strip().replace("\n"," ").replace("\xe2\x80\x93","")))
                    
    duedates = OrderedDict(zip(dates, assignments))
        #print (duedates)


    for x,y in  duedates.items():
            #duetext = x.encode('utf-8')+" "+ y.encode('utf-8')  
        duetext = x + "/"+str(now.year)+ " "+y
            #print (duetext)
        created_event = service.events().quickAdd(calendarId='primary',
        text=duetext).execute()
    print (created_event)
    print ("Events have been created please check your calendar")

if __name__ == '__main__':
    main()
