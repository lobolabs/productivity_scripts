
from __future__ import print_function
import httplib2
import os
import json

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import datetime
import sys
from optparse import OptionParser

if __debug__:
    from workflow import Workflow, ICON_WEB, web
    from workflow.notify import notify



# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


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
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
    return credentials

def get_spreadsheet_cell(service, spreadsheetId, projectStr="misc"):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    dateCells = "C15:H15"
    for sheet in spreadsheet['sheets']:
        sheetTitle = sheet['properties']['title']
        rangeName = "%s!%s" % (sheetTitle, dateCells)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheetId, range=rangeName).execute()
        values = result.get('values',[])
        colNum = 0
        # TODO: replace current year with actual cell year, see http://stackoverflow.com/q/42216491/766570
        for row in values:
            for column in row:
                dateStr = "%s %s" % (column, datetime.date.today().year)
                cellDate = datetime.datetime.strptime(dateStr, '%b %d %Y')
                if cellDate.date() == datetime.date.today():
                    return get_project_cell(service, spreadsheetId, projectStr, sheetTitle, colNum)
                colNum +=1

def get_sheet_title_and_column(service, spreadsheetId):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    dateCells = "C15:H15"
    for sheet in spreadsheet['sheets']:
        sheetTitle = sheet['properties']['title']
        rangeName = "%s!%s" % (sheetTitle, dateCells)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheetId, range=rangeName).execute()
        values = result.get('values',[])
        colNum = 0
        # TODO: replace current year with actual cell year, see http://stackoverflow.com/q/42216491/766570
        for row in values:
            for column in row:
                dateStr = "%s %s" % (column, datetime.date.today().year)
                cellDate = datetime.datetime.strptime(dateStr, '%b %d %Y')
                if cellDate.date() == datetime.date.today():
                    return sheetTitle, colNum

                colNum +=1

def get_projects_and_hours():
    service, spreadsheetId = get_service_and_spreadsheetId()
    sheetTitle, colNum = get_sheet_title_and_column(service, spreadsheetId)
    projectCells = 'B16:H19'
    initialProjectCellIndex = 16
    rangeName = "%s!%s" % (sheetTitle, projectCells)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName, majorDimension='COLUMNS').execute()
    return colNum, sheetTitle, result.get('values',[])[0], result.get('values',[])[colNum+1]

def get_project_cell(service, spreadsheetId, projectStr, sheetTitle, colNum):
    cols = ['c','d','e','f','g','h']
    columnLetter = cols[colNum]
    # it's not likely that there wil be more than 4 projects at a time
    # but if there is, do logic that fetches all rows before the "total hours" row starts
    projectCells = 'B16:%s19' % columnLetter
    initialProjectCellIndex = 16
    rangeName = "%s!%s" % (sheetTitle, projectCells)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName, majorDimension='COLUMNS').execute()
    projectNames = list(map((lambda x: x.lower()), result.get('values',[])[0]))
    rowIndex = [i for i, s in enumerate(projectNames) if projectStr.lower() in s][0]
    initialCellValue = result['values'][colNum+1][rowIndex]
    return '%s%s' % (columnLetter, initialProjectCellIndex + rowIndex), float(initialCellValue), projectNames[rowIndex]

def py_main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1tRziPgOwgPBCYIQZuwa7Me1vk5_tfsAWb3mcpPICxEI'

    projects, hours = get_projects_and_hours(service, spreadsheetId)
    index = 0
    items = []
    for project in projects:
        #sys.stderr.write("project: " + project + "\n")
        items.append({'title':project,
                    'subtitle':"hours logged: "+hours[index],
                    'icon':'ICON_WEB',
                    'autocomplete':project})
        index+=1

    print(items)

def get_service_and_spreadsheetId():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1tRziPgOwgPBCYIQZuwa7Me1vk5_tfsAWb3mcpPICxEI'

    return service, spreadsheetId

def main(wf):
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    if len(wf.args):
        sys.stderr.write("there ARE arguments!\n")
        sys.stderr.write(json.dumps(wf.args))
        projectStr = args[0]
        hours_to_add = float(args[1])
        colNum, sheetTitle, projects, hours = wf.cached_data('projects_hours', get_projects_and_hours, max_age=60)
        sheetTitle, cell, initialCellValue = get_project_cell(service, spreadsheetId, projectStr, sheetTitle, colNum)
        sys.stderr.write(json.dumps([sheetTitle, cell, initialCellValue]))
        rangeName = '%s!%s:%s' % (sheetTitle, cell, cell)
        sys.stderr.write("updating range" + rangeName)
        values = [[initialCellValue + hours]]
        body = {
            'values': values
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=rangeName, body=body, valueInputOption="USER_ENTERED").execute()

        notify('success!',
               'you have added %s hours to "%s", totalling %s hours' % (hours, projectStr, hours + initialCellValue))
    else:
        sys.stderr.write("there are no arguments!\n")
        colNum, sheetTitle, projects, hours = get_projects_and_hours()
        index = 0
        for project in projects:
            #sys.stderr.write("project: " + project + "\n")
            wf.add_item(title=project,
                        subtitle="hours logged: "+hours[index],
                        icon=ICON_WEB,
                        autocomplete=project, )
            index+=1
        wf.send_feedback()

    # sheetTitle, cell, initialCellValue, projectName = get_spreadsheet_cell(service, spreadsheetId, projectStr)
    # rangeName = '%s!%s:%s' % (sheetTitle, cell, cell)
    # values = [[initialCellValue + hours]]
    # body = {
    #         'values': values
    #        }
    # result = service.spreadsheets().values().update(
    #         spreadsheetId=spreadsheetId, range=rangeName, body=body, valueInputOption="USER_ENTERED").execute()
    #
    # notify('success!','you have added %s hours to "%s", totalling %s hours' % (hours, projectName, hours + initialCellValue))


if __name__ == '__main__':
    parser = OptionParser()
    (options, args) = parser.parse_args()
    if __debug__:
        wf = Workflow()
        sys.exit(wf.run(main))
    else:
        py_main()
