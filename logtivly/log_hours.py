
from __future__ import print_function
import json

import datetime
import sys
import os.path
import subprocess
from credentials import credentials

if __debug__:
    from workflow import Workflow, ICON_WEB, web
    from workflow.notify import notify

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
        return sheetTitle, 0

def get_projects_and_hours():
    service, spreadsheetId = credentials.get_service_and_spreadsheetId()
    sheetTitle, colNum = get_sheet_title_and_column(service, spreadsheetId)
    projectCells = 'B16:H19'
    initialProjectCellIndex = 16
    rangeName = "%s!%s" % (sheetTitle, projectCells)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName, majorDimension='COLUMNS').execute()
    values = result.get('values',[])
    return (
            colNum
            , sheetTitle
            , values and values[0]
            , values and values[colNum+1]
        )

def get_project_cell(projectStr, sheetTitle, colNum):
    service, spreadsheetId = credentials.get_service_and_spreadsheetId()
    cols = ['c','d','e','f','g','h']
    columnLetter = cols[colNum]
    # it's not likely that there wil be more than 4 projects at a time
    # but if there is, do logic that fetches all rows before the "total hours" row starts
    projectCells = 'B16:%s19' % columnLetter
    initialProjectCellIndex = 16
    rangeName = "%s!%s" % (sheetTitle, projectCells)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName, majorDimension='COLUMNS').execute()
    values = result.get('values',[])
    projectNames = list(map((lambda x: x.lower()), values and values[0]))
    rowIndex = [i for i, s in enumerate(projectNames) if projectStr.lower() in s] or [0]
    initialCellValue = values[colNum+1][rowIndex[0]] if len(values) > 1 and rowIndex else 0
    return '%s%s' % (columnLetter, initialProjectCellIndex + rowIndex[0]), float(initialCellValue), projectNames[rowIndex[0]] if projectNames else ''

def py_main():
    colNum, sheetTitle, projects, hours = get_projects_and_hours()
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

def main(wf):

    if not os.path.isfile('~/.credentials/'+credentials.CREDS_FILENAME):
        subprocess.Popen(['nohup', 'python', './credentials/credentials.py'])

    if len(wf.args):
        sys.stderr.write("there ARE arguments!\n")
        sys.stderr.write(json.dumps(wf.args))

        projectStr = " ".join(wf.args[0:len(wf.args)-1])
        addedhours = wf.args[len(wf.args)-1]
        sys.stderr.write("\n\nproject str: %s\n\n" % (projectStr))

        colNum, sheetTitle, projects, hours = wf.cached_data('projects_hours', get_projects_and_hours, max_age=60)
        cell, initialCellValue, retrievedProjectStr = get_project_cell(projectStr, sheetTitle, colNum)
        sys.stderr.write("\n\n"+json.dumps([sheetTitle, cell, initialCellValue, retrievedProjectStr, addedhours])+"\n\n")
        rangeName = '%s!%s:%s' % (sheetTitle, cell, cell)
        sys.stderr.write("updating range" + rangeName)
        values = [[initialCellValue + float(addedhours)]]
        body = {
            'values': values
        }

        service, spreadsheetId =  credentials.get_service_and_spreadsheetId()
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheetId, range=rangeName, body=body, valueInputOption="USER_ENTERED").execute()

        notify('success!',
               'you have added %s hours to "%s", totalling %s hours' % (float(addedhours), retrievedProjectStr, float(addedhours) + initialCellValue))
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



if __name__ == '__main__':
    if __debug__:
        wf = Workflow(libraries=['./'])
        sys.exit(wf.run(main))
    else:
        py_main()
