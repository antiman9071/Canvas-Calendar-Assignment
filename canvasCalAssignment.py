from dataclasses import dataclass
import icalendar
import pytz, zoneinfo, dateutil.tz  
from datetime import datetime
from pathlib import Path
from sys import argv
import os
import xlwt
import xlrd
from xlutils.copy import copy

@dataclass
class dailyTask:
    date:int
    dayOfTheWeek:int
    key:str

#TODO create a way to check for days of the week then write to excel
def grab_assignment(summary):
    check = summary.find('Office Hours')
    check2 = summary.find('office hours')
    check3 = summary.find('Office hours')
    check4 = summary.find('office Hours')
    if check == -1 and check2 == -1 and check3 == -1 and check4 == -1:
        start = summary.find('y:') + 1
        end = summary.find('[')

        if start != -1 and end != -1 and start < end:
            return summary[start:end]
        else:
            return None
    else:
        return None

def sort_assignments(calendarDict):
    workingList = list()
    for key in calendarDict.keys():
        value=calendarDict.get(key)
        workingList.append(dailyTask(int(value.strftime("%W"))-1, int(value.strftime("%u"))+1, key))
    return workingList
    
if len(argv) > 1:
    ics_path = Path(argv[1])
else:
    if os.name != "posix":
        ics_path = Path(os.getcwd() + "\Downloads\canvas_output.ics")
    else:
        ics_path = Path(os.getcwd() + "/Downloads/canvas_output.ics")

with ics_path.open(encoding='utf8') as fileIN:
    calendar = icalendar.Calendar.from_ical(fileIN.read())
calendarDict = dict()
for event in calendar.walk('VEVENT'):
    if grab_assignment(event.get("SUMMARY")) != None:
        calendarDict[grab_assignment(event.get("SUMMARY"))] = event["DTSTART"].dt

workingList = sort_assignments(calendarDict)
lastDate = 0
lastDOW = 0
rb = xlrd.open_workbook('classwork/Semester at a Glance- SP 25 LC.xls')
wb = copy(rb)
wksheet = wb.get_sheet(0)
data = None
for dailyTask in workingList:
    if(dailyTask.date == lastDate and dailyTask.dayOfTheWeek == lastDOW):
        data += "\n" + dailyTask.key
    else:
        wksheet.write(lastDate, lastDOW, data)
        data = dailyTask.key
    lastDOW = dailyTask.dayOfTheWeek
    lastDate = dailyTask.date
wb.save('classwork/Semester at a Glance- SP 25 LC.xls')
