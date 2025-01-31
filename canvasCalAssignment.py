import icalendar
import pytz, zoneinfo, dateutil.tz  
from datetime import datetime
from pathlib import Path
import xlwt
from sys import argv
import os

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

    for key in calendarDict.keys():
        if(isinstance(calendarDict.get(key),datetime)):
            return "done"
        else:
            return calendarDict.get(key)
    
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
    calendarDict[event.get("SUMMARY")] = event["DTSTART"].dt.strftime("%A|%m/%d/%y")
for key in calendarDict.keys():
    print("key=", grab_assignment(key), "value=",calendarDict.get(key))
print(sort_assignments(calendarDict))
