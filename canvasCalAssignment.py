import icalendar
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

if len(argv) > 1:
    ics_path = Path(argv[1])
else:
    ics_path = Path(os.getcwd() + "\Downloads\canvas_output.ics")

with ics_path.open(encoding='utf8') as fileIN:
    calendar = icalendar.Calendar.from_ical(fileIN.read())
    for event in calendar.walk('VEVENT'):
        print(grab_assignment(event.get("SUMMARY")))