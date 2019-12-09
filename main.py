from google_base import get_calendar_service
import edt_utils
import sys
import argparse
from icalendar import Calendar, Event

# Check the color for a specifi course
def get_color(colors_events, course_name):
    for colors_k, colors_v in colors_events.items():
        if colors_k in course_name:
            return colors_v
    return None


parser = argparse.ArgumentParser()
parser.add_argument("--file", type=str, help="Excel file you want to transform", default='edt.xls')
group = parser.add_mutually_exclusive_group(required=True)
group.add_argument("--store_to_ics", type=str, help="ics filename to transform your excel edt")
group.add_argument("--sent_to_google_calendar", type=str, help="Google Calendar Id to import your calendar")
group.add_argument("--clear_google_calendar", type=str, help="Google Calendar Id to clear")
args = parser.parse_args()

if args.store_to_ics:

    cal = Calendar()
    for course in edt_utils.get_events():
        cal.add_component(course.to_icalendar_event())
    with open(args.store_to_ics, 'wb') as f:
        f.write(cal.to_ical())

elif args.sent_to_google_calendar:

    calendar_id = args.sent_to_google_calendar

    # Get the google service/client
    service = get_calendar_service()
    
    colors_events = {'MULTICORE SYSTEM': 1,  'DATA MINING': 2, 'STRUCTURAL MODELS': 3, 'TIMES SERIES': 4, 'MACHINE LEARNING': 5, 'PLS': 6, 'ONLINE SURVEY': 7, 'SPLUNK': 8, 'MICRO': 9, 'RESEAU DE NEURONES': 10, 'ANGLAIS': 11, 'TOEIC': 11 ,'CRM': 7, 'INTRODUCTION TO BIG DATA': 8}    
    
    # For each course add it to my class calendar 
    for course in edt_utils.get_events():
        event = course.to_google_event(get_color(colors_events, course.course_name))
        print(event)
        service.events().insert(calendarId=calendar_id, body=event).execute()

elif args.clear_google_calendar:

    calendar_id = args.clear_google_calendar

    # Get the google service/client
    service = get_calendar_service()

    # Clear all code
    page_token = None
    while True:
        events = service.events().list(calendarId=calendar_id, pageToken=page_token).execute()
        for event in events['items']:
            service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()  
        page_token = events.get('nextPageToken')
        if not page_token:
            break   







