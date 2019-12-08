from google_base import get_calendar_service
import edt_utils
import sys

# Check the color for a specifi course
def get_color(colors_events, course_name):
    for colors_k, colors_v in colors_events.items():
        if colors_k in course_name:
            return colors_v
    return None

# Get the google service/client
service = get_calendar_service()
colors_events = {'MULTICORE SYSTEM': 1,  'DATA MINING': 2, 'STRUCTURAL MODELS': 3, 'TIMES SERIES': 4, 'MACHINE LEARNING': 5, 'PLS': 6, 'ONLINE SURVEY': 7, 'SPLUNK': 8, 'MICRO': 9, 'RESEAU DE NEURONES': 10, 'ANGLAIS': 11, 'TOEIC': 11 ,'CRM': 7, 'INTRODUCTION TO BIG DATA': 8}

events = edt_utils.get_events()

# For each course add it to my class calendar 
for course in events:
    event = course.to_google_event(get_color(colors_events, course.course_name))
    print(event)
    service.events().insert(calendarId='5uvaj1au30r7o1va10fc5v13p0@group.calendar.google.com', body=event).execute()
    