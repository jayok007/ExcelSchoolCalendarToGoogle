import pandas as pd
import re
import sys
from datetime import datetime

# Class containing all the information about a course
class Course:
    def __init__(self, date_start, date_end, course_name, teacher=None, place=None):
        self.date_start = date_start
        self.date_end = date_end
        self.course_name = course_name
        self.teacher = teacher
        self.place = place
    
    # Create the dictionnary to sent to the google API
    def to_google_event(self,color_id=None):
        event = {
        'summary': self.course_name,
        'start': {
            'dateTime': self.date_start.isoformat(),
            'timeZone': 'Europe/Paris',
        },
        'end': {
            'dateTime': self.date_end.isoformat(),
            'timeZone': 'Europe/Paris',
        }}
        
        if color_id:
            event['colorId']= color_id,

        if self.place:
            event['location'] = self.place
        if self.teacher:
            event['description']= self.teacher
        return event

    # Create an event in ics
    def to_icalendar_event(self):
        event = Event()
        event.add('summary', self.course_name)
        if self.teacher:
            event.add('description', f'{self.teacher}')
        event.add('dtstart', self.date_start)
        event.add('dtend', self.date_end)
        if self.place:
            event.add('location', self.place)
        return event

    def __str__(self):
        return f'{self.date_start} -> {self.date_end}\n\t Course: {self.course_name}\n\t\n\t Teacher: {self.teacher}\n\t \n\t Place: {self.place}\n\t'

def is_place(places, potential_place):
    for place in places:
        if place in potential_place:
            return True
    return False

# Return a timestamp and take as an input the date and  the course timerange
# 9h30-12h ->  return [datetime(date.year, date.month, date.day, 9,30),datetime(date.year, date.month, date.day, 12)]
def str_time_to_timestamp(pd_date, str_time):
    # Split on - -> then split on h 
    # 9h30-12h -> [[9,30],[12]]
    split_time = [list(filter(lambda x : len(x.strip()) > 0, x.split('h'))) for x in str_time.split('-')]
    res_dates = []
    for date in split_time:
        if len(date) == 1:
            res_dates.append(datetime(pd_date.year, pd_date.month, pd_date.day, int(date[0])))
        else:
            res_dates.append(datetime(pd_date.year, pd_date.month, pd_date.day, int(date[0]), int(date[1])))
    return res_dates


def get_events()
    # Read the edt as a panda data frame
    df = pd.read_excel('edt.xls')

    # Known substring in place
    places = ['EDF R&D', 'SALLE', 'salle']        

    # Get the shape of the imported matrix
    shape = df.shape
    first_column = df.iloc[:,0]
    days = ['LUNDI', 'MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI']
    # Get the start vertical start index for each day
    days_indexes = [(day, first_column[first_column == day].index[0]) for day in days]


    formated_courses = []
    # For each day
    for day_index ,day_info in enumerate(days_indexes):
        
        # For each column
        for i in range(3,shape[1]):
            day_name, start_day_index = day_info
            
            if type(df.iloc[start_day_index,i]) == pd.Timestamp:
                date = df.iloc[start_day_index, i]

                # If friday end of vertical range is end of matrix otherwise it's next day index 
                end_of_day_row_index = shape[0] if day_index == (len(days) -1) else days_indexes[day_index + 1][1]                
                
                # Browsing verticaly until the next day or the end of matrix (Friday)
                contents = list(set(filter(lambda x : type(x) != float, df.iloc[start_day_index+1:end_of_day_row_index, i])))
                contents = [list(filter(lambda x: len(x)> 0, content.split('\n'))) for content in contents]
                
                # Since excel is not well formated checking for smaller content processing then for for bigger one
                if len(contents) > 0:
                    for course in contents:
                        if len(course) == 2 and 'TOEIC' in course[1]:
                                date_start_end = str_time_to_timestamp(date, '08h00-19h00')
                                course_name = 'TOEIC (Horaire non défini)'
                                professor = 'Prof Anglais'
                                formated_courses.append(Course(date_start_end[0], date_start_end[1], course_name, professor))
                        elif len(course) == 3:
                            date_start_end = str_time_to_timestamp(date, course[0].lower())
                            course_name = course[1]
                            professor = course[2]
                            formated_courses.append(Course(date_start_end[0], date_start_end[1], course_name, professor))
                        elif len(course) >= 4:
                            # Double time range
                            if 'H' in course[0] and 'H' in course[1]:
                                date_start_end1 = str_time_to_timestamp(date, course[0].lower())
                                date_start_end2 = str_time_to_timestamp(date, course[1].lower())
                                course_name = course[2]
                                professor = course[3]
                                formated_courses.append(Course(date_start_end1[0], date_start_end1[1], course_name, professor))
                                formated_courses.append(Course(date_start_end2[0], date_start_end2[1], course_name, professor))
                            else:
                                date_start_end = str_time_to_timestamp(date, course[0].lower())
                                course_name = course[1]
                                professor = course[2]
                                start_index_other = 3
                                place = None
                                if is_place(places,course[3]):
                                    place = course[3]
                                formated_courses.append(Course(date_start_end[0], date_start_end[1], course_name, professor, place))
                        else:
                            print('Course not taken into account !', course)
    return formated_courses


