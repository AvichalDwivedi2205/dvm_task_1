import csv
import openpyxl
import pandas as pd
from openpyxl import load_workbook

# Function to create course dictionary
def create_course_dict(name, code, test_date, credits):
    return {'Course Name': name, 'Course Code': code, 'Test Date': test_date, 'Credits': credits}

# Function to read excel data and convert it into list of dictionaries
def read_excel_data(file_name):
    # Read the excel file
    df = pd.read_excel(file_name)
    
    # Convert dataframe to list of dictionaries
    data = df.to_dict(orient='records')
    
    return data

# Function to export data to csv
def export_to_csv(data, filename):
    # Convert list of dictionaries to dataframe
    df = pd.DataFrame(data)
    
    # Write dataframe to csv file
    df.to_csv(filename, index=False)

# Read excel data and convert it into list of dictionaries
excel_data = read_excel_data("input.xlsx")

# Filter the list to keep only the rows where clashes occur
clash_data = [row for row in excel_data if clashes(row)]

# Export the filtered list of dictionaries to a csv file
export_to_csv(clash_data, "timetable.csv")

class Course:
    def __init__(self, name, code, test_date, credits):
        self.name = name
        self.code = code
        self.credits = credits
        self.test_date = test_date
        self.sections = []

    def add_section(self, section):
        self.sections.append(section)

    def get_sections(self):
        return self.sections


class Section:
    def __init__(self, course, id, room, type):
        self.course = course
        self.section_id = id
        self.room = room
        self.day = []
        self.hours = []
        if id[0] == "L":
            self.type = "Lecture"
        elif id[0] == "T":
            self.type = "Tutorial"
        elif id[0] == "P":
            self.type = "Practise"


class Timetable:
    def __init__(self):
        self.courses = []
        self.enrolled_sections = []

    def enroll_course(self, course):
        self.courses.append(course)

    def add_enrolled_section(self, section):
        self.enrolled_sections.append(section)

    def check_clashes(self, section):
        for course in self.courses:
            if section in course.get_sections():
                for enrolled_section in self.enrolled_sections:
                    if enrolled_section.course != section.course:
                        if enrolled_section.day[0] == section.day[0]:
                            if enrolled_section.hours[0] == section.hours[0]:
                                return True
                            else:
                                return False
                        else:
                            return False
                    else:
                        if enrolled_section.type != section.type:
                            if enrolled_section.day[0] == section.day[0]:
                                if enrolled_section.hours[0] == section.hours[0]:
                                    return True
                                else:
                                    return False
                            else:
                                return False
        return False


# Export to CSV
def export_to_csv(data, filename):
    with open(filename, 'w', newline='') as f:
        writer = csv.writer(f)
        for row in data:
            writer.writerow(row)

# Example usage
timetable = Timetable()
course1 = Course("Math", "MATH101", "2022-06-01", 3)
section1 = Section(course1, "L01", "R101", "Lecture")
section1.day.append("M")
section1.hours.append("10:00")
section2 = Section(course1, "T01", "R102", "Tutorial")
section2.day.append("W")
section2.hours.append("14:00")
course1.add_section(section1)
course1.add_section(section2)
timetable.enroll_course(course1)
timetable.add_enrolled_section(section1)
timetable.add_enrolled_section(section2)
clash = timetable.check_clashes(section2)

if clash:
    print("There is a clash.")
else:
    print("There is no clash.")

# Export timetable data to CSV
export_data = [
    ["Course Name", "Course Code", "Test Date", "Credits"],
    [course1.name, course1.code, course1.test_date, course1.credits],
]
export_to_csv(export_data, "timetable.csv")