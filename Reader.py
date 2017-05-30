import os
import xlrd
from Student import Student
from validate_email import validate_email

def ReadStudentData(directory):
    """
    :param directory: OS directory to search for course files
    :return: A dictionary of Students found in spreadsheets
    """
    students = []
    firstname_col = 0
    lastname_col = 1
    email_col = 2
    score_col = 5

    for filename in os.listdir(directory):
        if filename.endswith(".xlsx") and filename.startswith('Course Detail'):
            course = None
            book = xlrd.open_workbook(directory + filename)
            first_sheet = book.sheet_by_index(0)

            for row_idx in range(0, first_sheet.nrows):
                # Identify the course by the 'Type' text in column 'A'
                if first_sheet.cell(row_idx, 0).value == 'Type':
                    course = first_sheet.cell(row_idx, 1).value

                email = first_sheet.cell(row_idx, email_col).value

                # Validate email address to know we're in the student data
                if validate_email(email):
                    firstname = first_sheet.cell(row_idx, firstname_col).value
                    lastname = first_sheet.cell(row_idx, lastname_col).value
                    score = first_sheet.cell(row_idx, score_col).value
                    student = Student(firstname, lastname, email, score, course)
                    students.append(student)

    return students