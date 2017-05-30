from openpyxl import load_workbook, worksheet

def WriteData(filename, students):
    firstname_col = 1
    lastname_col = 2
    email_col = 3
    score_col = 14
    course_col = 11

    wb = load_workbook(filename = filename)
    sheet = wb.get_sheet_by_name('Attendees')

    for row in sheet.iter_rows():
        c = sheet['B4'].value
        print(c)
        print(row)

