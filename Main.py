from Reader import ReadStudentData
from Updater import Update_Existing_Students

def main():
    #read_file_directory = '/Users/toddmiller/Downloads/'
    #write_file = '/Users/toddmiller/OneDrive Xperient/OneDrive - Xperient Software LLC/Customers.xlsx'
    read_file_directory = 'TestData/'
    write_file = 'TestData/file.xlsx'

    students = ReadStudentData(read_file_directory)

    Update_Existing_Students(write_file, students)

if __name__ == "__main__":
    main()
