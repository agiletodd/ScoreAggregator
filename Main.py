from Reader import ReadStudentData
from Writer import WriteData

def main():
    read_file_directory = '/Users/toddmiller/Downloads/'
    write_file = '/Users/toddmiller/OneDrive Xperient/OneDrive - Xperient Software LLC/Customers.xlsx'
    #read_file_directory = 'C:/Users/tmiller/Downloads/'
    #write_file = 'C:/Users/tmiller/Downloads/file.xlsx'

    students = ReadStudentData(read_file_directory)

    WriteData(write_file, students)

if __name__ == "__main__":
    main()
