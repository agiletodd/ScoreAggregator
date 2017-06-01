import unittest
from Reader import ReadStudentData
from Student import Student

class TestReadData(unittest.TestCase):
    def setUp(self):
        read_file_directory = 'TestData/'
        self.students = ReadStudentData(read_file_directory)

    def test_read_course_details(self):
        self.assertIsNotNone(self.students)

    def test_find_student_by_email(self):
        self.assertIsNotNone(Student.find_by_email('bostm@ctc.com'))

suite = unittest.TestLoader().loadTestsFromTestCase(TestReadData)
unittest.TextTestRunner(verbosity=2).run(suite)