from collections import defaultdict

class Student:
    email_index = defaultdict(list)

    def __init__(self, firstname, lastname, email, score, course):
        self.firstname = firstname
        self.lastname = lastname
        self.email = email
        self.score = score
        self.course = course
        Student.email_index[email].append(self)

    @classmethod
    def find_by_email(cls, email):
        return Student.email_index[email]
