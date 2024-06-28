#!/usr/bin/env python3

class Student:

    def __init__(self, matNr_, firstName_, lastName_, email_, topic_, advisor_):
        self.matNr = matNr_
        self.firstName = firstName_
        self.lastName = lastName_
        self.email = email_
        self.topic = topic_
        self.advisor = advisor_

class Advisor:

    def __init__(self, lastName_):
        self.lastName = lastName_
        self.students = []

    def addStudent(self, matNr_):
        self.students.append(int(matNr_))