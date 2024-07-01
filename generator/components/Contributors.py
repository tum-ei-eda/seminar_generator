#!/usr/bin/env python3

class Student:

    def __init__(self, matNr_, firstName_, lastName_, email_, topic_, advisor_):
        self.matNr = matNr_
        self.firstName = firstName_
        self.lastName = lastName_
        self.fullName = firstName_ + " " + lastName_
        self.email = email_
        self.topic = topic_
        self.advisor = advisor_
        self.grades = Grades(self)

    def addPaperGrade(self, grade_):
        self.grades.addPaper(grade_)

    def getPaperGrade(self):
        return self.grades.getPaper()

    def addPresentationGrade(self, examiner_, style_, content_):
        self.grades.addPresentation(examiner_, style_, content_)

    def presentationGradeExists(self, examiner_):
        return self.grades.presentationExists(examiner_)

    def getPresentationStyleGrade(self, examiner_):
        return self.grades.getPresentationStyle(examiner_)

    def getPresentationContentGrade(self, examiner_):
        return self.grades.getPresentationContent(examiner_)

class Grades:

    def __init__(self, student_):
        self.student = student_
        self.paper = 0
        self.paperGradeSet = False
        self.presentationDict = {}

    def addPaper(self, paper_):
        if not self.paperGradeSet:
            self.paper = paper_
            self.paperGradeSet = True
        else:
            print("ERROR: Attempting to override paper grade for \'" + self.student.fullName + "\'. THIS SHOULD NEVER HAPPEN!")
            sys.exit()

    def getPaper(self):
        return self.paper

    def addPresentation(self, examiner_, style_, content_):
        presGrade = PresentationGrades(style_, content_)
        if examiner_ in self.presentationDict:
            print("ERROR: Cannot add presentation grades of \'" + examiner_ + "\' to student \'" + self.student.fullName + "\' as there are already grades for this examiner. THIS SHOULD NEVER HAPPEN!")
            sys.exit()
        self.presentationDict[examiner_] = presGrade

    def presentationExists(self, examiner_):
        if examiner_ in self.presentationDict:
            print(self.student.fullName + "[" + examiner_ + "]")
            return True
        return False

    def getPresentationStyle(self, examiner_):
        return self.presentationDict[examiner_].style

    def getPresentationContent(self, examiner_):
        return self.presentationDict[examiner_].content

class PresentationGrades:

    def __init__(self, style_, content_):
        self.style = style_
        self.content = content_

class Advisor:

    def __init__(self, lastName_):
        self.lastName = lastName_
        self.students = []

    def addStudent(self, matNr_):
        self.students.append(int(matNr_))

class Seminar:

    def __init__(self, dir_):
        self.targetDir = dir_
        self.emptyGradingSheetsDir = self.targetDir / "EmptyGradingSheets"
        self.filledGradingSheetsDir = self.targetDir / "FilledGradingSheets"

        self.studentDict = {}
        self.advisorDict = {}
        self.sessions = []

    def getTargetDir(self):
        return self.targetDir

    def getEmptyGradingSheetsDir(self):
        return self.emptyGradingSheetsDir

    def getFilledGradingSheetsDir(self):
        return self.filledGradingSheetsDir

    def addStudent(self, student_):
        self.studentDict[student_.matNr] = student_

    def getStudent(self, matNr_):
        return self.studentDict[matNr_]

    def getStudentByName(self, name_):
        for stud_i in self.dictionary.values():
            if stud_i.fullName == name_:
                return stud_i
        return None

    def getStudentDict(self):
        return self.studentDict

    def createOrUpdateAdvisor(self, student_):
        if student_.advisor not in self.advisorDict:
            self.advisorDict[student_.advisor] = Advisor(student_.advisor)
        self.advisorDict[student_.advisor].addStudent(student_.matNr)

    def getAdvisorList(self):
        return self.advisorDict.values()

    def setSessions(self, sessions_):
        self.sessions = sessions_

    def getSessions(self):
        return self.sessions

    def getNumSessions(self):
        return len(self.sessions)