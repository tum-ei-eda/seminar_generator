#!/usr/bin/env python3

import argparse
import pathlib
import csv
import os
import sys
import re
import pickle

from xlrd import open_workbook,xldate_as_tuple

from components.Contributors import Student
from components.Contributors import Advisor
from components.Contributors import Seminar

from components.GradingSheet import EmptyGradingSheet
from components.GradingReport import GradingReport
from components.Agenda import AgendaPlanning
from components.Agenda import Agenda

def checkSpecialCharacter(str_):
    
    if str_ in ("-"):
        return False
    
    pattern = r"[^a-zA-Z\s]"
    return bool(re.search(pattern, str_))

def importSeminar(inDir_):
    
    print("Reading input directory: " + inDir_)
    inDir = pathlib.Path(inDir_).resolve()
    if not inDir.is_dir():
        print("ERROR: Input directory \'" + str(inDir) + "\' does not exist!")
        sys.exit()
    
    csvFiles = [x for x in inDir.glob("**/*") if (x.is_file and x.suffix==".csv")]
    if len(csvFiles) > 1:
        print("ERROR: More than one csv file located in \'" + str(inDir) + "\'")
        sys.exit()
    if len(csvFiles) < 1:
        print("ERROR: Failed to localize csv file in \'" + str(inDir) + "\'")
        sys.exit()
    fileContent = []
    with csvFiles[0].open('r') as f:
        for line_i in f:
            fileContent.append(line_i)


    print("Processing file content...")
    
    seminar = Seminar(inDir)

    csvDict = csv.DictReader(fileContent, delimiter=',')
    for student_i in csvDict:

        if student_i["PLACE"] == "Confirmed place":

            matNr_str = student_i["MATRICULATION NUMBER"]
            if matNr_str.startswith(""):
                matNr_str = matNr_str[1:]
            matNr = int(matNr_str)

            print(" > Adding student " + student_i["LAST NAME"] + " (" + str(matNr) + ")")

            # Read out advisor and topic
            advisor = student_i["NOTE"].split(';')[0]
            topic = student_i["NOTE"].split(';')[1]

            # Create student object and add to dictionary
            student = Student(  matNr, 
                                student_i["FIRST NAME"].replace(" ",""), 
                                student_i["LAST NAME"].replace(" ",""),
                                student_i["EMAIL"], 
                                topic, 
                                advisor)
            seminar.addStudent(student)
            seminar.createOrUpdateAdvisor(student)

    # Check student names for special characters
    for student_i in seminar.getStudentList():
        isNew = False
        if checkSpecialCharacter(student_i.firstName):
            student_i.updateFirstName(input(f"Student {student_i.fullName}: First name ({student_i.firstName}) contains unsupported characters. Please enter a version containing standard letters (a-zA-Z):"))
            isNew = True
        if checkSpecialCharacter(student_i.lastName):
            student_i.updateLastName(input(f"Student {student_i.fullName}: Last name ({student_i.lastName}) contains unsupported characters. Please enter a version containing standard letters (a-zA-Z):"))
            isNew = True
        if isNew:
            print(f"New name: {student_i.fullName}")

    # Create sheet for agenda planning
    print("Create agenda-planning sheet...")
    agendaPlan = AgendaPlanning()
    agendaPlan.create(seminar)
    agendaPlan.print(inDir_)

    # Store seminar object and return
    objFile = inDir / "seminar.pkl"
    with objFile.open("wb") as f:
        pickle.dump(seminar, f)

    return seminar


def loadSeminar(inDir_):
    
    print("Reading input directory: " + inDir_)
    inDir = pathlib.Path(inDir_).resolve()
    if not inDir.is_dir():
        print("ERROR: Input directory \'" + str(inDir) + "\' does not exist!")
        sys.exit()
    
    objFile = inDir / "seminar.pkl"
    if not objFile.is_file():
        print("ERROR: No stored object seminar.pkl. Use --new to import a new list of students!")
        sys.exit()

    with objFile.open("rb") as f:
        seminar = pickle.load(f)

    return seminar


def createAgenda(seminar_):

    agendaFile = seminar_.getTargetDir() / 'Agenda_Planning.xls'
    if not agendaFile.is_file():
        print("Error: Agenda.xls does not exist. Import a new student list to generate this file!")
        sys.exit()

    print("Read in session-plan...")
    # Open and read out examiner name
    wb = open_workbook(str(agendaFile))
    planning_sheet = wb.sheet_by_name('Planning')

    sessions={}
    row = 1 # Row offset for first student
    while(row < planning_sheet.nrows):
        matNr = int(planning_sheet.cell(row,1).value)
        sessionNr_str = planning_sheet.cell(row,4).value
        sessionIndex_str = planning_sheet.cell(row,5).value

        if sessionNr_str == "":
            print(f"ERROR: No session specified in session-plan row {row}. Assign sessions to all students before running \"create\".")
            sys.exit()
        sessionNr = int(sessionNr_str)

        if sessionIndex_str == "":
            sessionIndex = 100
        else:
            sessionIndex = int(sessionIndex_str)

        if not sessionNr in sessions:
            sessions[sessionNr] = []

        sessions[sessionNr].append((sessionIndex, seminar_.getStudent(matNr)))

        row += 1

    # Sort talks according to session-index:
    for nr_i, session_i in sessions.items():
        sortedTalks = [student for _, student in sorted(session_i, key=lambda x: x[0])]
        sessions[nr_i] = sortedTalks

    # Convert session dictionary into list
    sessionList = [value for key, value in sorted(sessions.items())]
    seminar_.setSessions(sessionList)

    print("Create agenda...")
    agenda = Agenda()
    agenda.create(seminar_)
    agenda.print(seminar_.getTargetDir())


def createGradingSheets(seminar_):

    print("Creating folders...")
    seminar_.getEmptyGradingSheetsDir().mkdir(parents=True, exist_ok=True)
    seminar_.getFilledGradingSheetsDir().mkdir(parents=True, exist_ok=True)

    print("Creating grading-sheets...")
    for advisor_i in seminar_.getAdvisorList():
        print(" > " + advisor_i.lastName)
                
        studentList = []
        for student_i in advisor_i.students:
            studentList.append(seminar_.getStudent(student_i))

        gradingSheet = EmptyGradingSheet(advisor_i.lastName)
        gradingSheet.createOverviewSheet()
        gradingSheet.createPaperSheet(studentList)
        for session_i in seminar_.getSessions():
            gradingSheet.createSessionSheet(session_i)
        gradingSheet.print(seminar_.getEmptyGradingSheetsDir())

    #Create default grading-sheet 'LastName'
    gradingSheet = EmptyGradingSheet('LastName')
    gradingSheet.createOverviewSheet()
    for session_i in seminar_.getSessions():
        gradingSheet.createSessionSheet(session_i)
    gradingSheet.print(seminar_.getEmptyGradingSheetsDir())
    
    
def createGradingReport(seminar_):

    print("Importing filled grading sheets from '" + str(seminar_.getFilledGradingSheetsDir()) + "'...")
    if not seminar_.getFilledGradingSheetsDir().is_dir():
        print("ERROR: Directory \'" + str(seminar_.getFilledGradingSheetsDir()) + "\' does not exist!")
        sys.exit()
    FilledGradingSheetFiles = [x for x in seminar_.getFilledGradingSheetsDir().glob("**/*")]
    
    examinerList = []

    for file_i in FilledGradingSheetFiles:
        print(" > Processing \'" + str(file_i) +"\'")

        # Open and read out examiner name
        wb = open_workbook(file_i)
        overview_sheet = wb.sheet_by_name('Overview')
        examiner = overview_sheet.cell(1,2).value
        examinerList.append(examiner)

        for sheet_i in wb.sheet_names():

            print(sheet_i)
            if sheet_i == "Overview":
                continue
            
            # Read paper grade
            if sheet_i == "Paper Grading":
                paper_sheet = wb.sheet_by_name("Paper Grading")

                # Find advisor
                advisor = seminar_.getAdvisor(examiner)

                row = 10 # Row offset for first paper. Make this less implicit
                #while(row < paper_sheet.nrows):
                for i in range(advisor.getNumStudents()):
                    print("ping")

                    matNr = paper_sheet.cell(row+1,2).value
                    paperGrade = int(paper_sheet.cell(row+2,1).value)
                    student = seminar_.getStudent(matNr)
                    if (paperGrade > 12) or (paperGrade < 0):
                        print("ERROR: Examiner \'" + examiner + "\' attempts to give illegal paper grade (" + str(paperGrade) + ") to \'" + student.fullName + " [" + str(student.matNr) + "]\'")
                        sys.exit()
                    student.addPaperGrade(paperGrade)

                    row = row+6 # Increase row to next paper. Make this less implicit

            # Read presentation grades
            if sheet_i.startswith("Session"):
                session_sheet = wb.sheet_by_name(sheet_i)

                row = 13 # Row offset for first presentation. Make this less implicit
                while(row < session_sheet.nrows):

                    matNr = session_sheet.cell(row+1,2).value
                    if not ((session_sheet.cell(row+2,1).value == "<ENTER POINTS (12-0)>" or session_sheet.cell(row+3,1).value == "<ENTER POINTS (12-0)>") \
                            or (session_sheet.cell(row+2,1).value == "" or session_sheet.cell(row+3,1).value == "")):
                        styleGrade = int(session_sheet.cell(row+2,1).value)
                        contentGrade = int(session_sheet.cell(row+3,1).value)
                        student = seminar_.getStudent(matNr)
                        if (styleGrade > 12) or (styleGrade < 0):
                            print("ERROR: Examiner \'" + examiner + "\' attempts to give illegal presentation style-grade (" + str(styleGrade) + ") to \'" + student.fullName + " [" + str(student.matNr) + "]\'")
                            sys.exit()
                        if (contentGrade > 12) or (contentGrade < 0):
                            print("ERROR: Examiner \'" + examiner + "\' attempts to give illegal presentation content-grade (" + str(contentGrade) + ") to \'" + student.fullName + " [" + str(student.matNr) + "]\'")
                            sys.exit()
                        student.addPresentationGrade(examiner, styleGrade, contentGrade)

                    row = row+6 # Increase row to next presentation. Make this less implicit


    print("Creating grading-report...")
    gradingReport = GradingReport(seminar_.getStudentDict())

    for examiner_i in examinerList:
        gradingReport.addExaminer(examiner_i)

    #for advisor_i in seminar_.getAdvisorList():
    #    gradingReport.addExaminer(advisor_i.lastName)

    ##gradingReport.addExaminer('Graeb')
    #gradingReport.addExaminer('Foik')
    #gradingReport.addExaminer('Gerl')
    #gradingReport.addExaminer('Prebeck')

    gradingReport.print(seminar_.getTargetDir())

if __name__ == '__main__':

    argParser = argparse.ArgumentParser()
    argParser.add_argument("input_dir", help="Path to input directory containing the student list csv-file")
    argParser.add_argument("--new", "-n", action="store_true", help="Import new student list (csv-file)")
    argParser.add_argument("--create", "-c", action="store_true", help="Create empty grading sheets")
    argParser.add_argument("--grade", "-g", action="store_true", help="Read filled grading sheets and create grading report")
    args = argParser.parse_args()

    inputDir=args.input_dir

    if args.new:
        seminar = importSeminar(inputDir)
    else:
        seminar = loadSeminar(inputDir)

    if args.create:
        createAgenda(seminar)
        createGradingSheets(seminar)
    if args.grade:
        createGradingReport(seminar)

    ## Capture script
    #while True:
    #    pass