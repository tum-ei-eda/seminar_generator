#!/usr/bin/env python3

import argparse
import pathlib
import csv
import os
import sys

from xlrd import open_workbook,xldate_as_tuple

from components.Contributors import Student
from components.Contributors import Advisor
from components.Contributors import Seminar

from components.GradingSheet import EmptyGradingSheet
from components.GradingReport import GradingReport

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

            # TODO: MatNr in CSV has a starting white space, which messes up cast to int. Currently removed manually. Find way to do this automatically!
            matNr = int(student_i["MATRICULATION NUMBER"])

            print(" > Adding student " + student_i["LAST NAME"] + " (" + str(matNr) + ")")

            # Read out advisor and topic
            advisor = student_i["NOTE"].split(';')[0]
            topic = student_i["NOTE"].split(';')[1]

            # Create student object and add to dictionary
            student = Student(  matNr, 
                                student_i["FIRST NAME"], 
                                student_i["LAST NAME"], # TODO: Special characters and white space not supported! Currently removed by hand in csv. Automate!!
                                student_i["EMAIL"], 
                                topic, 
                                advisor)
            seminar.addStudent(student)
            seminar.createOrUpdateAdvisor(student)

    print("Planning sessions...")

    # TODO: Hardcoded session planning. Automate!
    sessions = []

    session_1 = []
    session_1.append(seminar.getStudent(int("03710411"))) # Simson
    session_1.append(seminar.getStudent(int("03750080"))) # Deng
    session_1.append(seminar.getStudent(int("03765600"))) # Wang
    sessions.append(session_1)

    session_2 = []
    session_2.append(seminar.getStudent(int("03767986"))) # Berger
    session_2.append(seminar.getStudent(int("03677564"))) # Singh
    session_2.append(seminar.getStudent(int("03765678"))) # Song
    sessions.append(session_2)

    session_3 = []
    session_3.append(seminar.getStudent(int("03711460"))) # Häringer
    session_3.append(seminar.getStudent(int("03774286"))) # Dickgießer
    session_3.append(seminar.getStudent(int("03763760"))) # Huang
    sessions.append(session_3)

    seminar.setSessions(sessions)

    return seminar


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

            if sheet_i == "Overview":
                continue
            
            # Read paper grade
            if sheet_i == "Paper Grading":
                paper_sheet = wb.sheet_by_name("Paper Grading")

                row = 10 # Row offset for first paper. Make this less implicit
                while(row < paper_sheet.nrows):
                    
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
                    if not (session_sheet.cell(row+2,1).value == "<ENTER POINTS (12-0)>" or session_sheet.cell(row+3,1).value == "<ENTER POINTS (12-0)>"):
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
    argParser.add_argument("--create", "-c", action="store_true", help="Create empty grading sheets")
    argParser.add_argument("--grade", "-g", action="store_true", help="Read filled grading sheets and create grading report")
    args = argParser.parse_args()

    inputDir=args.input_dir

    seminar = importSeminar(inputDir)

    if args.create:
        createGradingSheets(seminar)
    if args.grade:
        createGradingReport(seminar)

    ## Capture script
    #while True:
    #    pass