#!/usr/bin/env python3

import argparse
import pathlib
import csv
import os
import sys

from components.Contributors import Student
from components.Contributors import Advisor

from components.GradingSheet import GradingSheet
from components.GradingReport import GradingReport


def main(inDir_):
    
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

    emptyGradingSheetsDir = inDir / "EmptyGradingSheets"
    emptyGradingSheetsDir.mkdir(parents=True, exist_ok=True)
    filledGradingSheetsDir = inDir / "FilledGradingSheets"
    filledGradingSheetsDir.mkdir(parents=True, exist_ok=True)

    print("Processing file content...")
    
    studentDict = {}
    advisorDict = {}
    
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
            studentDict[matNr] = student

            # Update existing advisor or create new one
            if advisor not in advisorDict:
                advisorDict[advisor] = Advisor(advisor)
            advisorDict[advisor].addStudent(matNr)

    print("Planning sessions...")

    # TODO: Hardcoded session planning. Automate!
    sessions = []

    session_1 = []
    session_1.append(studentDict[int("03710411")]) # Simson
    session_1.append(studentDict[int("03750080")]) # Deng
    session_1.append(studentDict[int("03765600")]) # Wang
    sessions.append(session_1)

    session_2 = []
    session_2.append(studentDict[int("03767986")]) # Berger
    session_2.append(studentDict[int("03677564")]) # Singh
    session_2.append(studentDict[int("03765678")]) # Song
    sessions.append(session_2)

    session_3 = []
    session_3.append(studentDict[int("03711460")]) # Häringer
    session_3.append(studentDict[int("03774286")]) # Dickgießer
    session_3.append(studentDict[int("03763760")]) # Huang
    sessions.append(session_3)
    '''

    session_1 = []
    session_1.append(studentDict[int("03765937")]) # Kedia
    session_1.append(studentDict[int("03767522")]) # Cakir
    session_1.append(studentDict[int("03750993")]) # Chakraborty
    session_1.append(studentDict[int("03771336")]) # Li
    sessions.append(session_1)

    session_2 = []
    session_2.append(studentDict[int("03766837")]) # Gowda
    session_2.append(studentDict[int("03767345")]) # Noor
    sessions.append(session_2)
    '''

    print("Creating grading-sheets...")
    for advisor_i in advisorDict.values():
        print(" > " + advisor_i.lastName)
        #print(" >  " + outDir)
                
        studentList = []
        for student_i in advisor_i.students:
            studentList.append(studentDict[student_i])

        gradingSheet = GradingSheet(advisor_i.lastName)
        gradingSheet.createOverviewSheet()
        gradingSheet.createPaperSheet(studentList)
        for session_i in sessions:
            gradingSheet.createSessionSheet(session_i)
        gradingSheet.print(emptyGradingSheetsDir)

    #Create default grading-sheet 'LastName'
    gradingSheet = GradingSheet('LastName')
    gradingSheet.createOverviewSheet()
    for session_i in sessions:
        gradingSheet.createSessionSheet(session_i)
    gradingSheet.print(emptyGradingSheetsDir)
    

    print("Creating grading-report...")
    gradingReport = GradingReport(studentDict)

    for advisor_i in advisorDict.values():
        gradingReport.addExaminer(advisor_i.lastName)

    #gradingReport.addExaminer('Graeb')
    gradingReport.addExaminer('Foik')
    gradingReport.addExaminer('Gerl')
    gradingReport.addExaminer('Prebeck')

    gradingReport.print(inDir)

if __name__ == '__main__':

    argParser = argparse.ArgumentParser()
    argParser.add_argument("input_dir", help="Path to input directory containing the student list csv-file")
    args = argParser.parse_args()

    inputDir=args.input_dir

    main(inputDir)

    ## Capture script
    #while True:
    #    pass