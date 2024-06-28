#!/usr/bin/env python3

from xlwt import Workbook, XFStyle, Borders, Pattern, Font, Alignment, Utils, Formula,easyxf

class GradingSheetConfig:

    def __init__(self):
        pointfieldstyle = XFStyle()
        pointfieldstyle.borders.left = Borders.THICK
        pointfieldstyle.borders.right = Borders.THICK
        pointfieldstyle.borders.top = Borders.THICK
        pointfieldstyle.borders.bottom = Borders.THICK
        pointfieldstyle.alignment.vert = Alignment.VERT_TOP
        pointfieldstyle.alignment.wrap = True
        pointfieldstyle.protection.cell_locked = False
        self.pointfieldstyle = pointfieldstyle

        talktitlestyle_gs= XFStyle()
        talktitlestyle_gs.font.bold=True
        self.talktitlestyle_gs = talktitlestyle_gs

        advisorstyle_gs = XFStyle()
        advisorstyle_gs.font.italic=True
        self.advisorstyle_gs = advisorstyle_gs

        sessiontitlestyle_gs = XFStyle()
        sessiontitlestyle_gs.font.bold=True
        self.sessiontitlestyle_gs = sessiontitlestyle_gs

        criteriastyle_gs = XFStyle()
        criteriastyle_gs.alignment.wrap = True
        self.criteriastyle_gs = criteriastyle_gs

        importantnoticestyle_gs= style = easyxf('font: bold 1, color red;')
        self.importantnoticestyle_gs = importantnoticestyle_gs

class GradingSheet:

    def __init__(self, userName_):

        self.userName = userName_
        self.sheet = Workbook(encoding='cp1252')
        self.cfg = GradingSheetConfig()

        self.numSessions = 0

    def createOverviewSheet(self):

        ags = self.sheet.add_sheet('Overview')
        ags.protect = False

        row = 1
        ags.write(row,0,'Graded by:')
        ags.write_merge(row,row,2,4,self.userName)

        row=row+2
        ags.write(row,0,'Steps to do:',self.cfg.sessiontitlestyle_gs)
        row = row+1
        ags.write(row,0,'Step 1')
        ags.write(row,1,'If not already filled, put your last name in field above')
        row = row+1
        ags.write(row,0,'Step 2')
        ags.write(row,1,'Save this file: If not already done, replace <Lastname> in filename with your last name')
        row = row+1
        ags.write(row,0,'Step 3')
        ags.write(row,1,'Use speadsheet \"Paper Grading\" to enter points for the paper of your students')
        row = row+1
        ags.write(row,0,'Step 4')
        ags.write(row,1,'Use speadsheet \"Session X\" to enter points for the talks')
        row = row+1
        ags.write(row,0,'Step 5')
        ags.write(row,1,'Send filled sheet to conrad.foik@tum.de')

        row = row+2
        ags.write(row,0,'Grading scheme:',self.cfg.sessiontitlestyle_gs)
        row = row+1
        ags.write(row,0,'We grade three categorize: 1. Paper, 2. Presentation Style, 3. Presentation Content')
        row = row+1
        ags.write(row,0,'In each category students can reach up to 12 points')
        row = row+1
        ags.write(row,0,'Point scheme: 12=outstanding, 11=very good, 10=good, 9=average, 8=so-so, 7-0=poor')

    def createPaperSheet(self, students_):

        ags = self.sheet.add_sheet('Paper Grading')
        ags.protect = False

        row = 0
        ags.write(row,0,'Grading Sheet for Paper(s)',self.cfg.sessiontitlestyle_gs)

        row = row+2
        ags.write(row,0,'Point scheme:')
        ags.write(row,1,'12=outstanding, 11=very good, 10=good, 9=average, 8=so-so, 7-0=poor')

        row = row+1
        ags.write(row,0,'Criteria:')
        ags.write_merge(row,row+4,1,7,'Is the field completely covered? Are all major principles explained? Has the topic been covered in-depth? Was the technical content covered correctly? Is the formatting correct and clear? Is the writing clear? Has the student developed his own way to explain the topic?')

        row = row+4
        for student_i in students_:
            row = row+3
            row = self.__addPaper(ags, row, student_i)


    def createSessionSheet(self, students_):

        self.numSessions = self.numSessions + 1

        ags = self.sheet.add_sheet(('Session ' + str(self.numSessions)))
        ags.protect = False

        row = 0

        ags.write(row,0,('Grading Sheet for Session ' + str(self.numSessions)),self.cfg.sessiontitlestyle_gs)

        row = row+2
        ags.write(row,0,'Point scheme:')
        ags.write(row,1,'12=outstanding, 11=very good, 10=good, 9=average, 8=so-so, 7-0=poor')

        row = row+2
        ags.write(row,0,'Criteria - Presentation Style:')
        ags.write_merge(row,row+2,1,7,'Slide design, structure, time management, presentation skill, voice, body language, involvment of audience, …')

        row = row+4
        ags.write(row,0,'Criteria - Presentation Content:')
        ags.write_merge(row,row+2,1,7,'Comprehensiveness, Correctness, Completeness, Technical depth, Answers in discussion,…')

        row = row+3
        for talkNr_i, student_i in enumerate(students_):
            row=row+2
            row = self.__addTalk(ags, row, student_i, (talkNr_i + 1))

    def print(self, outDir_):
        fileName=str(outDir_) + '/' + self.userName +  '_GradingSheetSeminar.xls'
        self.sheet.save(fileName)

    def __addPaper(self, ags_, row_, student_):

        row = row_
        ags_.write(row,0,'Paper:')
        ags_.write(row,1,student_.topic)

        row = row+1
        ags_.write(row,0,'Student:')
        ags_.write(row,1,(student_.firstName + " " + student_.lastName))

        row = row+1
        ags_.write(row,0,'Points:')
        ags_.write(row,1,'<ENTER YOUR POINTS HERE>')

        row = row+1
        ags_.write(row,0,'Comment:')
        ags_.write_merge(row,row,1,7, '<ENTER YOUR COMMENTS HERE>')

        return row

    def __addTalk(self, ags_, row_, student_, talkNr_):

        row = row_
        ags_.write(row,0,('Talk ' + str(talkNr_)),self.cfg.sessiontitlestyle_gs)
        ags_.write(row,1,student_.topic,self.cfg.sessiontitlestyle_gs)

        row = row+1
        ags_.write(row,0,'Student:')
        ags_.write(row,1,(student_.firstName + ' ' + student_.lastName))

        row = row+1
        ags_.write(row,0,'Points - Presentation Style:')
        ags_.write(row,1,'<ENTER POINTS (12-0)>')

        row = row+1
        ags_.write(row,0,'Points - Presentation Content:')
        ags_.write(row,1,'<ENTER POINTS (12-0)>')

        row = row+1
        ags_.write(row,0,'Comments:')
        ags_.write(row,1,'<ENTER YOUR COMMENTS>')

        return row