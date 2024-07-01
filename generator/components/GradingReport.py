#!/usr/bin/env python3

from xlwt import Workbook, XFStyle, Borders, Pattern, Font, Alignment, Utils, Formula,easyxf

# TODO: Make common config for all sheets?
class GradingReportConfig:

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


class GradingReport:

    def __init__(self, studentDict_):

        self.sheet = Workbook(encoding='cp1252')
        self.cfg = GradingReportConfig()
        self.studentDict = studentDict_

        self.overview = None
        self.paperRow = 0
        self.paperCol = 0
        self.talkStyleRow = 0
        self.talkStyleCol = 0
        self.talkContentRow = 0
        self.talkContentCol = 0

        self.examiners = []

    def addExaminer(self, examiner_):
        self.examiners.append(examiner_)


    def print(self, outDir_):

        self.__createOverviewSheet()

        for student_i in self.studentDict.values():
            self.__createStudentSheet(student_i)

        self.__finalizeOverviewSheet()

        fileName=str(outDir_) + '/GradingReport.xls'
        self.sheet.save(fileName)


    def __createOverviewSheet(self):

        ags = self.sheet.add_sheet('Overview')
        ags.protect = False

        row = 2
        ags.write(row,0,'Grading Scheme:',self.cfg.sessiontitlestyle_gs)
        row = row+1
        ags.write(row,0,'Points')
        ags.write(row,1,'Grade')
        row = row+1
        ags.write(row,0,'0')
        ags.write(row,1,'5,0')
        row = row+1
        ags.write(row,0,'1')
        ags.write(row,1,'4,7')
        row = row+1
        ags.write(row,0,'3')
        ags.write(row,1,'4,3')
        row = row+1
        ags.write(row,0,'5')
        ags.write(row,1,'4,0')
        row = row+1
        ags.write(row,0,'7')
        ags.write(row,1,'3,7')
        row = row+1
        ags.write(row,0,'9')
        ags.write(row,1,'3,3')
        row = row+1
        ags.write(row,0,'11')
        ags.write(row,1,'3,0')
        row = row+1
        ags.write(row,0,'13')
        ags.write(row,1,'2,7')
        row = row+1
        ags.write(row,0,'15')
        ags.write(row,1,'2,3')
        row = row+1
        ags.write(row,0,'17')
        ags.write(row,1,'2,0')
        row = row+1
        ags.write(row,0,'19')
        ags.write(row,1,'1,7')
        row = row+1
        ags.write(row,0,'21')
        ags.write(row,1,'1,3')
        row = row+1
        ags.write(row,0,'23')
        ags.write(row,1,'1,0')

        self.overview = ags


    def __finalizeOverviewSheet(self):

        ags = self.overview
        
        row = 2
        col = 5
        ags.write_merge(row, row+1,col,col,'First Name')
        col = col+1
        ags.write_merge(row, row+1,col,col,'Last Name')
        col = col+1
        ags.write_merge(row, row+1,col,col,'Matr.Nr.')
        col = col+1
        ags.write_merge(row, row+1,col,col,'Topic')
        col = col+1
        ags.write_merge(row, row+1,col,col,'Advisor')
        col = col+1
        ags.write_merge(row,row,col,col+2,'Written')
        ags.write(row+1,col,'Paper')
        col = col+1
        ags.write(row+1,col,'Sum')
        col = col+1
        ags.write(row+1,col,'Grade')
        col = col+1
        ags.write_merge(row,row,col,col+3,'Presentation')
        ags.write(row+1,col,'Style')
        col = col+1
        ags.write(row+1,col,'Content')
        col = col+1
        ags.write(row+1,col,'Sum')
        col = col+1
        ags.write(row+1,col,'Grade')
        col = col+1
        ags.write_merge(row, row+1,col,col,'Overall Grade')

        row = row +1
        for student_i in self.studentDict.values():
            row = row +1
            col = 5

            ags.write(row,col,student_i.firstName)
            col = col+1
            ags.write(row,col,student_i.lastName)
            col = col+1
            ags.write(row,col,student_i.matNr)
            col = col+1
            ags.write(row,col,student_i.topic)
            col = col+1
            ags.write(row,col,student_i.advisor)

            # Get paper points and calculate grade
            col = col+1
            ags.write(row,col,Formula(self.__getExcelLink(student_i.lastName,self.paperRow, self.paperCol)))
            paperLinkCell = self.__getExcelCellName(row,col)
            col = col+1
            ags.write(row,col,Formula("IF(" + paperLinkCell + "=\"ne\";\"ne\";2*" + paperLinkCell +")"))
            col = col + 1
            #TODO :Add look-up to grading table
            paperGradeCell = self.__getExcelCellName(row,col)

            # Get talk points and calculate grade
            col = col + 1
            ags.write(row,col,Formula(self.__getExcelLink(student_i.lastName,self.talkStyleRow,self.talkStyleCol)))
            talkStyleCell = self.__getExcelCellName(row,col)
            col = col +1
            ags.write(row,col,Formula(self.__getExcelLink(student_i.lastName,self.talkContentRow,self.talkContentCol)))
            talkContentCell = self.__getExcelCellName(row,col)
            col = col + 1
            ags.write(row,col,Formula("IF(" + talkStyleCell + "=\"ne\";\"ne\";IF(" + talkContentCell + "=\"ne\";\"ne\";" + talkStyleCell + "+" + talkContentCell + "))"))
            col = col + 1
            # TODO: Add look-up to grading table
            talkGradeCell = self.__getExcelCellName(row,col)

            # Calculate final grade
            col = col + 1
            ags.write(row,col,Formula("IF(" + paperGradeCell + "=\"ne\";\"ne\";IF(" + talkGradeCell + "=\"ne\";\"ne\";ROUNDDOWN(AVERAGE(" + paperGradeCell + ";" + talkGradeCell +");1)))"))

            
    def __createStudentSheet(self, student_):
        
        # TODO: Last name unique enough?
        ags = self.sheet.add_sheet((student_.lastName))
        ags.protect = False

        row = 1
        ags.write(row,0,'Name:')
        ags.write(row,1,(student_.firstName + " " + student_.lastName))
        row = row+1
        ags.write(row,0,'Topic:')
        ags.write(row,1,student_.topic)
        row = row+1
        ags.write(row,0,'Advisor:')
        ags.write(row,1,student_.advisor)

        row = row+2
        ags.write(row,0,'Paper:')
        ags.write(row,1,student_.getPaperGrade())
        self.paperRow = row
        self.paperCol = 1

        row = row+2
        ags.write(row,0,'Präsentation')
        row = row+1
        ags.write(row,2,'Style')
        ags.write(row,3,'Content')

        self.talkStyleCol = 2
        self.talkContentCol = 3

        isFirst = True
        for examiner_i in self.examiners:
            row = row + 1
            ags.write(row,1,examiner_i)
            if student_.presentationGradeExists(examiner_i):
                ags.write(row,2,student_.getPresentationStyleGrade(examiner_i))
                ags.write(row,3,student_.getPresentationContentGrade(examiner_i))
            if isFirst:
                firstExaminerRow = row
                isFirst = False
        lastExaminerRow = row

        firstStyleGradeCell = self.__getExcelCellName(firstExaminerRow,2)
        lastStyleGradeCell = self.__getExcelCellName(lastExaminerRow,2)

        firstContentGradeCell = self.__getExcelCellName(firstExaminerRow,3)
        lastContentGradeCell = self.__getExcelCellName(lastExaminerRow,3)

        row = row + 1
        ags.write(row,2,Formula("AVERAGE(" + firstStyleGradeCell + ":" + lastStyleGradeCell + ")"))
        ags.write(row,3,Formula("AVERAGE(" + firstContentGradeCell + ":" + lastContentGradeCell + ")"))

        self.talkStyleRow = row
        self.talkContentRow = row

    #def __getColLetter(self, int_):
    #    return chr(ord('@')+int_)

    def __getExcelCellName(self, row_, col_):
        return chr(ord('@')+(col_+1)) + str(row_+1)

    def __getExcelLink(self, name_, row_, col_):
        return (name_ + "!" + self.__getExcelCellName(row_,col_))