#!/usr/bin/env python3

from xlwt import Workbook, XFStyle, Borders, Pattern, Font, Alignment, Utils, Formula,easyxf
from xlrd import open_workbook

# TODO: Use openpyxl instead, to be able to have agenda_planning and agenda_auto sheets in common workbook

class AgendaConfig:

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

class AgendaPlanning:

    def __init__(self):

        self.sheet = Workbook(encoding='cp1252')
        self.cfg = AgendaConfig()

    def create(self, seminar_):

        ags = self.sheet.add_sheet('Planning')
        ags.protect = False

        ags.write(0,0,"Student")
        ags.write(0,1,"Matr.Nr.")
        ags.write(0,2,"Topic")
        ags.write(0,3,"Advisor")
        ags.write(0,4,"Session")
        ags.write(0,5,"Session Order (optional)")

        row = 0
        for student_i in seminar_.getStudentList():
            row += 1
            ags.write(row, 0, student_i.fullName)
            ags.write(row, 1, student_i.matNr)
            ags.write(row, 2, student_i.topic)
            ags.write(row, 3, student_i.advisor)

    def print(self, outDir_):
        fileName=str(outDir_) + '/Agenda_Planning.xls'
        self.sheet.save(fileName)

class Agenda:

    def __init__(self):

        self.sheet = Workbook(encoding='cp1252')
        self.cfg = AgendaConfig()

    def create(self, seminar_):
        
        ags = self.sheet.add_sheet('Agenda_auto')
        ags.protect = False

        # TODO: Adjust style and add images
        ags.write_merge(0,0,0,8,"Agenda for Hauptseminar VLSI-Entwurfsverfahren")
        ags.write_merge(1,1,1,7,"<TERM> - <DATE>")
        
        ags.write_merge(3,3,1,3,"Time")
        ags.write(3,5,"Session and Session Chair")
        ags.write(3,7,"Speaker")

        ags.write(5,2,"-")
        ags.write(5,5,"Welcome (<TBD>)")

        isFirst = True
        row = 7
        sessionNr = 1
        for session_i in seminar_.sessions:

            if isFirst:
                isFirst = False
            else:
                ags.write(row,2,"-")
                ags.write(row,5,"Break")
                row += 2
            
            ags.write(row,2,"-")
            sessionTitle = "Session " + str(sessionNr) + ": <TBD> (<TBD>)"
            ags.write(row,5,sessionTitle)
            
            row += 1
            for student_i in session_i:
                talkTitle = str(student_i.talkNr) + ": " + student_i.topic
                ags.write(row,5,talkTitle)
                ags.write(row,7,student_i.fullName)
                row += 1
            
            sessionNr += 1
            row += 1

    def print(self, outDir_):
        fileName=str(outDir_) + '/Agenda.xls'
        self.sheet.save(fileName)