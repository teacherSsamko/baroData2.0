from openpyxl import Workbook, load_workbook

class BaroRoll:
    def __init__(self, wb):
        self.wb = wb
        self.sheet = wb[wb.sheetnames[0]]
        self.count_students = self.sheet.max_row - 1
        self.count_subjects = self.sheet.max_column - 1
        self.student_names = []
        

    
    
