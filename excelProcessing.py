from PySide6.QtCore import QThread, QDate
from openpyxl import load_workbook
class ExcelTread(QThread):
    def __init__(self, mainWindow):
        super().__init__()
        self.mainWindow = mainWindow
        self.objects = 0
        self.ln = 0
    def setObjects(self, objects, ln=0):
        self.objects = objects
        self.ln = ln

    def run(self):
        formDate = self.mainWindow.dateEdit.date()        
        for index, obj in enumerate(self.objects):
            # print(obj.getInstituteName(), obj.getCode(), obj.getProfile(), obj.getDateStart())
            startDate =  QDate(int(obj.getDateStart()), 9, 1)
            if startDate >= formDate:
                return
            # obj.getView()
            dt = startDate.daysTo(formDate)            
            print(dt, dt // 365 + 1, (dt % 365) // 153) # Високосный???
            self.mainWindow.progressBar.setValue(int(((index + 1)/self.ln) * 100)) 
class ExcelProccessing():
    def __init__(self, filename):
        self.filename = filename
        self.book = load_workbook(filename)
        self.titleList = self.book['Титул']
        self.practiceList = self.book['Практики']
    def getInstituteName(self):
        return self.titleList['D38'].value # С заглавной??? 
    def getCode(self):
        return self.titleList['D27'].value
    def getProfile(self):
        return self.titleList['D30'].value
    def getDateStart(self):
        return self.titleList['W40'].value
    def getView(self):
        practiceView = []
        practiceTyp = []
        practiceSem = []
        for i in range(3, self.practiceList.max_row + 1):
            view = self.practiceList['A' + str(i)].value
            typ = self.practiceList['B' + str(i)].value
            course = self.practiceList['D' + str(i)].value
            sem = self.practiceList['E' + str(i)].value
            if view:
                if "Вид" in view:
                    practiceView.append(view[14:])
            if typ:
                practiceTyp.append(typ)
            if course:
                practiceSem.append((int(course) - 1) * 2 + sem) 
        practiceList = []
        for i in range(len(practiceView)):
            practiceList.append([practiceView[i], practiceTyp[i], practiceSem[i]])
        print(practiceList)
        return practiceList        
      
