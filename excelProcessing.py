from PySide6.QtCore import QThread, QDate
from openpyxl import load_workbook
cellColumn = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH"] 
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
            startDate =  QDate(int(obj.getDateStart()), 9, 1) # год из таблицы, месяц 09, день 01
            if startDate >= formDate:
                return
            # obj.getView()
            dt = startDate.daysTo(formDate)  
            # print(obj.getInstituteName(), obj.getCode(), obj.getProfile(), obj.getView()) 
            print(index + 1, obj.getView())          

            # print(dt // 365 + 1, (dt % 365) // 153) # Високосный???
            self.mainWindow.progressBar.setValue(int(((index + 1)/self.ln) * 100)) 
class ExcelProccessing():
    def __init__(self, filename):
        self.filename = filename
        self.book = load_workbook(filename)
        self.titleList = self.book['Титул']
        self.practiceList = self.book['Практики']
        self.planList = self.book['План']

    def getInstituteName(self):
        return self.titleList['D38'].value # С заглавной??? 
    def getCode(self):
        return self.titleList['D29'].value
    def getProfile(self):
        return self.titleList['D30'].value
    def getDateStart(self):
        return self.titleList['W40'].value
    def getView(self):
        practiceView = []
        practiceTyp = []
        practiceSem = []
        double = 0
        for name in cellColumn:
            if self.practiceList[name + "1"].value == "Курс":
                nameCourseColumn = name
                break
        nameSemColumn = cellColumn[cellColumn.index(nameCourseColumn) + 1]
        for i in range(3, self.practiceList.max_row + 1):
            view = self.practiceList["A" + str(i)].value # Нужно искать буквы!
            typ = self.practiceList["B" + str(i)].value
            course = self.practiceList[nameCourseColumn + str(i)].value
            sem = self.practiceList[nameSemColumn + str(i)].value
            if view:
                if "Вид" in view:
                    practiceView.append(view[14:]) # Убирает "Вид практики: "
                    double = 0
            if typ:
                practiceTyp.append(typ)
                double += 1
                if double > 1:
                    practiceView.append(practiceView[-1])
            if course:
                practiceSem.append((int(course) - 1) * 2 + sem) 
        practiceList = []
        zeColumn = "Неопределено"
        nameColumn = "Неопределено"
        hoursColumn = "Неопределено" 
        for col in cellColumn:
            if self.planList[col + "2"].value == "Наименование" or self.planList[col + "3"].value == "Наименование" or self.planList[col + "4"].value == "Наименование": 
                nameColumn = col 
            elif self.planList[col + "1"].value == "з.е." or self.planList[col + "2"].value == "з.е." or self.planList[col + "3"].value == "з.е.": # Брать экспертное или фактическое
                zeColumn = col 
            elif self.planList[col + "1"].value == "Итого акад.часов" or self.planList[col + "2"].value == "Итого акад.часов" or self.planList[col + "3"].value == "Итого акад.часов":
                hoursColumn = col
            if zeColumn != "Неопределено" and hoursColumn != "Неопределено" and nameColumn != "Неопределено":
                break
        else:
            raise Exception("Не удалось распарсить колонки")   
        for i in range(len(practiceView)):
            startIndex = 1    
            while self.planList[nameColumn + str(startIndex)].value != practiceTyp[i]:
                startIndex += 1
                if startIndex > self.planList.max_row:
                    break
            laborTime = str(self.planList[hoursColumn + str(startIndex)].value) + "/" + str(self.planList[zeColumn + str(startIndex)].value) 
            compList = self.book['Компетенции']
            currentCode = ""
            currentText = ""
            allComp = []
            for row in range(2, compList.max_row + 1):
                code = compList["B" + str(row)].value
                comp = compList["D" + str(row)].value
                if code:
                   currentCode = code
                   currentText = comp
                if comp == practiceTyp[i]:
                    allComp.append(currentCode + " - " + currentText)  
            practiceList.append([laborTime, allComp])    
        return practiceList        
      
