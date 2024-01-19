from PySide6.QtCore import QThread
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
        for index, obj in enumerate(self.objects):
            print(obj.getInstituteName(), obj.getCode(), obj.getProfile())
            self.mainWindow.progressBar.setValue(int(((index + 1)/self.ln) * 100)) 
class ExcelProccessing():
    def __init__(self, filename):
        self.filename = filename
        self.book = load_workbook(filename)
        self.titleList = self.book['Титул']
    def getInstituteName(self):
        return self.titleList['D38'].value # С заглавной??? 
    def getCode(self):
        return self.titleList['D27'].value
    def getProfile(self):
        return self.titleList['D30'].value
        
      
