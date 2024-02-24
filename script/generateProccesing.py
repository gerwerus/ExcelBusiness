from PySide6.QtCore import QThread
from docxtpl import DocxTemplate
import os


class WordThread(QThread):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.objects = 0
        self.ln = 0
    def setObjects(self, objects, ln=0):
        self.objects = objects
        self.ln = ln

    def run(self):
        for index, obj in enumerate(self.objects):
            obj.generate()
            self.mainMenu.progressBar.setValue(int(((index + 1)/self.ln) * 100))
            


class WordProccesing():
    def __init__(self, data, template="script/template.docx", destination="./bin"):
        self.data = data
        self.template = template
        self.destination = destination
    def setDestination(self, destination):
        self.destination = destination
    def generate(self):
        doc = DocxTemplate(self.template)
        context = {key: self.data[key] for key in self.data}
        doc.render(context)
        doc.save(os.path.join(self.destination, self.data["Направление"] + "_" + self.data["Тип"] + ".docx"))