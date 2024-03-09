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
        # Создание иерархии
        for index, obj in enumerate(self.objects):
            obj.generate(index)
            self.parent.changeProgressSignal.emit(int(((index + 1) / self.ln) * 100))
        self.parent.generateButton.setEnabled(True)


class WordProccesing:
    def __init__(self, data, template="script/template.docx", destination="./bin"):
        self.data = data
        self.template = template
        self.templateFilename = os.path.split(template)[1]
        self.destination = destination

    def setDestination(self, destination):
        self.destination = destination

    def generate(self, index):
        doc = DocxTemplate(self.template)
        context = {key: self.data[key] for key in self.data}
        doc.render(context)
        course = (int(self.data["Семестр"]) + 1) // 2
        doc.save(
            os.path.join(
                self.destination,
                str(index) + self.templateFilename,
            )
        )
