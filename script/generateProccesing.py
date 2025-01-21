from PySide6.QtCore import QThread
from docxtpl import DocxTemplate
import re
import os
from pathlib import Path


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
        try:
            for index, obj in enumerate(self.objects):
                obj.generate()
                self.parent.changeProgressSignal.emit(
                    int(((index + 1) / self.ln) * 100)
                )
        finally:
            self.parent.generateButton.setEnabled(True)


class WordProccesing:
    def __init__(self, data, template="script/template.docx", destination="./bin"):
        self.data = data
        self.template = template
        self.templateFilename = os.path.split(template)[1]
        self.destination = destination

    def setDestination(self, destination):
        self.destination = destination

    def _makeDir(self):
        path0 = Path(self.templateFilename).stem
        path1 = f"{self.data['Направление'].split(' ')[0]}_{self.data['Год выгрузки']}"
        path2 = re.sub(r"[^\w\s]+", "", self.data["Профиль"])
        path3 = str((int(self.data["Семестр"]) + 1) // 2)

        finalPath = Path(self.destination) / path0 / path1 / path2 / path3
        
        finalPath.mkdir(parents=True, exist_ok=True)

        return finalPath

    def generate(self):
        doc = DocxTemplate(self.template)
        context = {key: self.data[key] for key in self.data}
        if context['Год выгрузки']:
            context['ГодВыгрузки'] = context.pop('Год выгрузки')
        doc.render(context)
        course = str((int(self.data["Семестр"]) + 1) // 2)
        profile = re.sub(r"[^\w\s]+", "", self.data["Профиль"])
        doc.save(
            os.path.join(
                self._makeDir(),
                f"{self.data['Направление'].split(' ')[0]}_{profile}_{course}_{self.data['Вид']}.{self.templateFilename.split('.')[1]}",
            )
        )
