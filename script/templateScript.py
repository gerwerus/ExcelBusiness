from docxtpl import DocxTemplate
import os

def generate(data, template="script/template.docx", destination="./bin"):
    doc = DocxTemplate(template)
    context = {key: data[key] for key in data}
    doc.render(context)
    doc.save(os.path.join(destination, data["Направление"] + "_" + data["Тип"] + ".docx"))