from script.generateProccesing import WordThread, WordProccesing
from components.mixins import ProgressChangeMixin

import openpyxl
from PySide6.QtWidgets import (
    QLineEdit,
    QPushButton,
    QApplication,
    QTextEdit,
    QDateEdit,
    QDialog,
    QGroupBox,
    QMainWindow,
    QRadioButton,
    QFileDialog,
    QWidget,
    QDateEdit,
    QLabel,
    QGridLayout,
    QDialogButtonBox,
    QInputDialog,
    QProgressBar,
    QTableWidgetItem,
    QTableWidget,
    QMenu,
)


class TableComponent(QWidget, ProgressChangeMixin):
    def __init__(self, header: list[str]):
        super(TableComponent, self).__init__()
        layout = QGridLayout()
        self.table = QTableWidget()
        self.header = header
        self.table.setColumnCount(len(header))
        self.table.setHorizontalHeaderLabels(header)

        self.context_menu = QMenu(self)
        action_fill = self.context_menu.addAction("Заполнить")
        action_fill.triggered.connect(self.fillSelectedCells)

        self.clearButton = QPushButton("Очистить")
        self.clearButton.clicked.connect(self.clearTable)
        self.excelButton = QPushButton("Выгрузить в Excel")
        self.excelButton.clicked.connect(self.saveToExcel)
        self.generateButton = QPushButton("Сгенерировать")
        self.generateButton.clicked.connect(self.generateWord)

        self.progressBar = QProgressBar()
        self.progressBar.setValue(0)
        self.changeProgressSignal[float].connect(self.changeProgressSlot)

        layout.addWidget(self.table, 0, 0, 1, 4)
        layout.addWidget(self.progressBar, 1, 0, 1, 1)
        layout.addWidget(self.generateButton, 1, 1, 1, 1)
        layout.addWidget(self.excelButton, 1, 2, 1, 1)
        layout.addWidget(self.clearButton, 1, 3, 1, 1)

        self.setLayout(layout)

        self.wordThread_instance = WordThread(self)

    def contextMenuEvent(self, event):
        self.context_menu.exec(event.globalPos())

    def addLine(self, line: list[str]):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for index, el in enumerate(line):
            val = str(el).rstrip().lstrip().replace('"', "'")
            self.table.setItem(row, index, QTableWidgetItem(val))

    def clearTable(self):
        while self.table.rowCount() > 0:
            self.table.removeRow(0)

    def saveToExcel(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "All Files(*.xlsx)")
        if not filename:
            return

        data = [self.header] + self._get_all_data()

        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        for r, row in enumerate(data, start=1):
            for c, col in enumerate(row, start=1):
                worksheet.cell(row=r, column=c).value = col
        workbook.save(filename=filename)

    def _get_all_data(self) -> list:
        _list = []
        model = self.table.model()
        for row in range(model.rowCount()):
            _r = []
            for column in range(model.columnCount()):
                _r.append("{}".format(model.index(row, column).data() or ""))
            _list.append(_r)
        return _list

    def generateWord(self):
        templates = QFileDialog.getOpenFileNames(
            self, filter="Шаблоны (*.docx);; All files (*)", caption="Выберите шаблоны"
        )[0]
        if not (templates):
            return
        destination = QFileDialog.getExistingDirectory(self, "Папка назначения")
        if not (destination):
            return
        ln = len(templates) * self.table.rowCount()

        def getObjects():
            for template in templates:
                for row in range(self.table.rowCount()):
                    data = {}
                    for ind, el in enumerate(self.header):
                        data.update({el: self.table.item(row, ind).text()})
                    yield WordProccesing(data, template, destination)

        objects = getObjects()
        self.wordThread_instance.setObjects(objects, ln)
        self.progressBar.setValue(0)
        self.wordThread_instance.start()
        self.generateButton.setDisabled(self.wordThread_instance.isRunning())

    def fillSelectedCells(self):
        value, ok = QInputDialog.getText(self, "Заполнить ячейки", "Значение")
        if ok:
            for index in self.table.selectedIndexes():
                item = self.table.itemFromIndex(index)
                if item:
                    item.setText(value)
