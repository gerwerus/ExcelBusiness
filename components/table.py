from excelProcessing import ExcelProccessing, ExcelTread
from PySide6.QtCore import QDate
from PySide6.QtWidgets import (
    QLineEdit, QPushButton, QApplication, QTextEdit,
    QDateEdit, QDialog, QGroupBox, QMainWindow, QRadioButton,
    QFileDialog, QWidget, QDateEdit, QLabel, QGridLayout,
    QDialogButtonBox, QMessageBox, QTabWidget, QTableWidgetItem, QTableWidget
    )
class TableComponent(QWidget):
    def __init__(self, header: list[str]):
        super(TableComponent, self).__init__()
        layout = QGridLayout()
        self.table = QTableWidget()
        self.header = header
        self.table.setColumnCount(len(header))
        self.table.setHorizontalHeaderLabels(header)
        
        self.clearButton = QPushButton("Очистить")
        self.clearButton.clicked.connect(self.clearTable)
        self.excelButton = QPushButton("Выгрузить в Excel")
        self.clearButton.clicked.connect(self.generateWord)
        self.generateButton = QPushButton("Сгенерировать")
        # self.clearButton.clicked.connect(self.clearTable)

        layout.addWidget(self.table, 0, 0, 1, 4)
        layout.addWidget(self.generateButton, 1, 1, 1, 1)
        layout.addWidget(self.excelButton, 1, 2, 1, 1)
        layout.addWidget(self.clearButton, 1, 3, 1, 1)

        self.setLayout(layout)
    def addLine(self, line: list[str]):
        # if len(line) > len(self.header):
        #     raise Exception("Невозможно вставить в таблицу")
        row = self.table.rowCount()
        self.table.insertRow(row)
        for index, el in enumerate(line):
            self.table.setItem(row, index, QTableWidgetItem(str(el)))
    def clearTable(self):
        while (self.table.rowCount() > 0):
            self.table.removeRow(0)
    def excelUnload(self):
        pass
    def generateWord(self):
        pass
