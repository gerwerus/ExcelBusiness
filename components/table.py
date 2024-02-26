from excelProcessing import ExcelProccessing, ExcelTread
from openpyxl import Workbook
from PySide6.QtCore import QDate
from PySide6.QtWidgets import (
    QLineEdit, QPushButton, QApplication, QTextEdit,
    QDateEdit, QDialog, QGroupBox, QMainWindow, QRadioButton,
    QFileDialog, QWidget, QDateEdit, QLabel, QGridLayout,
    QDialogButtonBox, QMessageBox, QTabWidget, QTableWidgetItem, QTableWidget
    )
class TableComponent(QTableWidget):
    def __init__(self, header: list[str]):
        super(TableComponent, self).__init__()
        layout = QGridLayout()
        self.header = header
        self.setColumnCount(len(header))
        self.setRowCount(0)
        self.setHorizontalHeaderLabels(header)
        
        self.setLayout(layout)
    def addLine(self, line: list[str]):
        # if len(line) != len(self.header):
        #     raise Exception("Невозможно вставить в таблицу")
        row = self.rowCount()
        self.insertRow(row)
        for index, el in enumerate(line):
            self.setItem(row, index, QTableWidgetItem(str(el)))
    
    def clearTable(self):
        self.setRowCount(0)

    def exportToExcel(self, filename: str = "table_data.xlsx"):
        workbook = Workbook()
        sheet = workbook.active

        for columnIndex, header in enumerate(self.header, start=1):
            sheet.cell(row=1, column=columnIndex, value=header)

        for rowIndex, rowData in enumerate(self.get_all_data(), start=2):
            for columnIndex, cellValue in enumerate(rowData, start=1):
                sheet.cell(row=rowIndex, column=columnIndex, value=cellValue)

        workbook.save(filename)

    def get_all_data(self):
        data = []
        for row in range(self.rowCount()):
            rowData = []
            for column in range(self.columnCount()):
                item = self.item(row, column)
                rowData.append(item.text() if item else "")
            data.append(rowData)
        return data