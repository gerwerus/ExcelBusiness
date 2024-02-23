from excelProcessing import ExcelProccessing, ExcelTread
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