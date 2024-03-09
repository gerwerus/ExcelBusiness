from excelProcessing import ExcelProccessing, ExcelTread
from components.mixins import ProgressChangeMixin

from PySide6.QtCore import QDate, Slot, Signal
import os
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
    QMessageBox,
    QTabWidget,
    QScrollArea,
    QProgressBar,
)


class mainMenuComponent(QWidget, ProgressChangeMixin):
    def __init__(self, callback, *args, **kwargs):
        super(mainMenuComponent, self).__init__()
        self.excelTread_instance = ExcelTread(self)
        for k in kwargs:
            if k == "addLine":
                self.addLine = kwargs["addLine"]
        self.callback = callback
        self.files = []
        layout = QGridLayout()
        self.getFilesButton = QPushButton("Выбрать файлы")
        self.getFilesFromDirButton = QPushButton("Выбрать Папку")
        self.dateEdit = QDateEdit()
        self.dateEdit.setDate(QDate.currentDate())
        # self.dateEdit.date()
        self.progressBar = QProgressBar()
        self.progressBar.setValue(0)
        self.changeProgressSignal[float].connect(self.changeProgressSlot)

        self.getFilesButton.clicked.connect(self.getFiles)
        self.getFilesFromDirButton.clicked.connect(self.getFilesFromDirectory)

        layout.addWidget(QLabel("Дата формирования:"), 0, 0)
        layout.addWidget(self.dateEdit, 0, 1)
        layout.addWidget(self.getFilesButton, 1, 0, 1, 1)
        layout.addWidget(self.getFilesFromDirButton, 1, 1, 1, 1)
        layout.addWidget(self.progressBar, 2, 0, 1, 2)

        self.setLayout(layout)

    def getFiles(self):
        self.files = QFileDialog.getOpenFileNames(
            self, filter="Excel files (*.xls, *xlsx);; All files (*)"
        )[0]
        self.startTread()

    def getFilesFromDirectory(self):
        for top, __, files in os.walk(
            str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        ):
            for name in files:
                if name.endswith(".xls") or name.endswith(".xlsx"):
                    self.files.append(os.path.join(top, name))
        # print(self.files, self.files.__len__())
        self.files = list(set(self.files))
        self.startTread()

    def startTread(self):
        ln = len(self.files)
        if not (ln):
            return
        objects = (ExcelProccessing(i) for i in self.files)
        self.excelTread_instance.setObjects(objects, ln)
        self.progressBar.setValue(0)
        self.excelTread_instance.start()
        running = self.excelTread_instance.isRunning()
        self.getFilesButton.setDisabled(running)
        self.getFilesFromDirButton.setDisabled(running)
