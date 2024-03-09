from components import table, mainMenu
from PySide6.QtCore import QDate
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


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__()
        self.setWindowTitle("Excel Plan")
        self.resize(600, 400)
        # self.setFixedSize(600, 400)
        self.files = []
        layout = QGridLayout()
        # Добавление Вкладок
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.TabPosition.North)

        self.table = table.TableComponent(
            [
                "Институт",
                "Направление",
                "Профиль",
                "Семестр",
                "Вид",
                "Тип",
                "Трудоёмкость",
                "Начало",
                "Окончание",
                "Компетенции",
                "Год выгрузки",
            ]
        )
        self.mainMenu = mainMenu.mainMenuComponent(
            self.callback, addLine=self.table.addLine
        )

        self.tabs.addTab(self.mainMenu, "Выбор файлов")
        self.tabs.addTab(self.table, "Данные")
        layout.addWidget(self.tabs)
        #
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def callback(self, files):
        self.files = files
