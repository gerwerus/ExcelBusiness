from PySide6.QtWidgets import (
    QLineEdit, QPushButton, QApplication, QTextEdit,
    QVBoxLayout, QDialog, QGroupBox, QMainWindow, QRadioButton,
    QFileDialog, QWidget, QDateEdit, QLabel, QGridLayout,
    QDialogButtonBox, QMessageBox, QTabWidget, QScrollArea
    )
class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__()
        self.setWindowTitle("Excel Plan")
        self.setFixedSize(600, 400)
        layout = QGridLayout()
        self.getFilesButton = QPushButton("Загрузить файлы")
        self.getFilesButton.clicked.connect(self.getFiles)
        layout.addWidget(self.getFilesButton, 0, 0)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
    def getFiles(self):
        files = QFileDialog.getOpenFileNames(self, filter="Excel files (*.xls, *xlsx);; All files (*)")[0]  
        print(files)