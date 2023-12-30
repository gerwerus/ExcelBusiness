import sys
from style import style
from PySide6.QtWidgets import (QLineEdit, QPushButton, QApplication,
    QVBoxLayout, QDialog, QMainWindow, QDialogButtonBox, QMessageBox)
from mainWindow import MainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(style)
    window = MainWindow(app)
    window.show()
    sys.exit(app.exec())