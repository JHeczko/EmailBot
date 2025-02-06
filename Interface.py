import os.path

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QPixmap, QPalette, QAction
from PySide6.QtWidgets import QWidget, QApplication, QMainWindow, QFileDialog, QToolBar, QStatusBar, QToolButton


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Przetwarzanie excela')
        self.setMinimumSize(800, 600)
        # palete = self.palette()
        # palete.setColor(QPalette.Window, "#999999")
        # self.setPalette(palete)

        bitmap = QPixmap(os.path.join(os.getcwd() + "logo.ico"))
        icon = QIcon(bitmap)
        self.setWindowIcon(icon)

        toolbar = QToolBar(self)

        button_open = QAction('Otwórz plik', self)
        button_open.triggered.connect(self.file_open())

        button_save = QAction("Zapisz plik", self)
        button_open.triggered.connect(self.file_save())

        toolbar.addAction(button_open)
        toolbar.addSeparator()
        toolbar.addAction(button_save)
        self.addToolBar(toolbar)

        menu = self.menuBar()
        file_menu = menu.addMenu("&Plik")
        file_menu.addAction(button_open)
        file_menu.addAction(button_save)

    def file_open(self):
        print('s')

    def file_save(self):
        print('s')

if __name__ == '__main__':
    app = QApplication([])
    window = Window()
    window.show()

    app.exec()