import os.path

import openpyxl
from openpyxl import load_workbook
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QPixmap, QPalette, QAction
from PySide6.QtWidgets import QWidget, QApplication, QMainWindow, QFileDialog, QToolBar, QMessageBox


class Window(QMainWindow):
    def __init__(self):
        # init of local variables
        super().__init__()
        self.workbook : openpyxl.Workbook = None

        # setting up window sizes and icons
        self.setWindowTitle('Przetwarzanie excela')
        self.setMinimumSize(800, 600)
        bitmap = QPixmap(os.path.join(os.getcwd() + "./public/logo.ico"))
        icon = QIcon(bitmap)
        self.setWindowIcon(icon)

        # setting up buttons for menu
        button_open = QAction('Otwórz plik', self)
        button_open.triggered.connect(self.file_open)
        button_save = QAction("Zapisz plik", self)
        button_save.triggered.connect(self.file_save)

        # adding toolbar for the app
        toolbar = QToolBar(self)
        toolbar.addAction(button_open)
        toolbar.addSeparator()
        toolbar.addAction(button_save)
        self.addToolBar(toolbar)

        # adding menu for the app
        menu = self.menuBar()
        file_menu = menu.addMenu("&Plik")
        file_menu.addAction(button_open)
        file_menu.addAction(button_save)

    def file_open(self):
        path = QFileDialog().getOpenFileName(QWidget(self), 'Open file', os.getcwd(), "Excel Files (*.xlsx)")[0]
        if path != '' and path is not None:
            try:
                self.workbook = load_workbook(path[0])
                dlg = QMessageBox(self)
                dlg.setWindowTitle("Wszystko ok")
                dlg.setText(f"Załadowano plik excela")
                dlg.exec()
            except Exception as e:
                dlg = QMessageBox.critical(self, "Błąd ładowania", "Nie udało się załadować pliku")
        else:
            dlg = QMessageBox.critical(self, "Błąd ładowania", "Nie wybrano pliku")


    def file_save(self):
        if self.workbook is None:
            dlg = QMessageBox.critical(self, "Nie wczytano notatnika", "Nie wczytano notatnika do zapisu")
            return

        path = QFileDialog().getSaveFileName(QWidget(self), 'Save file', os.getcwd(), "Excel Files (*.xlsx)")
        print(path)
        if path != '' and path is not None:
            print(path)
            self.workbook.save(path)
        else:
            dlg = QMessageBox.critical(self,"Nie wybrano ścieżki do zapisu", "Nieprawnie wybrana ścieżka zapisu")
if __name__ == '__main__':
    app = QApplication([])
    window = Window()
    window.show()

    app.exec()