import os.path

import openpyxl
from openpyxl import load_workbook
from PySide6.QtCore import Qt, QSize, QSysInfo
from PySide6.QtGui import QIcon, QPixmap, QPalette, QAction, QGuiApplication
from PySide6.QtWidgets import QWidget, QApplication, QMainWindow, QFileDialog, QToolBar, QMessageBox, QStackedLayout, \
    QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton


class Window(QMainWindow):
    def __init__(self):
        # init of local variables
        super().__init__()
        self.workbook : openpyxl.Workbook = None
        self.labels = []
        self.info_labels = ["Imie i Nazwisko Mamy", "Mail", "3-30 dni", "31-60 dni", "61-365 dni"]
        self.comboboxes = []
        self.file_name = QLabel('')
        self.file_name.setAlignment(Qt.AlignmentFlag.AlignCenter)

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
        button_help = QAction("Pomoc", self)
        button_help.triggered.connect(self.help_popup)

        # adding toolbar for the app only for macusers
        if QSysInfo == 'macos':
            toolbar = QToolBar(self)
            toolbar.addAction(button_open)
            toolbar.addSeparator()
            toolbar.addAction(button_save)
            toolbar.addSeparator()
            toolbar.addAction(button_help)
            self.addToolBar(toolbar)

        # adding menu for the app
        menu = self.menuBar()
        file_menu = menu.addMenu("&Plik")
        menu.addAction(button_help)
        file_menu.addAction(button_open)
        file_menu.addAction(button_save)

        # create the main widget with stack(central widget)
        self.main_widget = QWidget()
        self.main_stack = QStackedLayout()
        self.main_stack.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.main_widget.setLayout(self.main_stack)
        self.setCentralWidget(self.main_widget)

        # =-==-=-=-=-=-= FIRST SCENE AKA NOTHING BOX =-==-=-=-=-=-=
        self.window1= QWidget()
        window1_layout = QVBoxLayout()
        window1_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.window1.setLayout(window1_layout)

        enter = QLabel('Witamy w przetwarzaniu excela')
        enter.setAlignment(Qt.AlignmentFlag.AlignCenter)
        opis = QLabel('Jeśli potrzebujesz pomocy kliknij na górze guzik o nazwie "Pomoc", jeśli chcesz załadować plik i zacząć go przetwarzać kliknij "Otwórz Plik". Po otwarciu i przetworzeniu pliku zapisz go guzikiem ')
        window1_layout.addWidget(enter)
        window1_layout.addWidget(opis)

        # =-==-=-=-=-=-= SECOND SCENE AKA LOADED FILE=-==-=-=-=-=-=
        self.window2 = QWidget()
        window2_layout = QVBoxLayout()
        window2_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # adding text with file name
        window2_layout.addWidget(self.file_name)

        # making sequence of comboboxes
        window2_layout_combobox = QHBoxLayout()
        window2_layout_combobox.setAlignment(Qt.AlignmentFlag.AlignCenter)

        window2_layout.addLayout(window2_layout_combobox)
        self.window2.setLayout(window2_layout)

        for name in self.info_labels:
            temp_lay = QVBoxLayout()
            temp_lay.addWidget(QLabel(name))

            temp_lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
            temp_lay.setContentsMargins(30, 30, 30, 30)
            temp_lay.setSpacing(10)

            combobox_temp = QComboBox()
            combobox_temp.addItems(self.labels)
            temp_lay.addWidget(combobox_temp)
            self.comboboxes.append(combobox_temp)

            window2_layout_combobox.addLayout(temp_lay)

        # buttons for navigation
        window2_layout_button = QHBoxLayout()
        window2_layout_button.setAlignment(Qt.AlignmentFlag.AlignCenter)
        button_next = QPushButton()
        button_cancel = QPushButton()
        button_cancel.setText("Anuluj")
        button_next.setText("Przetwórz")

        button_cancel.pressed.connect(self.back_button)
        button_next.pressed.connect(self.next_button)


        window2_layout_button.addWidget(button_next)
        window2_layout_button.addWidget(button_cancel)
        window2_layout.addLayout(window2_layout_button)

        # =-==-=-=-=-=-= ADDING EVERYTHING TOGETHER =-==-=-=-=-=-=
        self.main_stack.addWidget(self.window1)
        self.main_stack.addWidget(self.window2)

        self.main_stack.setCurrentWidget(self.window1)

    def file_open(self):
        path = QFileDialog().getOpenFileName(QWidget(self), 'Open file', os.getcwd(), "Excel Files (*.xlsx)")[0]
        if path == '':
            return
        if path != '' and path is not None:
            try:
                self.workbook = load_workbook(path)
                QMessageBox.information(self,"Wszystko ok", f"Załadowano plik excela")
                self.labels =  [x.value for x in self.workbook.active[1]]
                self.file_name.setText(f"Praca na pliku: \"{os.path.basename(path)}\"")

                self.main_stack.setCurrentWidget(self.window2)
                for combobox in self.comboboxes:
                    combobox.clear()
                    combobox.addItems(self.labels)
                    combobox.setCurrentIndex(0)
            except Exception as e:
                QMessageBox.critical(self, "Błąd ładowania", f"Nie udało się załadować pliku {e}")
        else:
            QMessageBox.critical(self, "Błąd ładowania", "Nie wybrano pliku")


    def file_save(self):
        if self.workbook is None:
            QMessageBox.critical(self, "Nie wczytano notatnika", "Nie wczytano notatnika do zapisu")
            return

        path = QFileDialog().getSaveFileName(QWidget(self), 'Save file', os.getcwd(), "Excel Files (*.xlsx)")[0]
        if path != '' and path is not None:
            self.workbook.save(path)
        else:
            return

    def help_popup(self): pass

    def back_button(self):
        self.main_stack.setCurrentWidget(self.window1)
        self.workbook.close()
        self.workbook = None
        self.labels = []

    def next_button(self): pass
if __name__ == '__main__':
    app = QApplication([])
    window = Window()
    window.show()
    app.exec()