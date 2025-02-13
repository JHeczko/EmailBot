import ctypes
import os.path

import openpyxl
from openpyxl import load_workbook
from PySide6.QtCore import Qt, QSysInfo, QTranslator, QLocale, QLibraryInfo
from PySide6.QtGui import QIcon, QPixmap, QAction
from PySide6.QtWidgets import QWidget, QApplication, QMainWindow, QFileDialog, QMessageBox, QStackedLayout, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton, QToolButton
from Parsing import edit_excel

class HelpWindow(QWidget):
    def __init__(self):
        super().__init__()

        # setting up for new window
        self.setWindowTitle("Pomoc")
        self.setMinimumSize(500,300)
        self.setMaximumSize(700,500)
        bitmap = QPixmap(os.path.join(os.getcwd() + "./public/logo.ico"))
        icon = QIcon(bitmap)
        self.setWindowIcon(icon)

        # main layout
        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # button layout and text stack layout
        self.main_stack = QStackedLayout()
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(10)
        buttons_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # configuring help text
        # =========== 1 HELP ABOUT OPENING DATA AND FORMAT OF DATA ===========
        help_window1 = QWidget()
        help_layout1 = QVBoxLayout()
        help_layout1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_layout1.setSpacing(30)
        help_window1.setLayout(help_layout1)

        help_imie_h1 = QLabel("<h1>Format Imienia i Nazwiska</h1>")
        help_imie_h1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_imie_tresc = QLabel('Imie i nazwisko mamy w bazowej tabeli powinno być sformatowane w taki sposób, aby stanowiło dwa człony, czyli mamy imię na przykład "Katarzyna" i po tym imieniu, nazwisko np "Baranowska-Kowalska" (może być na odwrót, o tym później). Jeśli w tabeli znajdą się dwie Mamy o różnym imieniu lub nazwisku, to wtedy dzieci tych dwóch osób będą traktowane dla każdej z tych osób osobno. W kolumnie <b>EMAIL</b> powinny być maile oddzielone ";", jeśli jest jeden mail to wystaczy wpisać tylko tego maila bez średniak(";"). Liczby zadłużeń muszą być liczbami.')
        help_imie_tresc.setWordWrap(True)

        help_layout1.addWidget(help_imie_h1)
        help_layout1.addWidget(help_imie_tresc)
        self.main_stack.addWidget(help_window1)

        # =========== 2 HELP ABOUT PROCESSING DARA ===========
        help_window2 = QWidget()
        help_layout2 = QVBoxLayout()
        help_layout2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_layout2.setSpacing(30)
        help_window2.setLayout(help_layout2)

        help_przetwarzanie_h1 = QLabel("<h1>Przetwarzanie</h1>")
        help_przetwarzanie_h1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_przetwarzanie_tresc = QLabel('Aby zacząć przetwarzać należy najpierw wczytać plik, robi się to guzikiem <b>"Otwórz Plik"</b> znajduję się on w menu <b>Plik</b>. Następnie po otwarciu pliku w oknie wyświetlą się listy wraz z kolumnami. Należy wybrać dla każdej listy odpowiednią kolumnę. Na przykład dla listy z dopiskiem "Imie i Nazwisko Mamy", należy wybrać, kolumne tytuł kolumny, która odpowiada, kolumnie z odpowiednimi danymi. Można również wybrać odpowiednią opcję sformatowanych danych, czy w kolumnie mamy dane w formacie <b>IMIE-NAZWISKO</b>, czy <b>NAZWISKO-IMIE</b>, chodzi o to w jakiej kolejności ułożone są imiona, na przykład czy w tabeli z imionami mamy <b>"Katarzyna Nowak"</b> czy <b>"Nowak Katarzyna"</b>')
        help_przetwarzanie_tresc.setWordWrap(True)

        help_layout2.addWidget(help_przetwarzanie_h1)
        help_layout2.addWidget(help_przetwarzanie_tresc)
        self.main_stack.addWidget(help_window2)

        # =========== 3 HELP ABOUT SAVING FILE ===========
        help_window3 = QWidget()
        help_layout3 = QVBoxLayout()
        help_layout3.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_layout3.setSpacing(30)
        help_window3.setLayout(help_layout3)

        help_zapis_h1 = QLabel("<h1>Zapis</h1>")
        help_zapis_h1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_zapis_tresc = QLabel('Po przetworzeniu pliku wyświetli nam się strona, która powie, że wszystko poszło ok, to nie oznacza, że dane zostały poprawnie przetworzone, oznacza to, że program nie napotkał żadnego błędu na drodze, może się zdarzyć, że kolumny zostały źle zaznaczone, a więc dane również źle przetworzone i wynik może być bez sensu, należy więc zweryfikować końcowy plik. Aby ten końcowy plik dostać należy kliknąć, w menu <b>"Plik"</b> przycisk <b>"Zapisz Plik"</b>')
        help_zapis_tresc.setWordWrap(True)

        help_layout3.addWidget(help_zapis_h1)
        help_layout3.addWidget(help_zapis_tresc)
        self.main_stack.addWidget(help_window3)


        self.main_stack.setCurrentIndex(0)

        # configuring buttons
        next_button = QToolButton()
        next_button.setArrowType(Qt.RightArrow)
        next_button.setToolButtonStyle(Qt.ToolButtonIconOnly)
        next_button.pressed.connect(self.next)

        previous_button = QToolButton()
        previous_button.setArrowType(Qt.LeftArrow)
        previous_button.setToolButtonStyle(Qt.ToolButtonIconOnly)
        previous_button.pressed.connect(self.previous)

        buttons_layout.addWidget(previous_button)
        buttons_layout.addWidget(next_button)

        # configuring page counter
        self.strona = QLabel(f"Strona {self.main_stack.currentIndex()+1}/{self.main_stack.count()}")
        self.strona.setAlignment(Qt.AlignmentFlag.AlignRight)

        # adding everything together
        main_layout.addLayout(self.main_stack)
        main_layout.addLayout(buttons_layout)
        main_layout.addWidget(self.strona)

        self.setLayout(main_layout)

    def next(self):
        ind = self.main_stack.currentIndex()
        if ind + 1 < self.main_stack.count():
            self.main_stack.setCurrentIndex(ind + 1)
            self.strona.setText(f"Strona {self.main_stack.currentIndex()+1}/{self.main_stack.count()}")

    def previous(self):
        ind = self.main_stack.currentIndex()
        if ind -1 >= 0:
            self.main_stack.setCurrentIndex(ind -1)
            self.strona.setText(f"Strona {self.main_stack.currentIndex()+1}/{self.main_stack.count()}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # initailizing all local variables needed for workflow of program
        self.workbook : openpyxl.Workbook = None
        self.workbook_edited : openpyxl.Workbook= None
        self.labels = []
        self.info_labels = ["Imie i Nazwisko Mamy", "Mail", "3-30 dni", "31-60 dni", "61-365 dni"]
        self.comboboxes = []
        self.file_name = QLabel('')
        self.window_help = None
        self.theme = 0 # does not matter what value here it will be set later

        # |=|=|=|=|=|=|=|=| CUSTOMIZATION |=|=|=|=|=|=|=|=|

        # setting up window sizes and icons
        self.setWindowTitle('Przetwarzanie excela')
        self.setMinimumSize(800, 600)
        bitmap = QPixmap(os.path.join(os.getcwd() + "./public/logo.ico"))
        icon = QIcon(bitmap)
        self.setWindowIcon(icon)
        self.file_name.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # setting up buttons for menu
        button_open = QAction('Otwórz plik', self)
        button_open.triggered.connect(self.file_open)
        button_save = QAction("Zapisz plik", self)
        button_save.triggered.connect(self.file_save)
        button_help = QAction("Pomoc", self)
        button_help.triggered.connect(self.help_popup)
        button_mode = QAction("Motyw ciemny/jasny",self)
        button_mode.setCheckable(True)
        button_mode.setChecked(True)
        button_mode.triggered.connect(self.switch_modes)

        # =-=-=-=-=--= PLATFORM CUSTOMIZATION =-=-=-=-=--=
        if QSysInfo.productType() == 'macos':
            menu = self.menuBar()
            file_menu = menu.addMenu("&Plik")
            help_menu = menu.addMenu("&Pomoc")
            help_menu.addAction(button_help)
            file_menu.addAction(button_open)
            file_menu.addAction(button_save)
        else:
            menu = self.menuBar()
            file_menu = menu.addMenu("&Plik")
            menu.addAction(button_help)
            menu.addAction(button_mode)
            file_menu.addAction(button_open)
            file_menu.addAction(button_save)
            # this line is only for windows QSS working nicely for windows only, mac os has nativly nice themes for apps
            self.switch_modes(True)

        # |=|=|=|=|=|=|=|=| SCENE CREATION |=|=|=|=|=|=|=|=|

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
        window1_layout.setSpacing(30)
        self.window1.setLayout(window1_layout)

        enter = QLabel('<h1>Witamy w przetwarzaniu excela</h1>')
        enter.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        opis = QLabel('Jeśli potrzebujesz pomocy kliknij na górze guzik o nazwie <b>"Pomoc"</b>, jeśli chcesz załadować plik i zacząć go przetwarzać kliknij <b>"Otwórz Plik"</b>. Po otwarciu i przetworzeniu pliku zapisz go guzikiem <b>"Zapisz Plik"</b>')
        opis.setStyleSheet("""
            font-size: 20px;
            text-align: justify;
        """)
        opis.setWordWrap(True)

        window1_layout.addWidget(enter)
        window1_layout.addWidget(opis)

        # =-==-=-=-=-=-= SECOND SCENE AKA LOADED FILE=-==-=-=-=-=-=
        # TODO need a selectio for mode processing and then we are done
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

        # =-==-=-=-=-=-= THIRD SCENE AKA ENDING/SAVING FILE=-==-=-=-=-=-=
        self.window3 = QWidget()
        window3_layout = QVBoxLayout()
        window3_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        window3_layout.setSpacing(20)
        self.window3.setLayout(window3_layout)

        w3_text = QLabel("<h1>Wszystko poszło ok</h1>")
        w3_text.setAlignment(Qt.AlignmentFlag.AlignCenter)

        w3_text2 = QLabel('"Plik teraz należy zapisać trzeba przejść do menu <b>"Plik" -> "Zapisz plik"</b>, następnie wybieramy lokalizację zapisu i nazwę pliku')
        w3_text2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        w3_text2.setStyleSheet("""
              font-size: 20px;
              text-align: justify;
          """)

        window3_layout.addWidget(w3_text)
        window3_layout.addWidget(w3_text2)

        window3_button_layout = QHBoxLayout()
        window3_button_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        w3_button = QPushButton("Wróć")
        w3_button.pressed.connect(self.back_button)
        window3_button_layout.addWidget(w3_button)
        window3_layout.addLayout(window3_button_layout)

        # =-==-=-=-=-=-= ADDING EVERYTHING TOGETHER =-==-=-=-=-=-=
        self.main_stack.addWidget(self.window1)
        self.main_stack.addWidget(self.window2)
        self.main_stack.addWidget(self.window3)

        self.main_stack.setCurrentWidget(self.window1)

    def file_open(self):
        '''
        Opening a file from a disk that user select, also if there is already opened and preprocessed file it will clean it up from prorgam memory
        :return:
        '''
        # mechanism for checking if there is already a processed file in memory if so then aks user if he want to delete it
        if self.workbook_edited is not None:
            result = QMessageBox.question(self, "Nie zapisano",
                                          "Masz niezapisany przetworzony arkusz, czy napewno chcesz otworzyć nowy arkusz i porzucić stary niezapisując go?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if result == QMessageBox.No:
                return
            else:
                try:
                    self.workbook_edited.close()
                    self.workbook_edited = None
                    self.main_stack.setCurrentWidget(self.window1)
                except Exception as e: pass

        # standard loading steps from here
        path = QFileDialog().getOpenFileName(QWidget(self), 'Open file', os.getcwd(), "Excel Files (*.xlsx)")[0]
        if self.workbook is not None:
            try:
                self.workbook.close()
            except Exception: pass
        if path == '':
            return
        if path != '' and path is not None:
            try:
                self.workbook = load_workbook(path)
                self.labels = [x.value for x in self.workbook.active[1]]
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
        '''
        Saving file and clearing memory from everything, basicly doing back button, but without asking of anything
        :return:
        '''
        if self.workbook_edited is None:
            QMessageBox.critical(self, "Nie przetworzono arkusza", "Arkusz jest nieprzetworzony, nie mozna zapisac")
            return

        path = QFileDialog().getSaveFileName(QWidget(self), 'Save file', os.getcwd(), "Excel Files (*.xlsx)")[0]
        if path != '' and path is not None:
            try:
                self.workbook_edited.save(path)
                self.main_stack.setCurrentWidget(self.window1)
                try:
                    self.workbook.close()
                except Exception as e:
                    pass

                try:
                    self.workbook_edited.close()
                except Exception as e:
                    pass
                self.workbook = None
                self.workbook_edited = None
                self.labels = []
                QMessageBox.information(self,"Wszystko ok", f"Zapisano poprawnie plik pod nazwą {os.path.basename(path)}")
            except Exception as e:
                self.back_button()
                QMessageBox.critical(self, "Błąd zapisu", "Coś nie tak z zapisywaniem pliku")
        else:
            return

    def back_button(self):
        # it is a checker if there is already some proccessed file in the memory of program
        if self.workbook_edited is not None:
            result = QMessageBox.question(self, "Nie zapisano",
                                          "Masz niezapisany przetworzony arkusz, czy napewno chcesz porzucić zrobione do teraz zmiany", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if result == QMessageBox.No:
                return

        # cleanup steps
        self.main_stack.setCurrentWidget(self.window1)
        try:
            self.workbook.close()
        except Exception as e: pass
        try:
            self.workbook_edited.close()
        except Exception as e: pass
        self.workbook = None
        self.workbook_edited = None
        self.labels = []

    def next_button(self):
        try:
            indexes = []
            for box in self.comboboxes:
                indexes.append(box.currentIndex())
            self.workbook_edited = edit_excel(self.workbook, *indexes)
            self.main_stack.setCurrentWidget(self.window3)
        except Exception as e:
            self.back_button()
            QMessageBox.critical(self, "Błąd przetwarzania", "Coś nie tak z przetwarzaniem. Upewnij się, że kolumny są poprawnie wybrane oraz dane są poprawnie przygotowane")

    def help_popup(self):
        def cleanup():
            self.window_help = None

        if self.window_help is None:
            self.window_help = HelpWindow()
            self.window_help.closeEvent = lambda event: (cleanup(),event.accept())
            self.switch_modes(self.theme)
            self.window_help.show()
        else:
            self.window_help.raise_()

    def switch_modes(self,dark_mode):
        dark_mode_css = """
            QMainWindow {
                background-color: #242424;
            }
            QWidget {
                background-color: #242424;
                color: #FFFFFF;
            }
            
            QPushButton{
                background-color: #212121;
                color: #FFFFFF;
                padding: 3px;
                border: 2px solid #404040;
                float:left;
            }
            QPushButton:hover{
                background-color: #aba7ab;
                color: #212121;
                padding: 3px;
                border: 2px solid #919191;
                float:left;
            }
            
            QComboBox {
                background-color: #333333;
                color: white;
                border: 1px solid #555;
                padding: 5px;
            }
            QComboBox::drop-down {
                border: none;
                background: #444;
                color: white;
            }
            QComboBox QAbstractItemView {
                background-color: #333333;
                color: white;
                selection-background-color: #555555;
            }
           QComboBox QAbstractItemView::item:hover,
           QComboBox QAbstractItemView::item:selected {
                background-color: gray; /* Change background color on hover */
                color: black; /* Change text color on hover */
            }
    
    
            QMenuBar{
                background-color: #222;
            }
            QMenuBar::item {
                padding: 3px 15px;
                width: 100%;
                float:left
            }
            QMenuBar::item:selected { /* when selected using mouse or keyboard */
                background: #a8a8a8;
                color: black;
            }
            
           QMenu {
               background-color: #333333;
               border: 1px solid #555555;
           }
           QMenu::item {
               padding: 6px 20px;
               color: white;
           }
           QMenu::item:selected {
               background-color: #555555;
           }
        """
        white_mode_css = """
                QMainWindow {
                    background-color: #F5F5F5;
                }
                
                QWidget {
                    background-color: #F5F5F5;
                    color: #212121;
                }
                
                QPushButton {
                    background-color: #E0E0E0;
                    color: #212121;
                    padding: 3px;
                    border: 2px solid #B0B0B0;
                    float: left;
                }
                QPushButton:hover {
                    background-color: #D6D6D6;
                    color: #000000;
                    padding: 3px;
                    border: 2px solid #8E8E8E;
                    float: left;
                }
                
                QComboBox {
                    background-color: #FFFFFF;
                    color: black;
                    border: 1px solid #AAAAAA;
                    padding: 5px;
                }
                QComboBox::drop-down {
                    border: none;
                    background: #DDDDDD;
                    color: black;
                }
                QComboBox QAbstractItemView {
                    background-color: #FFFFFF;
                    color: black;
                    selection-background-color: #DDDDDD;
                }
                QComboBox QAbstractItemView::item:hover,
                QComboBox QAbstractItemView::item:selected {
                    background-color: #C0C0C0;
                    color: black;
                }
                
                QMenuBar {
                    background-color: #E0E0E0;
                }
                QMenuBar::item {
                    padding: 3px 15px;
                    width: 100%;
                    float: left;
                }
                QMenuBar::item:selected { /* when selected using mouse or keyboard */
                    background: #C0C0C0;
                    color: black;
                }
                
                QMenu {
                    background-color: #FFFFFF;
                    border: 1px solid #AAAAAA;
                }
                QMenu::item {
                    padding: 6px 20px;
                    color: black;
                }
                QMenu::item:selected {
                    background-color: #DDDDDD;
                }
        """
        self.theme = dark_mode

        windows_to_switch = [self,self.window_help]
        for window in windows_to_switch:
            if window is None:
                continue

            if dark_mode:
                window.setStyleSheet(dark_mode_css)
            else:
                window.setStyleSheet(white_mode_css)
            if QSysInfo.productType() == 'windows':
                try:
                    hwnd = window.winId()  # Get window handle

                    # Define the attributes for DWM (Desktop Window Manager)
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 20  # Dark mode title bar
                    DWMWA_CAPTION_COLOR = 35  # Title bar color
                    DWMWA_TEXT_COLOR = 36  # Title text color

                    # Convert dark mode flag to ctypes
                    dark_mode_flag = ctypes.c_int(1 if dark_mode else 0)

                    # Set dark/light mode
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(dark_mode_flag), ctypes.sizeof(dark_mode_flag)
                    )

                    # Set custom title bar color (only works when `DWMWA_USE_IMMERSIVE_DARK_MODE` is enabled)
                    title_bar_color = 0x4d4d4d if dark_mode else 0xFFFFFF  # Dark or white
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, DWMWA_CAPTION_COLOR, ctypes.byref(ctypes.c_int(title_bar_color)), 4
                    )

                    # Set title text color (optional)
                    text_color = 0xFFFFFF if dark_mode else 0x000000
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, DWMWA_TEXT_COLOR, ctypes.byref(ctypes.c_int(text_color)), 4)
                except Exception as e:
                    print("Something wrong with windows native color changing for toolbars")

if __name__ == '__main__':
    app = QApplication([])
    translator = QTranslator()
    if translator.load(QLocale("pl_PL"), "qtbase", "_", QLibraryInfo.location(QLibraryInfo.TranslationsPath)):
        app.installTranslator(translator)
    window = MainWindow()
    window.show()
    app.exec()