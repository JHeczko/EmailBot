from PySide6.QtCore import QTranslator, QLocale, QLibraryInfo
from PySide6.QtWidgets import QApplication

from Interface import MainWindow

if __name__ == '__main__':
    app = QApplication([])
    translator = QTranslator()
    if translator.load(QLocale("pl_PL"), "qtbase", "_", QLibraryInfo.location(QLibraryInfo.TranslationsPath)):
        app.installTranslator(translator)
    window = MainWindow()
    window.show()
    app.exec()