from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import QDate

from gui.main_window import MainWindow
from gui.initial_window import InitialWindow

import sys

def main():
    """
    Main function to run the application.
    """

    app = QApplication([])
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # Set the application icon
    app.setWindowIcon(QIcon('fsrg.ico'))
        
    initial_window = InitialWindow()
    if initial_window.exec_() == QDialog.Accepted:
        window = MainWindow()
        window.show()
        app.exec_()

if __name__ == "__main__":
    main()

