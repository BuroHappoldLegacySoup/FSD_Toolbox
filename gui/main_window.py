from PyQt5.QtWidgets import QMainWindow, QWidget, QHBoxLayout, QListWidget, QStackedWidget

from gui.project_page import ProjectPage
from version import __version__

class MainWindow(QMainWindow):
    """
    MainWindow class that sets up the GUI.
    """

    def __init__(self):
        """
        Initialize the MainWindow.
        """
        super().__init__()

        self.setWindowTitle("FSRG Tool_v{}".format(__version__))

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QHBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.sidebar = QListWidget()
        self.sidebar.addItems(["Project"])
        #self.sidebar.currentRowChanged.connect(self.display_content)
        self.sidebar.setFixedWidth(150)
        self.sidebar.setStyleSheet("background-color: rgb(196, 214, 0); color: white; font-weight: bold;")

        self.stack = QStackedWidget()

        self.page1 = ProjectPage()
        self.page1.setMinimumWidth(900)
        """self.page1.next_clicked.connect(self.go_to_design_page)

        self.page2 = None
        self.page3 = None
        self.page4 = None"""

        self.stack.addWidget(self.page1)

        self.layout.addWidget(self.sidebar)
        self.layout.addWidget(self.stack)

        self.setStyleSheet("background-color: white;")
