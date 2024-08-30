from PyQt5.QtWidgets import QDialog, QGroupBox, QPushButton, QGridLayout, QVBoxLayout, QCheckBox

class InitialWindow(QDialog):
    """
    This class represents the initial window that the user sees when the application starts.
    """

    def __init__(self, parent=None):
        """
        Initialize the InitialWindow with a parent widget.
        """
        super().__init__(parent)

        self.setWindowTitle("FSRG Tool")
        self.setStyleSheet("background-color: white;")

        layout = QGridLayout(self)

        start_box = QGroupBox("Before You Start:")
        start_box.setMinimumWidth(350)
        start_box.setObjectName("start_box")
        start_box.setStyleSheet("#start_box {background-color: white; border: 3px solid rgb(196, 214, 0); font-size: 15px; font-weight: bold; border-radius: 6px; margin-top: 12px;}QGroupBox::title {subcontrol-origin: margin; left: 3px; padding: 0px 0px 5px 0px;}")

        box_layout = QVBoxLayout()
        start_box.setLayout(box_layout)

        self.check1 = QCheckBox("RFEM6 is open.")
        #self.check2 = QCheckBox("The Web API license is active and and the server is set to automatically start (see program settings)")
        self.check3 = QCheckBox("No auto-save pop-ups from RFEM6 are visible")
        self.check4 = QCheckBox("You have a printout report saved in RFEM6")
        box_layout.addWidget(self.check1)
        #box_layout.addWidget(self.check2)
        box_layout.addWidget(self.check3)
        box_layout.addWidget(self.check4)

        layout.addWidget(start_box)

        self.continue_button = QPushButton("Continue")
        self.continue_button.setEnabled(False)
        layout.addWidget(self.continue_button)

        self.check1.stateChanged.connect(self.update_continue_button_state)
        #self.check2.stateChanged.connect(self.update_continue_button_state)
        self.check3.stateChanged.connect(self.update_continue_button_state)
        self.check4.stateChanged.connect(self.update_continue_button_state)

        self.continue_button.clicked.connect(self.accept)

    def update_continue_button_state(self):
        """
        Update the state of the continue button based on the state of the checkboxes.
        The continue button is enabled only if all checkboxes are checked.
        """
        self.continue_button.setEnabled((self.check1.isChecked() and self.check3.isChecked() and self.check4.isChecked()))

