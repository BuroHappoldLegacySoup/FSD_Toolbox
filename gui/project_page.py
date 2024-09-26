from PyQt5.QtWidgets import QPushButton, QVBoxLayout, QWidget, QLabel, QLineEdit, QGridLayout, QGroupBox, QHBoxLayout, QFileDialog, QDialog, QRadioButton, QButtonGroup, QApplication, QComboBox
from PyQt5.QtCore import pyqtSignal
from gui.rep_gen import RepGen as RG
from RFEM.initModel import Client, Model, openFile


class ModelSelectionDialog(QDialog):
    def __init__(self, model_list, parent=None):
        super(ModelSelectionDialog, self).__init__(parent)
        self.setWindowTitle('Select Model')
        self.layout = QVBoxLayout(self)
        
        self.button_group = QButtonGroup(self)
        for model in model_list:
            radio_button = QRadioButton(model)
            self.button_group.addButton(radio_button)
            self.layout.addWidget(radio_button)
        
        self.ok_button = QPushButton('OK', self)
        self.ok_button.clicked.connect(self.accept)
        self.layout.addWidget(self.ok_button)
        
    def get_selected_model(self):
        selected_button = self.button_group.checkedButton()
        if selected_button:
            return selected_button.text()
        return None

class ProjectPage(QWidget):
    """
    This class represents the first page. This is where all the project related information is collected.
    """
    next_clicked = pyqtSignal(dict)

    def __init__(self):
        super().__init__()
        self.user_inputs = {}
        self.initUI()

    def initUI(self):
        layout = QGridLayout(self)

        project_box = QGroupBox("Project")
        project_box.setObjectName("project_box")
        project_box.setStyleSheet("#project_box {background-color: white; border: 3px solid rgb(196, 214, 0); font-size: 15px; font-weight: bold; border-radius: 6px; margin-top: 12px;}QGroupBox::title {subcontrol-origin: margin; left: 3px; padding: 0px 0px 5px 0px;}")
        layout.addWidget(project_box)

        project_layout = QVBoxLayout(project_box)

        self.project_title_label = QLabel('Project Title')
        self.project_title = QLineEdit('Project Title')
        project_layout.addWidget(self.project_title_label)
        project_layout.addWidget(self.project_title)

        self.report_title_label = QLabel('Report Title')
        self.report_title = QLineEdit('Report Title')
        project_layout.addWidget(self.report_title_label)
        project_layout.addWidget(self.report_title)

        self.doc_no_label = QLabel('Document Number')
        self.doc_no = QLineEdit('XXXX-BHE-XX-XX-XX-X-XXXX')
        project_layout.addWidget(self.doc_no_label)
        project_layout.addWidget(self.doc_no)

        self.project_no_label = QLabel('Project Number')
        self.project_no = QLineEdit('Project Number')
        project_layout.addWidget(self.project_no_label)
        project_layout.addWidget(self.project_no)

        self.author_label = QLabel('Author')
        self.author = QLineEdit('Author')
        project_layout.addWidget(self.author_label)
        project_layout.addWidget(self.author)

        self.printout_reports_label = QLabel('Number of printout reports in the RFEM6 model')
        self.printout_reports = QComboBox()
        self.printout_reports.addItems([str(i) for i in range(1, 11)])
        self.printout_reports.setCurrentIndex(3)  # Default value is 4 (index 3)
        project_layout.addWidget(self.printout_reports_label)
        project_layout.addWidget(self.printout_reports)

        layout.addWidget(project_box)

        upload_rfem_model_button = QPushButton('Upload RFEM Model')
        upload_rfem_model_button.setStyleSheet("background-color: white;")
        upload_rfem_model_button.clicked.connect(self.upload_rfem_model)

        use_active_model_button = QPushButton('Use Active Model')
        use_active_model_button.setStyleSheet("background-color: white;")
        use_active_model_button.clicked.connect(self.use_active_model)

        button_widget = QWidget()
        button_layout = QHBoxLayout(button_widget)
        button_layout.addWidget(upload_rfem_model_button)
        button_layout.addWidget(use_active_model_button)

        layout.addWidget(button_widget, layout.rowCount(), 0)

        self.save_button = QPushButton('Save and Proceed')
        self.save_button.setStyleSheet("background-color: white;")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.generate_rfem_report)

        layout.addWidget(self.save_button, layout.rowCount(), 0)

        # Add the new button below "Save and Proceed"
        self.update_button = QPushButton('Update Existing Report')
        self.update_button.setStyleSheet("background-color: white;")
        self.update_button.setEnabled(False)
        self.update_button.clicked.connect(self.update_existing_report)

        layout.addWidget(self.update_button, layout.rowCount(), 0)

    def generate_rfem_report(self):
        rg = RG(
            self.project_title.text(),
            self.report_title.text(),
            self.doc_no.text(),
            self.project_no.text(),
            self.author.text(),
            self.printout_reports.currentText(),
            self.model
        )
        rg.generate_rfem_report_as_html()

    def update_existing_report(self):
        # Implement the logic for updating an existing report
        pass

    def upload_rfem_model(self):
        # Open a file dialog and let the user select a .rf6 file
        rfem_fp, _ = QFileDialog.getOpenFileName(self, "Open RFEM Model", "", "RFEM Model Files (*.rf6)")

        if rfem_fp:
            self.save_button.setEnabled(True)
            self.update_button.setEnabled(True)
            self.model = openFile(rfem_fp)
        else:
            return None

    def use_active_model(self):
        client = Client('http://localhost:8081/wsdl')
        model_list = client.service.get_model_list().name
        if len(model_list) > 1 :
            dialog = ModelSelectionDialog(model_list)
            if dialog.exec_() == QDialog.Accepted:
                selected_model = dialog.get_selected_model()
                self.model = Model(False,str(selected_model))
        else:
            self.model = Model(False,str(model_list))
        self.save_button.setEnabled(True)
        self.update_button.setEnabled(True)

if __name__ == "__main__":
    app = QApplication([])
    window = ProjectPage()
    window.show()
    app.exec_()
