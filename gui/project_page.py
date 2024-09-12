from PyQt5.QtWidgets import QPushButton, QVBoxLayout, QWidget, QLabel, QLineEdit, QGridLayout, QGroupBox, QHBoxLayout, QFileDialog, QDialog, QRadioButton, QButtonGroup, QApplication, QComboBox, QMessageBox
from PyQt5.QtCore import pyqtSignal
from html2word import HTMLToWordConverter
from replacement import DocumentWordReplacer, WordReplacement
from RFEM.initModel import Client, Model, openFile
import psutil, os
from RFEM.Reports.printoutReport import PrintoutReport
from info import TableInfo, ImageInfo
import time
from file_manager import FileManager as fm
import sys
import tempfile

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


    def generate_rfem_report(self):
        folder_path = fm.create_folder_desktop("FSRG")
        report_count = int(self.printout_reports.currentText())  # Set the expected number of reports here

        report_paths = [
            os.path.join(folder_path, f"pr{i+1}.html")
            for i in range(report_count)
        ]

        temp_files = [
            os.path.join(folder_path, f"report_op{i+1}.docx")
            for i in range(report_count)
        ]

        word_path = fm.resource_path("Template.docx")
        self.print_debug_info(word_path)

        for i in range(report_count):
            PrintoutReport.exportToHTML(i+1,report_paths[i],model=self.model)
            self.wait_for_file_size_stabilization(report_paths[i])
            if i == 0:
                report = HTMLToWordConverter(word_path, report_paths[i])
            else: 
                report = HTMLToWordConverter(temp_files[i-1], report_paths[i])
            report._delete_last_page_in_template()
            report.process_html_file()
            report.extract_image_files()
            report.extract_captions()
            report.add_images_to_word_document()
            report.save(temp_files[i])

            if i > 0:
                os.remove(temp_files[i-1])

        replacer = DocumentWordReplacer(temp_files[-1])
        replacer.add_replacement('Projekttitel', self.project_title.text())
        replacer.add_replacement('Berichttitel', self.report_title.text())
        replacer.add_replacement('XXXX-BHE-XX-XX-XX-X-XXXX', self.doc_no.text())
        replacer.add_replacement('Projektnummer', self.project_no.text())
        replacer.add_replacement('[Author]', self.author.text())
        modified_file_path = replacer.replace_words(folder_path)
        os.remove(temp_files[-1])
        
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Report successfully generated.\nPlease check the folder named FSRG on your Desktop.")
        msg.setWindowTitle("Report Generated!")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
        self.model.clientModel.service.close_connection()
        return modified_file_path

    def wait_for_file_size_stabilization(self, file_path):
        previous_size = -1
        while True:
            if os.path.exists(file_path):
                current_size = os.path.getsize(file_path)
                if current_size > 0 and current_size == previous_size:
                    break
                previous_size = current_size
            else:
                print(f"File not found: {file_path}")
            time.sleep(1)

    def print_debug_info(self, word_path):
        print("-------- Debug Information --------")
        print(f"Running as: {'Executable' if getattr(sys, 'frozen', False) else 'Script'}")
        print(f"Template.docx path: {word_path}")
        print(f"Template.docx exists: {os.path.exists(word_path)}")
        print(f"Current working directory: {os.getcwd()}")
        print(f"Temporary directory: {tempfile.gettempdir()}")
        print(f"sys._MEIPASS: {getattr(sys, '_MEIPASS', 'Not set - not running as PyInstaller bundle')}")
        print("------------------------------------")

    def upload_rfem_model(self):
        # Open a file dialog and let the user select a .rf6 file
        rfem_fp, _ = QFileDialog.getOpenFileName(self, "Open RFEM Model", "", "RFEM Model Files (*.rf6)")

        if rfem_fp:
            self.save_button.setEnabled(True)
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
            self.model = Model(False,str(model_list[0]))
        self.save_button.setEnabled(True)

if __name__ == "__main__":
    app = QApplication([])
    window = ProjectPage()
    window.show()
    app.exec_()