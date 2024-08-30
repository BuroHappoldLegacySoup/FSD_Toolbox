from PyQt5.QtWidgets import QPushButton, QVBoxLayout, QWidget, QLabel, QLineEdit, QGridLayout, QGroupBox, QHBoxLayout, QFileDialog, QDialog, QRadioButton, QButtonGroup, QApplication
from PyQt5.QtCore import pyqtSignal
from html2word import HTMLToWordConverter
from replacement import DocumentWordReplacer, WordReplacement
from html2png2word import ImageInfo, HTMLImageExtractor
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
        report_path_1 = folder_path + r"\pr1.html"
        PrintoutReport.exportToHTML(1,report_path_1,self.model)
        
        previous_size = -1
        while True:
            current_size = os.path.getsize(report_path_1)
            if current_size > 0 and current_size == previous_size:
                break
            previous_size = current_size
            time.sleep(1)
        
        word_path = fm.resource_path("Template.docx")
        print("-------- Debug Information --------")
        print(f"Running as: {'Executable' if getattr(sys, 'frozen', False) else 'Script'}")
        print(f"Template.docx path: {word_path}")
        print(f"Template.docx exists: {os.path.exists(word_path)}")
        print(f"Current working directory: {os.getcwd()}")
        print(f"Temporary directory: {tempfile.gettempdir()}")
        print(f"sys._MEIPASS: {getattr(sys, '_MEIPASS', 'Not set - not running as PyInstaller bundle')}")
        print("------------------------------------")

        temp_1 = folder_path + r"\report_op1.docx"
        report_pt1 = HTMLToWordConverter(word_path , report_path_1)
        report_pt1._delete_last_page_in_template()
        report_pt1.process_html_file()
        report_pt1.extract_image_files()
        report_pt1.extract_captions()
        report_pt1.add_images_to_word_document()

        # Save the Word document
        report_pt1.save(temp_1)

        report_path_2 = folder_path + r"\pr2.html"
        PrintoutReport.exportToHTML(2,report_path_2,self.model)
        previous_size = -1
        while True:
            current_size = os.path.getsize(report_path_2)
            if current_size > 0 and current_size == previous_size:
                break
            previous_size = current_size
            time.sleep(1)
        temp_2 = folder_path + r"\report_op2.docx"

        report_pt2 = HTMLToWordConverter(temp_1,report_path_2)
        report_pt2._delete_last_page_in_template()
        report_pt2.process_html_file()
        report_pt2.extract_image_files()
        report_pt2.extract_captions()
        report_pt2.add_images_to_word_document()
        # Save the Word document
        report_pt2.save(temp_2)
        fm.delete_file(temp_1)

        
        report_path_3 = folder_path + r"\pr3.html"
        PrintoutReport.exportToHTML(3,report_path_3,self.model)
        previous_size = -1
        while True:
            current_size = os.path.getsize(report_path_3)
            if current_size > 0 and current_size == previous_size:
                break
            previous_size = current_size
            time.sleep(1)
        temp_3 = folder_path + r"\report_op3.docx"
        report_pt3 = HTMLToWordConverter(temp_2,report_path_3)
        report_pt3._delete_last_page_in_template()
        report_pt3.process_html_file()
        report_pt3.extract_image_files()
        report_pt3.extract_captions()
        report_pt3.add_images_to_word_document()
        report_pt3.save(temp_3)
        fm.delete_file(temp_2)


        replacer = DocumentWordReplacer(temp_3)
        
        replacer.add_replacement('Projekttitel', self.project_title.text())
        replacer.add_replacement('Berichttitel', self.report_title.text())
        replacer.add_replacement('XXXX-BHE-XX-XX-XX-X-XXXX', self.doc_no.text())
        replacer.add_replacement('Projektnummer', self.project_no.text())
        replacer.add_replacement('[Author]', self.author.text())
        
        modified_file_path = replacer.replace_words(folder_path)
        fm.delete_file(temp_3)

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
