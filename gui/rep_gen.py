from file_manager import FileManager as fm
from RFEM.Reports.printoutReport import PrintoutReport
from html2word import HTMLToWordConverter
from replacement import DocumentWordReplacer
from PyQt5.QtWidgets import QMessageBox
import os
import time
import sys
import tempfile

class RepGen:

    def __init__(self, project_title, report_title, doc_no, project_no, author, printout_reports, model):
        self.project_title = project_title
        self.report_title = report_title
        self.doc_no = doc_no
        self.project_no = project_no
        self.author = author
        self.printout_reports = printout_reports
        self.model = model

    def generate_rfem_report_as_html(self):
        folder_path = fm.create_folder_desktop("FSRG")
        report_count = int(self.printout_reports)  # Set the expected number of reports here

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
            PrintoutReport.exportToHTML(i+1, report_paths[i], model=self.model)
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
        replacer.add_replacement('Projekttitel', self.project_title)
        replacer.add_replacement('Berichttitel', self.report_title)
        replacer.add_replacement('XXXX-BHE-XX-XX-XX-X-XXXX', self.doc_no)
        replacer.add_replacement('Projektnummer', self.project_no)
        replacer.add_replacement('[Author]', self.author)
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