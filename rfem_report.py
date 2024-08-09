from dataclasses import dataclass
from RFEM.initModel import openFile, closeAllModels
from RFEM.Reports.printoutReport import PrintoutReport

@dataclass
class RFEMReport:
    """
    A class used to represent an RFEM Report

    ...

    Attributes
    ----------
    file_path : str
        a formatted string to declare file path
    report_path : str
        a formatted string to declare report path

    Methods
    -------
    open_model():
        Opens the RFEM model file
    export_report():
        Exports the RFEM model report to HTML
    generate_report():
        Generates the RFEM model report
    """

    file_path: str
    report_path: str

    def _open_model(self):
        """
        Opens the RFEM model file
        """
        openFile(self.file_path)

    def _export_report(self):
        """
        Exports the RFEM model report to HTML
        """
        PrintoutReport.exportToHTML(1, self.report_path)

    def generate_report(self):
        """
        Generates the RFEM model report by opening the model,
        exporting the report, and then closing all models.
        """
        self._open_model()
        self._export_report()
        closeAllModels()

if __name__ == "__main__":
    rfem_fp = r"C:\Users\vmylavarapu\Downloads\Decke_ü_8OG-10OG_210122.rf6"
    report_path = r"C:\Users\vmylavarapu\Desktop\pr.html"

    report = RFEMReport(rfem_fp, report_path)
    report.generate_report()
