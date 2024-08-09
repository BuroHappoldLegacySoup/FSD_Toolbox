from dataclasses import dataclass
from typing import List
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup
from docx.enum.text import WD_ALIGN_PARAGRAPH

@dataclass
class TableInfo:
    """
    Dataclass to store information about a table to be extracted and converted.
    
    Attributes:
        heading_text (str): The text of the heading to search for in the HTML.
        title (str): The title to be used for the table in the Word document.
    """
    heading_text: str
    title: str

class HTMLToWordTableConverter:
    """
    A class to convert HTML tables to Word document tables.
    
    This class provides functionality to extract tables from HTML content,
    convert them to Word tables, and save them in a Word document.
    """

    def __init__(self, doc_path: str):
        """
        Initialize the converter with an existing Word document.
        
        Args:
            doc_path (str): The path to the existing Word document.
        """
        self.doc = Document(doc_path)

    @staticmethod
    def extract_table_after_heading(soup: BeautifulSoup, heading_text: str) -> str:
        """
        Extract the HTML table that follows a specific heading.
        
        Args:
            soup (BeautifulSoup): The BeautifulSoup object containing the HTML content.
            heading_text (str): The text of the heading to search for.
        
        Returns:
            str: The HTML string of the table if found, None otherwise.
        """
        heading = soup.find('h2', string=heading_text)
        if heading:
            target_table = heading.find_next('table')
            if target_table:
                return str(target_table)
        return None

    @staticmethod
    def apply_cell_formatting(cell, color_hex: str):
        """
        Apply background color to a cell.
        
        Args:
            cell: The Word table cell to format.
            color_hex (str): The hex color code to apply as background.
        """
        if color_hex:
            cell_xml = cell._element
            shading = OxmlElement('w:shd')
            shading.set(qn('w:fill'), color_hex)
            cell_xml.get_or_add_tcPr().append(shading)

    @staticmethod
    def remove_table_borders(table):
        """
        Remove borders from the table.
        
        Args:
            table: The Word table to remove borders from.
        """
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblBorders = OxmlElement('w:tblBorders')
        for border in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            border_element = OxmlElement(f'w:{border}')
            border_element.set(qn('w:val'), 'none')
            border_element.set(qn('w:space'), '0')
            border_element.set(qn('w:sz'), '0')
            tblBorders.append(border_element)
        tblPr.append(tblBorders)

    def create_word_table_from_html(self, html_content: str, title: str):
        """
        Create a Word table from HTML content and add it to the document with a title.
        
        Args:
            html_content (str): The HTML content of the table.
            title (str): The title to be added before the table in the Word document.
        """
        soup = BeautifulSoup(html_content, 'html.parser')
        table = soup.find('table')
        self.doc.add_heading(title, level=1)

        headers = table.find_all('tr')[0].find_all(['th', 'td'])
        max_columns = sum(int(th.get('colspan', 1)) for th in headers)
        rows = table.find_all('tr')
        num_rows = len(rows)

        word_table = self.doc.add_table(rows=num_rows, cols=max_columns)
        word_table.style = 'Table Grid'
        self.remove_table_borders(word_table)

        self._fill_table_content(word_table, rows, max_columns)
        self._remove_empty_columns(word_table)
        self._remove_empty_rows(word_table)
        self._apply_row_colors(word_table)

    def _fill_table_content(self, word_table, rows, max_columns):
        """
        Fill the Word table with content from the HTML table.
        
        Args:
            word_table: The Word table to fill.
            rows: The rows from the HTML table.
            max_columns (int): The maximum number of columns in the table.
        """
        for row_idx, row in enumerate(rows):
            cells = row.find_all(['th', 'td'])
            col_idx = 0
            while cells and col_idx < max_columns:
                cell = cells.pop(0)
                colspan = int(cell.get('colspan', 1))
                cell_text = cell.get_text(strip=True)
                cell_style = cell.get('style', '')
                cell_color = None
                if 'background-color:' in cell_style:
                    color_hex = cell_style.split('background-color:')[1].split(';')[0].strip()
                    if color_hex.startswith('#'):
                        cell_color = color_hex[1:]
                self._fill_cell_content(word_table, row_idx, col_idx, cell_text, cell_color)
                col_idx += colspan

    def _fill_cell_content(self, word_table, row: int, col: int, content: str, color_hex: str):
        """
        Fill the content and background color of a cell in the Word table.
        
        Args:
            word_table: The Word table containing the cell.
            row (int): The row index of the cell.
            col (int): The column index of the cell.
            content (str): The text content to fill in the cell.
            color_hex (str): The hex color code for the cell background.
        """
        word_cell = word_table.cell(row, col)
        word_cell.text = content
        if color_hex:
            self.apply_cell_formatting(word_cell, color_hex)

    def _remove_empty_columns(self, word_table):
        """
        Remove empty columns from the Word table.
        
        Args:
            word_table: The Word table to process.
        """
        columns_to_remove = [i for i in range(len(word_table.columns)) if self._is_column_empty(word_table, i)]
        for col_idx in reversed(columns_to_remove):
            for row in word_table.rows:
                cell = row.cells[col_idx]
                cell._element.getparent().remove(cell._element)

    @staticmethod
    def _is_column_empty(word_table, col_idx: int) -> bool:
        """
        Check if a column is entirely empty.
        
        Args:
            word_table: The Word table to check.
            col_idx (int): The index of the column to check.
        
        Returns:
            bool: True if the column is empty, False otherwise.
        """
        return all(row.cells[col_idx].text.strip() == '' for row in word_table.rows)

    def _remove_empty_rows(self, word_table):
        """
        Remove empty rows from the Word table.
        
        Args:
            word_table: The Word table to process.
        """
        rows_to_remove = [row for row in word_table.rows if self._is_row_empty(row)]
        for row in rows_to_remove:
            tbl = word_table._tbl
            tbl.remove(row._tr)

    @staticmethod
    def _is_row_empty(row) -> bool:
        """
        Check if a row is entirely empty.
        
        Args:
            row: The row to check.
        
        Returns:
            bool: True if the row is empty, False otherwise.
        """
        return all(cell.text.strip() == '' for cell in row.cells)

    def _delete_last_page_in_template(self):
        for element in reversed(self.doc.element.body):
            if element.tag.endswith('sectPr'):
                self.doc.element.body.remove(element)
                break

    def _apply_row_colors(self, word_table):
        """
        Apply the background color of the second cell to the remaining cells of each row.
        
        Args:
            word_table: The Word table to process.
        """
        for row in word_table.rows:
            if len(row.cells) > 1:
                second_cell_color = row.cells[1]._element.xpath('.//w:shd/@w:fill')
                if second_cell_color:
                    second_cell_color = second_cell_color[0]
                    for cell in row.cells[2:]:
                        self.apply_cell_formatting(cell, second_cell_color)

    def save(self, filename: str):
        """
        Save the Word document to a file.
        
        Args:
            filename (str): The name of the file to save the document to.
        """
        self.doc.save(filename)

    def process_html_file(self, file_path: str, table_info_list: List[TableInfo]):
        """
        Process an HTML file, extract tables, and create a Word document.
        
        Args:
            file_path (str): The path to the HTML file to process.
            table_info_list (List[TableInfo]): A list of TableInfo objects specifying the tables to extract.
        """
        with open(file_path, 'r', encoding='utf-8') as file:
            large_html_content = file.read()

        soup = BeautifulSoup(large_html_content, 'html.parser')

        self._delete_last_page_in_template()

        for table_info in table_info_list:
            table_html = self.extract_table_after_heading(soup, table_info.heading_text)
            if table_html:
                self.create_word_table_from_html(table_html, table_info.title)

        print(f"The Word document '{file_path}' was successfully processed.")

# Example usage
if __name__ == "__main__":
    converter = HTMLToWordTableConverter()

    # Define the tables to extract
    table_info_list = [
        TableInfo('1.1 Materials', 'Materials'),
        TableInfo('1.2 Surfaces', 'Surfaces'),
        TableInfo('3.1 Nodal Supports', 'Nodal Supports'),
        TableInfo('4.1 Line Supports', 'Line Supports'),
        TableInfo('1.3 Sections', 'Cross sections'),
        TableInfo('4.2 Line Hinges', 'Line Hinges'),
        TableInfo('4.1 Member Hinges', 'Member Hinges'),
        TableInfo('1.6 Members', 'Members'),
        TableInfo('7.1 Load Cases', 'Load Cases')
    ]

    # Process the HTML file
    converter.process_html_file(r"C:\Users\vmylavarapu\Desktop\pr.html", table_info_list)

    # Save the Word document
    converter.save('Output.docx')
    print("The Word document 'Output.docx' was successfully created.")