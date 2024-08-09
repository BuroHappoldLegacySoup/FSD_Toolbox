from rfem_report import RFEMReport
from html2word import HTMLToWordTableConverter, TableInfo
from replacement import DocumentWordReplacer, WordReplacement


def main():

    # generate the printout report in as an HTML
    # preprequisite: A printout report exists
    rfem_fp = r"C:\Users\vmylavarapu\Downloads\test.rf6"
    report_path = r"C:\Users\vmylavarapu\Desktop\pr.html"
    word_path = r"C:\Users\vmylavarapu\Desktop\Template.docx"
    new_word_path = r"C:\Users\vmylavarapu\Desktop\report.docx"

    report = RFEMReport(rfem_fp, report_path)
    report.generate_report()

    converter = HTMLToWordTableConverter(word_path)

    # Define the tables to extract
    table_info_list = [
        TableInfo('1.1 Materials', 'Materials'),
        #TableInfo('1.2 Surfaces', 'Surfaces'),
        TableInfo('3.1 Nodal Supports', 'Nodal Supports'),
        TableInfo('4.1 Line Supports', 'Line Supports'),
        TableInfo('1.3 Sections', 'Cross sections'),
        TableInfo('4.2 Line Hinges', 'Line Hinges'),
        TableInfo('4.1 Member Hinges', 'Member Hinges'),
        TableInfo('1.6 Members', 'Members'),
        TableInfo('7.1 Load Cases', 'Load Cases')
    ]

    # Process the HTML file
    converter.process_html_file(report_path, table_info_list)
    converter.save(new_word_path)
    replacer = DocumentWordReplacer(new_word_path)
    
    replacer.add_replacement('Projekttitel', 'new')
    replacer.add_replacement('Berichttitel', 'replacement1')
    replacer.add_replacement('XXXX-BHE-XX-XX-XX-X-XXXX', 'replacement2')
    
    modified_file_path = replacer.replace_words()


    return None

if __name__ == "__main__":
    main()
