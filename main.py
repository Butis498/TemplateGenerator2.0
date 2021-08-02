from FileScanner import FileScanner
from TempleteGenerator import TempleteGenerator
import os

if __name__ == '__main__':

    folder_name = 'Cards'

    cwd = os.getcwd() 
    directory = str(os.path.join(cwd,folder_name))

    if not os.path.isdir(directory):
        os.mkdir(directory)

    fileScannerXls = FileScanner(directory , ".xlsx")
    fileScannerDocs = FileScanner(directory , ".docx")

    docx_files = fileScannerDocs.files
    xlsx_files = fileScannerXls.files

    templeteGenerator = TempleteGenerator()
    
    for xls in xlsx_files:
        templeteGenerator.replaceContentXls(xls)

    for doc in docx_files:
        templeteGenerator.replaceContentDoc(doc)