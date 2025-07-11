import docx
import pythoncom
import win32com
from win32com import client
import os

def Doc_to_Docx(path):
    pythoncom.CoInitialize()
    word=win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(os.getcwd()+'/'+path)
    print(doc)


def Figure(File_Path):
    EXT=File_Path.split('.')[-1]
    if EXT == 'doc':  #先将doc转化成docx
        Doc_to_Docx(File_Path)
    else :
        doc=docx.Document(File_Path)
        print(doc.paragraphs[0].text)


if __name__ == "__main__":
    Figure('Test_doc/abc.doc')



