import docx
import pythoncom
import win32com
from win32com import client
import os


Keyword_list=["无人机","巡检","联系人"]

def Doc_to_Docx(path):
    #利用win32com将doc转化为docx
    pythoncom.CoInitialize()
    word=win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(os.getcwd()+'/'+path)  #需要绝对路径
    doc.SaveAs(os.getcwd()+'/'+path+"x",12)  #将doc后缀改成docx
    doc.Close()
    word.Quit()

def Figure_doc(File_Path):
    EXT=File_Path.split('.')[-1]
    if EXT == 'doc':  #先将doc转化成docx
        Doc_to_Docx(File_Path)
        File_Path += "x"

    doc=docx.Document(File_Path)  #读取word文件
    paragraphs = doc.paragraphs   #处理文字内容
    for para in paragraphs:
        if para.text.find("无人机") != -1:
            print(para.text)

    tables=doc.tables  #处理表格内容
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                if any(kw in cell.text for kw in Keyword_list):
                    print(cell.text)



if __name__ == "__main__":
    Figure_doc('Test_doc/table.docx')



