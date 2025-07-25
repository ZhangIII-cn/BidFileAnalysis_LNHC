import unzip
import json
import chardet
import codecs
import os
import file_analyzer


def copy_worksppace(file_path):
    #因为目录过长，win接口无法读取文件，需要先临时移动到workspace目录进行文件分析
    workspace_path= ""
    return workspace_path


def figure_doc(path):
    work_path = copy_worksppace(path)
    file_analyzer.Figure_doc(work_path)

def figure_xls():
    print(2)
    pass

def figure_pdf():
    print(3)
    pass

def dfs_extract(target_dir,output_dir,ifRoot=False,ifFolder=False,father=None):
    # 递归解析所有文件
    # 返回值递归回根目录位置，标记当前Folder的状态，决定输出文件的位置
    # 返回值 1：符合筛选条件  2：不符合筛选条件  3：内部存在未知文件，需要人工核查
    #print(target_dir)
    if ifRoot:  #根目录默认解压所有压缩包并遍历
        unzip.extract_all_archives(target_dir,output_dir)
        for Folder in os.listdir(output_dir):
            for File in os.listdir(output_dir+'/'+Folder):
                File_strs = File.split('.')
                File_EXT = File_strs[-1]
                File_Name = File_strs[0]  #不包括后缀名的name
                if len(File_strs) == 1 : # File为文件夹 无后缀
                    dfs_extract( output_dir+'/'+Folder+'/'+File, output_dir+'/'+Folder+'/'+File,False,True)
                elif (File_EXT == 'zip'):
                    dfs_extract( output_dir+'/'+Folder+'/'+File,output_dir+'/'+Folder+'/'+File_Name,False,False)
    else:    #非根目录
        if not ifFolder:
            unzip.extract_zip(target_dir,output_dir)
            dfs_extract(output_dir,output_dir,False,True) #解压完一定是文件夹
        else :
            dir = target_dir if ifFolder else output_dir
            for File in os.listdir(dir):
                File_strs = File.split('.')
                File_EXT = File_strs[-1]
                File_Name = File_strs[0]  # 不包括后缀名的name
                if len(File_strs) == 1 : # File为文件夹 无后缀
                    dfs_extract(dir + '/'  + File, dir + '/' + File, False, True)
                elif (File_EXT == 'zip'):
                    dfs_extract(dir + '/' + File, dir + '/'  + File_Name, False,False)
                elif (File_EXT == 'doc' or File_EXT == 'docx'):
                    #print(dir+'/'+File)
                    figure_doc(dir+'/'+File)
                elif (File_EXT == 'xls' or File_EXT == 'xlsx'):
                    figure_xls()
                elif (File_EXT == 'pdf'):
                    figure_pdf()
                else :
                    print("存在未知格式文件："+dir+File)
                    return 3

    #file_ext = file.split('.')
    return


if __name__ == "__main__":
    #--------------------------------利用同一目录json读取源和目的目录-------------------------------
    #target_dir = 'C:/Users/Administrator/Desktop/临时文件'
    #output_directory = 'C:/Users/Administrator/Desktop/临时文件/New'
    with open('dir.json', 'r', newline='',encoding='utf-8') as rf:
        data=json.load(rf)
        target_dir = data['target_dir']
        output_dir = data['output_dir']

    dfs_extract(target_dir, output_dir,ifRoot=True)  #解压全部文件

