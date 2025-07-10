import unzip
import json
import chardet
import codecs
import os

def dfs_extract(target_dir,output_dir,ifRoot=False,ifFolder=False,father=None):  #递归解析所有文件
    print(target_dir)
    if ifRoot:  #根目录默认解压所有压缩包并遍历
        unzip.extract_all_archives(target_dir,output_dir)
        for Folder in os.listdir(output_dir):
            for File in os.listdir(output_dir+'/'+Folder):
                File_str = File.split('.')
                File_EXT = File_str[-1]
                File_Name = File_str[0]  #不包括后缀名的name
                if len(File_str == 1): # File为文件夹 无后缀
                    dfs_extract( output_dir+'/'+Folder+'/'+File, False,True)
                elif (File_EXT == 'zip'):
                    dfs_extract(output_dir+'/'+Folder+'/'+File,output_dir+'/'+Folder+'/'+File_Name,False)
    else:    #非根目录
        unzip.extract_zip(target_dir,output_dir)
        for File in os.listdir(output_dir):
            File_EXT = File.split('.')[-1]
            File_Name = File.split('.')[0]  # 不包括后缀名的name

            if (File_EXT == 'zip'):
                dfs_extract(output_dir + '/' + File, output_dir + '/'  + File_Name, False)
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

