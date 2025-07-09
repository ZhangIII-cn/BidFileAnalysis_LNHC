import unzip
import json
import chardet
import codecs
import os

def dfs_extract(target_dir,output_dir,ifroot=False,father=None):  #递归解析所有文件
    if ifroot:
        unzip.extract_all_archives(target_dir,output_dir)
        for Folder in os.listdir(output_dir):
            for File in os.listdir(output_dir+'/'+Folder):
                print(Folder+'/'+File)

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

    dfs_extract(target_dir, output_dir,ifroot=True)  #解压全部文件

