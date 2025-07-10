import os
import sys
import zipfile  # Python 内置，无需安装
import py7zr  # 需要 pip install py7zr
import rarfile  # 需要 pip install rarfile
import patoolib  # 需要 pip install patool
from patoolib.util import PatoolError
import logging

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_zip(file_path, output_dir):
    """解压 ZIP 文件"""
    # print(output_dir)
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            for file in zip_ref.namelist():
                # print(file_path.split('.')[-2],":",file)
                zip_ref.extract(file,output_dir)
            #zip_ref.extractall(output_dir)
        return True
    except Exception as e:
        logger.error(f"ZIP解压错误: {str(e)}")
        return False


def extract_7z(file_path, output_dir):
    """解压 7z 文件"""
    try:
        with py7zr.SevenZipFile(file_path, mode='r') as z:
            z.extractall(output_dir)
        return True
    except Exception as e:
        logger.error(f"7z解压错误: {str(e)}")
        return False


def extract_rar(file_path, output_dir):
    """解压 RAR 文件"""
    try:
        with rarfile.RarFile(file_path, 'r') as rar_ref:
            rar_ref.extractall(output_dir)
        return True
    except Exception as e:
        logger.error(f"RAR解压错误: {str(e)}")
        return False


def extract_with_patool(file_path, output_dir):
    """使用 patool 作为后备方案"""
    try:
        patoolib.extract_archive(file_path, outdir=output_dir)
        return True
    except PatoolError as e:
        logger.error(f"patool解压错误: {str(e)}")
        return False


def extract_all_archives(directory, output_dir=None):
    """
    解压指定目录下的所有压缩文件
    :param directory: 包含压缩文件的目录路径
    :param output_dir: 解压输出目录（默认为每个压缩文件同名的文件夹）
    :param password: 可选密码，用于加密压缩文件
    """
    if not os.path.isdir(directory):
        logger.error(f"目录 '{directory}' 不存在")
        return False

    # 支持的压缩文件扩展名和处理方法
    supported_extensions = {
        '.zip': extract_zip,
        '.7z': extract_7z,
        '.rar': extract_rar,
    }

    success = True
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        if not os.path.isfile(file_path):
            continue

        # 获取文件扩展名
        ext = os.path.splitext(file)[1].lower()

        try:
            # 创建输出目录
            if output_dir is None:
                output_subdir = os.path.join(directory, os.path.splitext(file)[0])
            else:
                output_subdir = os.path.join(output_dir, os.path.splitext(file)[0])

            os.makedirs(output_subdir, exist_ok=True)

            # 选择解压方法
            if ext in supported_extensions:
                result = supported_extensions[ext](file_path, output_subdir)
            else:
                # 检查是否是支持的压缩文件
                if patoolib.is_archive(file_path):
                    result = extract_with_patool(file_path, output_subdir)
                else:
                    logger.warning(f"跳过不支持的文件: {file}")
                    continue

            if result:
                # logger.info(f"成功解压到: {output_subdir}")
                pass
            else:
                success = False
                logger.error(f"解压失败: {file}")

        except Exception as e:
            success = False
            logger.error(f"处理 {file} 时发生错误: {str(e)}")

    return success


