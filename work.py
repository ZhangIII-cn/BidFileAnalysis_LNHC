import unzip


if __name__ == "__main__":
    target_dir = 'C:/Users/Administrator/Desktop/临时文件'
    output_directory = 'C:/Users/Administrator/Desktop/临时文件/New'
    unzip.extract_all_archives(target_dir, output_directory)

