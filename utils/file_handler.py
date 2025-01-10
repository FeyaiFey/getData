import os
from datetime import datetime

class FileHandler:
    @staticmethod
    def ensure_dir(directory):
        """确保目录存在"""
        os.makedirs(directory, exist_ok=True)

    @staticmethod
    def generate_unique_filename(original_filename, prefix=''):
        """生成唯一的文件名"""
        timestamp = datetime.now().strftime('%Y%m%d')
        if prefix:
            return f"{prefix}_{timestamp}_{original_filename}"
        return f"{timestamp}_{original_filename}"

    @staticmethod
    def save_file(content, filepath):
        """保存文件内容"""
        with open(filepath, 'wb') as f:
            f.write(content)

    @staticmethod
    def join_paths(*paths):
        """连接路径"""
        return os.path.join(*paths) 