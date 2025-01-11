import os
from datetime import datetime
from typing import Union, BinaryIO
from utils.log_handler import LogHandler

class FileHandler:
    """文件处理工具类，用于处理文件相关操作"""
    
    logger = LogHandler().get_logger('FileHandler')
    
    @classmethod
    def ensure_dir(cls, directory: str):
        """确保目录存在，如果不存在则创建
        
        Args:
            directory: 目录路径
        """
        try:
            if not os.path.exists(directory):
                os.makedirs(directory)
                cls.logger.debug("创建目录: {}", directory)
        except Exception as e:
            cls.logger.error("创建目录失败: {}, 错误: {}", 
                           directory, LogHandler.format_error(e))
            raise
    
    @classmethod
    def generate_unique_filename(cls, filename: str) -> str:
        """生成唯一的文件名，避免重复
        
        Args:
            filename: 原始文件名
            
        Returns:
            str: 唯一的文件名
        """
        name, ext = os.path.splitext(filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{name}_{timestamp}{ext}"
    
    @classmethod
    def join_paths(cls, *paths: str) -> str:
        """连接路径
        
        Args:
            *paths: 路径片段
            
        Returns:
            str: 连接后的完整路径
        """
        return os.path.join(*paths)
    
    @classmethod
    def save_file(cls, content: Union[bytes, BinaryIO], filepath: str):
        """保存文件内容
        
        Args:
            content: 文件内容
            filepath: 文件保存路径
        """
        try:
            # 确保目录存在
            cls.ensure_dir(os.path.dirname(filepath))
            
            # 保存文件
            with open(filepath, 'wb') as f:
                if isinstance(content, bytes):
                    f.write(content)
                else:
                    f.write(content.read())
                    
            cls.logger.debug("文件保存成功: {}", filepath)
            
        except Exception as e:
            cls.logger.error("保存文件失败: {}, 错误: {}", 
                           filepath, LogHandler.format_error(e))
            raise
    
    @classmethod
    def get_file_size(cls, filepath: str) -> int:
        """获取文件大小
        
        Args:
            filepath: 文件路径
            
        Returns:
            int: 文件大小（字节）
        """
        try:
            return os.path.getsize(filepath)
        except Exception as e:
            cls.logger.error("获取文件大小失败: {}, 错误: {}", 
                           filepath, LogHandler.format_error(e))
            return 0 