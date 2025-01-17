import os
from datetime import datetime
from typing import Union, BinaryIO
from utils.log_handler import LogHandler

class FileHandler:
    """文件处理工具类，用于处理文件相关操作
    
    主要功能：
    1. 目录创建和检查
    2. 文件保存和读取
    3. 生成唯一文件名
    4. 路径处理和拼接
    5. 文件大小获取
    """
    
    logger = LogHandler().get_logger('FileHandler', file_level='DEBUG', console_level='INFO')
    
    @classmethod
    def ensure_dir(cls, directory: str):
        """确保目录存在，如果不存在则创建
        
        Args:
            directory: 目录路径
            
        Raises:
            OSError: 创建目录失败时抛出
        """
        try:
            if not os.path.exists(directory):
                os.makedirs(directory)
                cls.logger.debug("创建目录: %s", directory)
        except Exception as e:
            cls.logger.error("创建目录失败 [%s]: %s", directory, LogHandler.format_error(e))
            raise
    
    @classmethod
    def generate_unique_filename(cls, filename: str) -> str:
        """生成唯一的文件名，避免重复
        
        通过添加时间戳来确保文件名的唯一性。
        格式：原文件名_YYYYMMDD_HHMMSS.扩展名
        
        Args:
            filename: 原始文件名
            
        Returns:
            str: 生成的唯一文件名
        """
        name, ext = os.path.splitext(filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{name}_{timestamp}{ext}"
    
    @classmethod
    def join_paths(cls, *paths: str) -> str:
        """连接路径
        
        使用系统特定的路径分隔符连接多个路径片段。
        
        Args:
            *paths: 路径片段列表
            
        Returns:
            str: 连接后的完整路径
        """
        return os.path.join(*paths)
    
    @classmethod
    def save_file(cls, content: Union[bytes, BinaryIO], filepath: str):
        """保存文件内容
        
        支持保存字节内容或文件对象。
        自动创建必要的目录结构。
        
        Args:
            content: 要保存的内容，可以是字节串或文件对象
            filepath: 保存的目标路径
            
        Raises:
            OSError: 文件保存失败时抛出
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
                    
            file_size = cls.get_file_size(filepath)
            cls.logger.debug("已保存文件 [%s] - 大小: %d 字节", filepath, file_size)
            
        except Exception as e:
            cls.logger.error("保存文件失败 [%s]: %s", filepath, LogHandler.format_error(e))
            raise
    
    @classmethod
    def get_file_size(cls, filepath: str) -> int:
        """获取文件大小
        
        Args:
            filepath: 文件路径
            
        Returns:
            int: 文件大小（字节），如果文件不存在或无法访问则返回0
        """
        try:
            return os.path.getsize(filepath)
        except Exception as e:
            cls.logger.error("获取文件大小失败 [%s]: %s", filepath, LogHandler.format_error(e))
            return 0 
    
    @classmethod
    def init_project_directories(cls):
        """初始化项目所需的所有目录结构
        
        根据配置文件创建所有必要的目录：
        1. 下载目录（wip、shipping）
        2. 归档目录（archive）
        3. 汇总目录（Summary）
        4. 供应商特定目录
        
        只在目录不存在时才创建。
        """
        try:
            # 基础目录结构
            base_dirs = [
                # 下载目录
                "downloads/wip/池州华宇进度表",
                "downloads/wip/汉旗进度表",
                "downloads/wip/PSMC进度表",
                "downloads/wip/CSMC_FAB1进度表",
                "downloads/wip/CSMC_FAB2进度表",
                "downloads/wip/荣芯进度表",
                
                # 送货单目录
                "downloads/shipping/池州华宇送货单",
                "downloads/shipping/汉旗送货单",
                "downloads/shipping/芯丰送货单",
                
                # 归档目录
                "downloads/archive/huayu",
                "downloads/archive/hanqi",
                "downloads/archive/xinfeng",
                
                # 汇总目录
                "downloads/shipping/Summary/池州华宇",
                "downloads/shipping/Summary/山东汉旗",
                "downloads/shipping/Summary/江苏芯丰",
                
                # 截图目录
                "screenshots"
            ]
            
            # 检查并创建必要的目录
            created_count = 0
            for dir_path in base_dirs:
                if not os.path.exists(dir_path):
                    os.makedirs(dir_path)
                    cls.logger.debug("已创建目录: %s", dir_path)
                    created_count += 1
                    
            if created_count > 0:
                cls.logger.info("已创建 %d 个新目录", created_count)
            else:
                cls.logger.info("所有必要的目录已存在")
                
            return True
            
        except Exception as e:
            cls.logger.error("初始化目录结构失败: %s", LogHandler.format_error(e))
            return False 