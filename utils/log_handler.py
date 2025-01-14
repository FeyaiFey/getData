import logging
import os
from datetime import datetime
from logging.handlers import RotatingFileHandler

class LogHandler:
    """日志处理器，负责配置和管理日志"""
    
    def __init__(self):
        """初始化日志处理器"""
        # 确保日志目录存在
        self.log_dir = os.path.join(os.getcwd(), 'logs')
        os.makedirs(self.log_dir, exist_ok=True)
        
        # 设置日志格式
        self.formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
    def get_logger(self, name: str) -> logging.Logger:
        """
        获取指定名称的日志记录器
        
        Args:
            name: 日志记录器名称
            
        Returns:
            logging.Logger: 配置好的日志记录器
        """
        logger = logging.getLogger(name)
        
        # 如果已经配置过，直接返回
        if logger.handlers:
            return logger
            
        # 设置日志级别
        logger.setLevel(logging.DEBUG)
        
        # 创建并配置文件处理器
        today = datetime.now().strftime('%Y%m%d')
        log_file = os.path.join(self.log_dir, f'{today}.log')
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=10*1024*1024,  # 10MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(self.formatter)
        
        # 创建并配置控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(self.formatter)
        
        # 添加处理器到日志记录器
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger
        
    @staticmethod
    def format_error(e: Exception) -> str:
        """
        格式化异常信息
        
        Args:
            e: 异常对象
            
        Returns:
            str: 格式化后的异常信息
        """
        return f"{type(e).__name__}: {str(e)}" 