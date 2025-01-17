import logging
import os
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler

class LogHandler:
    """日志处理器，负责配置和管理日志"""
    
    # 日志级别映射
    LOG_LEVELS = {
        'DEBUG': logging.DEBUG,
        'INFO': logging.INFO,
        'WARNING': logging.WARNING,
        'ERROR': logging.ERROR
    }
    
    def __init__(self):
        """初始化日志处理器"""
        # 确保日志目录存在
        self.log_dir = os.path.join(os.getcwd(), 'logs')
        os.makedirs(self.log_dir, exist_ok=True)
        
        # 设置日志格式 - 更简洁的格式
        self.formatter = logging.Formatter(
            '[%(asctime)s] %(levelname)-8s [%(name)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 清理旧日志文件
        self._cleanup_old_logs()
        
    def get_logger(self, name: str, file_level: str = 'DEBUG', console_level: str = 'INFO') -> logging.Logger:
        """
        获取指定名称的日志记录器
        
        Args:
            name: 日志记录器名称
            file_level: 文件日志级别
            console_level: 控制台日志级别
            
        Returns:
            logging.Logger: 配置好的日志记录器
        """
        logger = logging.getLogger(name)
        
        # 如果已经配置过，直接返回
        if logger.handlers:
            return logger
            
        # 设置日志级别为最低级别
        logger.setLevel(logging.DEBUG)
        
        # 创建并配置文件处理器
        today = datetime.now().strftime('%Y%m%d')
        log_file = os.path.join(self.log_dir, f'{today}.log')
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=5*1024*1024,  # 5MB
            backupCount=3,
            encoding='utf-8'
        )
        file_handler.setLevel(self.LOG_LEVELS.get(file_level.upper(), logging.DEBUG))
        file_handler.setFormatter(self.formatter)
        
        # 创建并配置控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(self.LOG_LEVELS.get(console_level.upper(), logging.INFO))
        console_handler.setFormatter(self.formatter)
        
        # 添加处理器到日志记录器
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger
        
    def _cleanup_old_logs(self, days: int = 7):
        """
        清理指定天数之前的日志文件
        
        Args:
            days: 保留的天数，默认7天
        """
        try:
            cutoff_date = datetime.now() - timedelta(days=days)
            for filename in os.listdir(self.log_dir):
                if not filename.endswith('.log'):
                    continue
                    
                file_path = os.path.join(self.log_dir, filename)
                file_date = datetime.fromtimestamp(os.path.getctime(file_path))
                
                if file_date < cutoff_date:
                    os.remove(file_path)
        except Exception as e:
            print(f"清理日志文件失败: {self.format_error(e)}")
        
    @staticmethod
    def format_error(e: Exception) -> str:
        """
        格式化异常信息，包含异常类名和详细信息
        
        Args:
            e: 异常对象
            
        Returns:
            str: 格式化后的异常信息
        """
        return f"{type(e).__name__}: {str(e)}" 