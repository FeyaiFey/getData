import logging
import os
from datetime import datetime

class LogHandler:
    _instance = None
    _initialized = False

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(LogHandler, cls).__new__(cls)
        return cls._instance

    def __init__(self):
        if not LogHandler._initialized:
            self._setup_logger()
            LogHandler._initialized = True

    def _setup_logger(self):
        """设置日志配置"""
        # 创建logs目录
        log_dir = 'logs'
        os.makedirs(log_dir, exist_ok=True)

        # 生成日志文件名
        current_date = datetime.now().strftime('%Y%m%d')
        log_file = os.path.join(log_dir, f'email_downloader_{current_date}.log')

        # 配置根日志记录器
        self.logger = logging.getLogger('EmailDownloader')
        self.logger.setLevel(logging.DEBUG)

        # 文件处理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)

        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        # 日志格式
        formatter = logging.Formatter(
            '%(asctime)s [%(levelname)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        # 添加处理器
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

    def info(self, message):
        """记录信息日志"""
        self.logger.info(message)

    def error(self, message):
        """记录错误日志"""
        self.logger.error(message)

    def warning(self, message):
        """记录警告日志"""
        self.logger.warning(message)

    def debug(self, message):
        """记录调试日志"""
        self.logger.debug(message)

    def critical(self, message):
        """记录严重错误日志"""
        self.logger.critical(message) 