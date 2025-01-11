import logging
import os
from datetime import datetime
from typing import Optional
import json

class LogHandler:
    """日志处理工具类，用于统一管理日志"""
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not hasattr(self, 'initialized'):
            self.initialized = True
            self._setup_logger()
    
    def _setup_logger(self):
        """设置日志配置"""
        # 创建日志目录
        log_dir = 'logs'
        os.makedirs(log_dir, exist_ok=True)
        
        # 设置日志格式
        log_format = '%(asctime)s [%(levelname)s] [%(name)s] %(message)s'
        date_format = '%Y-%m-%d %H:%M:%S'
        formatter = logging.Formatter(log_format, date_format)
        
        # 设置控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.INFO)
        
        # 设置文件处理器
        log_file = os.path.join(log_dir, f'app_{datetime.now().strftime("%Y%m%d")}.log')
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
        
        # 配置根日志记录器
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(console_handler)
        root_logger.addHandler(file_handler)
    
    def get_logger(self, name: str) -> logging.Logger:
        """获取指定名称的日志记录器
        
        Args:
            name: 日志记录器名称
            
        Returns:
            logging.Logger: 日志记录器实例
        """
        return logging.getLogger(name)
    
    @staticmethod
    def format_error(e: Exception) -> str:
        """格式化异常信息
        
        Args:
            e: 异常对象
            
        Returns:
            str: 格式化后的异常信息
        """
        return f"{type(e).__name__}: {str(e)}"
        
    def _load_process_dates(self):
        """加载处理日期记录"""
        try:
            if os.path.exists('process_dates.json'):
                with open('process_dates.json', 'r', encoding='utf-8') as f:
                    dates = json.load(f)
                    self.process_dates = {k: datetime.strptime(v, '%Y-%m-%d') 
                                       for k, v in dates.items()}
        except Exception as e:
            self.get_logger('LogHandler').error(
                "加载处理日期记录失败: %s", self.format_error(e))
            self.process_dates = {}
    
    def _save_process_dates(self):
        """保存处理日期记录"""
        try:
            dates = {k: v.strftime('%Y-%m-%d') 
                    for k, v in self.process_dates.items()}
            with open('process_dates.json', 'w', encoding='utf-8') as f:
                json.dump(dates, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.get_logger('LogHandler').error(
                "保存处理日期记录失败: %s", self.format_error(e))
    
    def get_last_process_date(self, rule_name: str) -> Optional[datetime]:
        """获取上次处理日期
        
        Args:
            rule_name: 规则名称
            
        Returns:
            Optional[datetime]: 上次处理日期，如果没有则返回None
        """
        return self.process_dates.get(rule_name)
    
    def update_process_date(self, rule_name: str, date: datetime):
        """更新处理日期
        
        Args:
            rule_name: 规则名称
            date: 处理日期
        """
        self.process_dates[rule_name] = date
        self._save_process_dates()
    
    def should_process_date(self, rule_name: str, date: datetime) -> bool:
        """检查是否应该处理该日期
        
        Args:
            rule_name: 规则名称
            date: 待检查的日期
            
        Returns:
            bool: 是否应该处理
        """
        last_date = self.get_last_process_date(rule_name)
        if not last_date:
            return True
        return date > last_date 