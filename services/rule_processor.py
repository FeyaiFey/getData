import re
import yaml
from typing import List, Dict, Any, Optional, Pattern
from models.email_message import EmailMessage
from utils.log_handler import LogHandler
import os

class RuleProcessor:
    """规则处理器，负责加载和匹配邮件规则
    
    主要功能：
    1. 加载邮件规则配置文件
    2. 编译和管理正则表达式模式
    3. 匹配邮件主题、发件人、收件人
    4. 匹配附件名称
    """

    def __init__(self):
        """初始化规则处理器
        
        初始化过程：
        1. 加载规则配置文件
        2. 初始化日志记录器
        3. 编译正则表达式模式
        
        Raises:
            FileNotFoundError: 规则配置文件不存在时抛出
            Exception: 初始化过程中的其他错误
        """
        self.rules = self._load_rules()
        self.logger = LogHandler().get_logger('RuleProcessor', file_level='DEBUG', console_level='INFO')
        self._compiled_patterns: Dict[str, List[Pattern]] = {}
        
        try:
            self._compile_patterns()
            self.logger.debug("已加载规则文件: %s", "config/email_rules.yaml")
        except Exception as e:
            self.logger.error("初始化规则处理器失败: %s", LogHandler.format_error(e))
            raise

    def _load_rules(self) -> dict:
        """加载邮件规则配置
        
        从配置文件加载邮件处理规则，包括主题匹配、发件人匹配、收件人匹配等规则。
        
        Returns:
            dict: 规则配置字典，包含所有处理规则
            
        Raises:
            FileNotFoundError: 配置文件不存在时抛出
        """
        try:
            config_path = os.path.join("config", "email_rules.yaml")
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                return config.get('rules', {})
        except Exception as e:
            self.logger.error("加载规则配置失败: %s", LogHandler.format_error(e))
            return {}

    def _compile_patterns(self):
        """预编译所有正则表达式模式
        
        将配置文件中的正则表达式模式预先编译，提高匹配效率。
        包括主题匹配模式和附件名称匹配模式。
        """
        for rule in self.rules:
            rule_name = rule['name']
            self._compiled_patterns[rule_name] = {
                'subject': [self._compile_pattern(p) for p in rule['subject_contains']],
                'attachment': [self._compile_pattern(p) for p in rule['attachment_name_pattern']]
            }

    def _compile_pattern(self, pattern: str) -> Pattern:
        """编译单个正则表达式模式
        
        Args:
            pattern: 正则表达式字符串
            
        Returns:
            Pattern: 编译后的正则表达式对象
            
        Raises:
            re.error: 正则表达式编译失败时抛出
        """
        try:
            return re.compile(pattern, re.IGNORECASE)
        except re.error as e:
            self.logger.error("编译正则表达式失败: %s, 错误: %s", 
                            pattern, LogHandler.format_error(e))
            raise

    def _check_subject(self, rule: Dict[str, Any], subject: str) -> bool:
        """检查邮件主题是否匹配规则
        
        Args:
            rule: 规则配置字典
            subject: 邮件主题
            
        Returns:
            bool: 如果主题匹配任一规则返回True，否则返回False
        """
        patterns = self._compiled_patterns[rule['name']]['subject']
        for pattern in patterns:
            try:
                if pattern.search(subject):
                    self.logger.debug("主题 [%s] 匹配规则: %s", subject, pattern.pattern)
                    return True
            except Exception as e:
                self.logger.warning("主题匹配异常 [%s]: %s", pattern.pattern, LogHandler.format_error(e))
        return False

    def _check_sender(self, rule: Dict[str, Any], sender: str) -> bool:
        """检查发件人是否匹配规则
        
        Args:
            rule: 规则配置字典
            sender: 发件人地址
            
        Returns:
            bool: 如果发件人匹配任一规则返回True，否则返回False
        """
        if not rule['sender_contains']:
            return True
            
        sender = str(sender).lower()
        for keyword in rule['sender_contains']:
            if keyword.lower() in sender:
                self.logger.debug("发件人 [%s] 匹配关键词: %s", sender, keyword)
                return True
        return False

    def _check_receiver(self, rule: Dict[str, Any], receiver: str) -> bool:
        """检查收件人是否匹配规则
        
        Args:
            rule: 规则配置字典
            receiver: 收件人地址
            
        Returns:
            bool: 如果收件人匹配任一规则返回True，否则返回False
        """
        if not rule['receiver_contains']:
            return True
            
        receiver = str(receiver).lower()
        for keyword in rule['receiver_contains']:
            if keyword.lower() in receiver:
                self.logger.debug("收件人 [%s] 匹配关键词: %s", receiver, keyword)
                return True
        return False

    def get_matching_rule(self, email_msg: EmailMessage) -> Optional[Dict[str, Any]]:
        """获取匹配的规则配置
        
        检查邮件是否匹配任一规则，返回第一个匹配的规则。
        
        Args:
            email_msg: 邮件对象
            
        Returns:
            Optional[Dict[str, Any]]: 匹配的规则配置，如果没有匹配则返回None
        """
        for rule in self.rules:
            if self._check_single_rule(rule, email_msg):
                self.logger.debug("邮件 [%s] 匹配规则: %s", email_msg.subject, rule['name'])
                return rule
        self.logger.debug("邮件 [%s] 不匹配任何规则", email_msg.subject)
        return None

    def _check_single_rule(self, rule: Dict[str, Any], email_msg: EmailMessage) -> bool:
        """检查邮件是否匹配单个规则
        
        按顺序检查主题、发件人、收件人是否都匹配规则。
        
        Args:
            rule: 规则配置字典
            email_msg: 邮件对象
            
        Returns:
            bool: 如果所有条件都匹配返回True，否则返回False
        """
        if not self._check_subject(rule, email_msg.subject):
            return False
            
        if not self._check_sender(rule, email_msg.sender):
            return False
            
        if not self._check_receiver(rule, email_msg.to):
            return False
            
        return True

    def match_attachment_name(self, rule: Dict[str, Any], filename: str) -> bool:
        """检查附件名称是否匹配规则的模式
        
        Args:
            rule: 规则配置字典
            filename: 附件文件名
            
        Returns:
            bool: 如果文件名匹配任一模式返回True，否则返回False
        """
        filename = filename.lower()
        patterns = self._compiled_patterns[rule['name']]['attachment']
        
        for pattern in patterns:
            try:
                if pattern.search(filename):
                    self.logger.debug("附件 [%s] 匹配规则: %s", filename, pattern.pattern)
                    return True
            except Exception as e:
                self.logger.warning("附件名称匹配异常 [%s]: %s", pattern.pattern, LogHandler.format_error(e))
                
        return False

    def get_rule_name(self, email_msg: EmailMessage) -> str:
        """获取匹配的规则名称
        
        Args:
            email_msg: 邮件对象
            
        Returns:
            str: 匹配的规则名称，如果没有匹配则返回"未知规则"
        """
        rule = self.get_matching_rule(email_msg)
        return rule['name'] if rule else "未知规则" 