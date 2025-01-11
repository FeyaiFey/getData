import re
import yaml
from typing import List, Dict, Any, Optional, Pattern
from models.email_message import EmailMessage
from utils.log_handler import LogHandler

class RuleProcessor:
    """规则处理器，负责加载和匹配邮件规则"""

    def __init__(self, rules_file: str = 'email_rules.yaml'):
        """初始化规则处理器
        
        Args:
            rules_file: 规则配置文件路径
        """
        self.logger = LogHandler().get_logger('RuleProcessor')
        self.rules_file = rules_file
        self.rules: List[Dict[str, Any]] = []
        self._compiled_patterns: Dict[str, List[Pattern]] = {}
        
        try:
            self.rules = self._load_rules(rules_file)
            self._compile_patterns()
            self.logger.info("成功初始化规则处理器，加载规则文件: %s", rules_file)
        except Exception as e:
            self.logger.error("初始化规则处理器失败: %s, 错误: %s", 
                            rules_file, LogHandler.format_error(e))
            raise

    def _load_rules(self, rules_file: str) -> List[Dict[str, Any]]:
        """加载规则配置文件
        
        Args:
            rules_file: 规则配置文件路径
            
        Returns:
            List[Dict[str, Any]]: 规则配置列表
        """
        try:
            with open(rules_file, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            self.logger.info("成功加载规则配置文件: %s", rules_file)
            return config['rules']
        except Exception as e:
            self.logger.error("加载规则配置文件失败: %s, 错误: %s", 
                            rules_file, LogHandler.format_error(e))
            raise

    def _compile_patterns(self):
        """预编译所有正则表达式模式"""
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
            rule: 规则配置
            subject: 邮件主题
            
        Returns:
            bool: 是否匹配
        """
        patterns = self._compiled_patterns[rule['name']]['subject']
        for pattern in patterns:
            try:
                if pattern.search(subject):
                    self.logger.debug("主题匹配成功: %s", pattern.pattern)
                    return True
            except Exception as e:
                self.logger.warning("主题匹配出错: %s, 错误: %s", 
                                  pattern.pattern, LogHandler.format_error(e))
        self.logger.debug("主题不匹配: %s", subject)
        return False

    def _check_sender(self, rule: Dict[str, Any], sender: str) -> bool:
        """检查发件人是否匹配规则
        
        Args:
            rule: 规则配置
            sender: 发件人地址
            
        Returns:
            bool: 是否匹配
        """
        if not rule['sender_contains']:
            return True
            
        sender = str(sender).lower()
        for keyword in rule['sender_contains']:
            if keyword.lower() in sender:
                self.logger.debug("发件人匹配成功: %s", sender)
                return True
        self.logger.debug("发件人不匹配: %s", sender)
        return False

    def _check_receiver(self, rule: Dict[str, Any], receiver: str) -> bool:
        """检查收件人是否匹配规则
        
        Args:
            rule: 规则配置
            receiver: 收件人地址
            
        Returns:
            bool: 是否匹配
        """
        if not rule['receiver_contains']:
            return True
            
        receiver = str(receiver).lower()
        for keyword in rule['receiver_contains']:
            if keyword.lower() in receiver:
                self.logger.debug("收件人匹配成功: %s", receiver)
                return True
        self.logger.debug("收件人不匹配: %s", receiver)
        return False

    def match_rule(self, email_msg: EmailMessage) -> bool:
        """检查邮件是否匹配任一规则
        
        Args:
            email_msg: 邮件对象
            
        Returns:
            bool: 是否匹配任一规则
        """
        return self.get_matching_rule(email_msg) is not None

    def get_matching_rule(self, email_msg: EmailMessage) -> Optional[Dict[str, Any]]:
        """获取匹配的规则配置
        
        Args:
            email_msg: 邮件对象
            
        Returns:
            Optional[Dict[str, Any]]: 匹配的规则配置，如果没有匹配则返回 None
        """
        for rule in self.rules:
            if self._check_single_rule(rule, email_msg):
                self.logger.debug("邮件匹配规则: %s", rule['name'])
                return rule
        self.logger.debug("邮件不匹配任何规则: %s", email_msg.subject)
        return None

    def _check_single_rule(self, rule: Dict[str, Any], email_msg: EmailMessage) -> bool:
        """检查邮件是否匹配单个规则
        
        Args:
            rule: 规则配置
            email_msg: 邮件对象
            
        Returns:
            bool: 是否匹配规则
        """
        # 按顺序检查各个条件，如果任一条件不满足则立即返回 False
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
            rule: 规则配置
            filename: 附件文件名
            
        Returns:
            bool: 是否匹配规则
        """
        filename = filename.lower()
        patterns = self._compiled_patterns[rule['name']]['attachment']
        
        for pattern in patterns:
            try:
                if pattern.search(filename):
                    self.logger.debug("附件名称匹配成功: %s", filename)
                    return True
            except Exception as e:
                self.logger.warning("附件名称匹配出错: %s, 错误: %s", 
                                  pattern.pattern, LogHandler.format_error(e))
                
        self.logger.debug("附件名称不匹配: %s", filename)
        return False

    def get_rule_name(self, email_msg: EmailMessage) -> str:
        """获取匹配的规则名称，用于日志输出"""
        rule = self.get_matching_rule(email_msg)
        return rule['name'] if rule else "未知规则" 