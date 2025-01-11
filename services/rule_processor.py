import re
import yaml
from typing import List, Dict, Any, Optional
from models.email_message import EmailMessage
from utils.log_handler import LogHandler

class RuleProcessor:
    def __init__(self, rules_file: str = 'email_rules.yaml'):
        """初始化规则处理器"""
        self.logger = LogHandler()
        try:
            self.rules = self._load_rules(rules_file)
            self.logger.info(f"成功初始化规则处理器，加载规则文件: {rules_file}")
        except Exception as e:
            self.logger.error(f"初始化规则处理器失败: {str(e)}")
            raise

    def _load_rules(self, rules_file: str) -> List[Dict[str, Any]]:
        """加载规则配置文件"""
        try:
            with open(rules_file, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            self.logger.info(f"成功加载规则配置文件: {rules_file}")
            return config['rules']
        except Exception as e:
            self.logger.error(f"加载规则配置文件失败: {str(e)}")
            raise

    def match_rule(self, email_msg: EmailMessage) -> bool:
        """检查邮件是否匹配任一规则"""
        return self.get_matching_rule(email_msg) is not None

    def get_matching_rule(self, email_msg: EmailMessage) -> Optional[Dict[str, Any]]:
        """获取匹配的规则配置"""
        for rule in self.rules:
            if self._check_single_rule(rule, email_msg):
                self.logger.debug(f"邮件匹配规则: {rule['name']}")
                return rule
        self.logger.debug(f"邮件不匹配任何规则: {email_msg.subject}")
        return None

    def _check_single_rule(self, rule: Dict[str, Any], email_msg: EmailMessage) -> bool:
        """检查邮件是否匹配单个规则"""
        # 检查主题（使用正则表达式匹配）
        if rule['subject_contains']:
            subject_matched = False
            for pattern in rule['subject_contains']:
                try:
                    if re.search(pattern, email_msg.subject, re.IGNORECASE):
                        subject_matched = True
                        self.logger.debug(f"主题匹配成功: {pattern}")
                        break
                except re.error:
                    self.logger.warning(f"无效的主题正则表达式: {pattern}")
                    continue
            if not subject_matched:
                self.logger.debug(f"主题不匹配: {email_msg.subject}")
                return False

        # 检查发件人
        if rule['sender_contains']:
            sender_email = str(email_msg.sender).lower()
            if not any(keyword.lower() in sender_email 
                      for keyword in rule['sender_contains']):
                self.logger.debug(f"发件人不匹配: {email_msg.sender}")
                return False
            self.logger.debug(f"发件人匹配成功: {email_msg.sender}")

        # 检查收件人
        if rule['receiver_contains']:
            receivers = str(email_msg.to).lower()
            if not any(keyword.lower() in receivers
                      for keyword in rule['receiver_contains']):
                self.logger.debug(f"收件人不匹配: {email_msg.to}")
                return False
            self.logger.debug(f"收件人匹配成功: {email_msg.to}")

        return True

    def match_attachment_name(self, filename: str) -> bool:
        """检查附件名称是否匹配任一规则的模式"""
        filename = filename.lower()
        for rule in self.rules:
            for pattern in rule['attachment_name_pattern']:
                try:
                    if re.match(pattern.lower(), filename):
                        self.logger.debug(f"附件名称匹配成功: {filename}")
                        return True
                except re.error:
                    self.logger.warning(f"无效的附件名正则表达式: {pattern}")
                    continue
        self.logger.debug(f"附件名称不匹配任何规则: {filename}")
        return False

    def get_rule_name(self, email_msg: EmailMessage) -> str:
        """获取匹配的规则名称，用于日志输出"""
        rule = self.get_matching_rule(email_msg)
        return rule['name'] if rule else "未知规则" 