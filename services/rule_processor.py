import re
import yaml
from typing import List, Dict, Any, Optional
from models.email_message import EmailMessage

class RuleProcessor:
    def __init__(self, rules_file: str = 'email_rules.yaml'):
        """初始化规则处理器"""
        self.rules = self._load_rules(rules_file)

    def _load_rules(self, rules_file: str) -> List[Dict[str, Any]]:
        """加载规则配置文件"""
        with open(rules_file, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config['rules']

    def match_rule(self, email_msg: EmailMessage) -> bool:
        """检查邮件是否匹配任一规则"""
        return self.get_matching_rule(email_msg) is not None

    def get_matching_rule(self, email_msg: EmailMessage) -> Optional[Dict[str, Any]]:
        """获取匹配的规则配置"""
        for rule in self.rules:
            if self._check_single_rule(rule, email_msg):
                return rule
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
                        break
                except re.error:
                    print(f"警告: 无效的主题正则表达式: {pattern}")
                    continue
            if not subject_matched:
                return False

        # 检查发件人
        if rule['sender_contains']:
            sender_email = str(email_msg.sender).lower()
            if not any(keyword.lower() in sender_email 
                      for keyword in rule['sender_contains']):
                return False

        # 检查收件人
        if rule['receiver_contains']:
            receivers = str(email_msg.to).lower()
            if not any(keyword.lower() in receivers
                      for keyword in rule['receiver_contains']):
                return False

        return True

    def match_attachment_name(self, filename: str) -> bool:
        """检查附件名称是否匹配任一规则的模式"""
        filename = filename.lower()
        for rule in self.rules:
            for pattern in rule['attachment_name_pattern']:
                try:
                    if re.match(pattern.lower(), filename):
                        return True
                except re.error:
                    print(f"无效的附件名正则表达式: {pattern}")
                    continue
        return False

    def get_rule_name(self, email_msg: EmailMessage) -> str:
        """获取匹配的规则名称，用于日志输出"""
        rule = self.get_matching_rule(email_msg)
        return rule['name'] if rule else "未知规则" 