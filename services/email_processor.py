from typing import List, Set
from services.email_service import EmailService
from services.rule_processor import RuleProcessor
from utils.log_handler import LogHandler
from models.email_message import EmailMessage

class EmailProcessor:
    """邮件处理器，负责处理未读邮件和下载附件"""
    
    def __init__(self):
        """初始化邮件处理器"""
        self.email_service = EmailService()
        self.rule_processor = RuleProcessor()
        self.logger = LogHandler().get_logger('EmailProcessor')
        self._processed_subjects: Set[str] = set()  # 用于跟踪已处理的邮件主题
        
    def _is_processed(self, subject: str) -> bool:
        """检查邮件是否已处理"""
        return subject in self._processed_subjects
        
    def _mark_as_processed(self, subject: str):
        """将邮件标记为已处理"""
        self._processed_subjects.add(subject)
        
    def _process_single_email(self, email_msg: EmailMessage) -> List[str]:
        """处理单个邮件
        
        Args:
            email_msg: 邮件对象
            
        Returns:
            List[str]: 下载的附件文件路径列表
        """
        try:
            self.logger.info("处理邮件: %s", email_msg.subject)
            
            # 检查是否已处理过
            if self._is_processed(email_msg.subject):
                self.logger.info("邮件已处理过，跳过: %s", email_msg.subject)
                return []
                
            # 标记为已处理
            self._mark_as_processed(email_msg.subject)
            
            # 获取匹配的规则
            matching_rule = self.rule_processor.get_matching_rule(email_msg)
            if not matching_rule:
                self.logger.info("邮件不匹配任何规则: %s", email_msg.subject)
                return []
                
            self.logger.info("邮件匹配规则: %s", matching_rule['name'])
            
            # 加载完整邮件内容
            if not self.email_service.load_full_message(email_msg):
                self.logger.error("加载邮件内容失败: %s", email_msg.subject)
                return []
            
            # 获取附件
            downloaded_files = self.email_service.get_attachments(email_msg, matching_rule)
            if not downloaded_files:
                self.logger.info("没有找到匹配的附件，邮件保持未读状态")
                return []
                
            # 标记为已读
            self.email_service.mark_as_read(email_msg)
            self.logger.info("成功处理邮件: %s", email_msg.subject)
            return downloaded_files
            
        except Exception as e:
            self.logger.error("处理邮件时出错: %s, 错误: %s", 
                            email_msg.subject, LogHandler.format_error(e))
            return []
            
    def process_unread_emails(self) -> None:
        """处理所有未读邮件"""
        try:
            # 获取未读邮件
            unread_emails = self.email_service.get_unread_emails()
            if not unread_emails:
                self.logger.info("没有未读邮件")
                return
                
            # 处理每封邮件
            total_files = []
            for email_msg in unread_emails:
                files = self._process_single_email(email_msg)
                total_files.extend(files)
                
            if not total_files:
                self.logger.info("没有下载任何附件")
            else:
                self.logger.info("成功下载 %d 个附件", len(total_files))
                
        except Exception as e:
            self.logger.error("处理未读邮件时出错: %s", LogHandler.format_error(e))
        finally:
            self.email_service.disconnect() 