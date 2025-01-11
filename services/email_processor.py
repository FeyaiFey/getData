from typing import List
from services.email_service import EmailService
from services.rule_processor import RuleProcessor
from utils.log_handler import LogHandler

class EmailProcessor:
    def __init__(self):
        self.email_service = EmailService()
        self.rule_processor = RuleProcessor()
        self.logger = LogHandler()

    def process_unread_emails(self) -> List[str]:
        """处理所有未读邮件"""
        try:
            # 获取所有未读邮件
            emails = self.email_service.get_unread_emails()
            if not emails:
                self.logger.info("没有未读邮件")
                return []

            # 记录找到的未读邮件数量
            self.logger.info(f"找到 {len(emails)} 封未读邮件")
            
            all_downloaded_files = []
            processed_subjects = set()  # 用于跟踪已处理的邮件主题

            for email_msg in emails:
                # 检查是否已处理过相同主题的邮件
                if email_msg.subject in processed_subjects:
                    self.logger.debug(f"跳过重复邮件: {email_msg.subject}")
                    continue
                
                processed_subjects.add(email_msg.subject)
                
                try:
                    self.logger.info(f"处理邮件: {email_msg.subject}")
                    # 获取匹配的规则
                    matching_rule = self.rule_processor.get_matching_rule(email_msg)
                    if matching_rule:
                        self.logger.info(f"邮件匹配规则: {matching_rule['name']}")
                        files = self.email_service.get_attachments(email_msg, matching_rule)
                        if files:  # 只有成功下载了附件才添加到列表
                            all_downloaded_files.extend(files)
                            self.logger.info(f"已下载 {len(files)} 个附件")
                        else:
                            self.logger.info("没有找到匹配的附件，邮件保持未读状态")
                    else:
                        self.logger.info("邮件不匹配任何规则，保持未读状态")
                except Exception as e:
                    self.logger.error(f"处理单个邮件时出错: {str(e)}")
                    continue

            if all_downloaded_files:
                self.logger.info(f"总共下载了 {len(all_downloaded_files)} 个附件")
            else:
                self.logger.info("没有下载任何附件")

            return all_downloaded_files
        except Exception as e:
            self.logger.error(f"处理邮件时出错: {str(e)}")
            raise
        finally:
            self.email_service.disconnect() 