from typing import List
from services.email_service import EmailService
from services.rule_processor import RuleProcessor

class EmailProcessor:
    def __init__(self):
        self.email_service = EmailService()
        self.rule_processor = RuleProcessor()

    def process_unread_emails(self) -> List[str]:
        """处理所有未读邮件"""
        try:
            # 获取所有未读邮件
            emails = self.email_service.get_unread_emails()
            if emails:
                print(f"找到 {len(emails)} 封未读邮件")
            else:
                print("没有未读邮件")
                return []

            all_downloaded_files = []
            for email_msg in emails:
                try:
                    print(f"\n处理邮件: {email_msg.subject}")
                    # 获取匹配的规则
                    matching_rule = self.rule_processor.get_matching_rule(email_msg)
                    if matching_rule:
                        print(f"邮件匹配规则: {matching_rule['name']}")
                        files = self.email_service.get_attachments(email_msg, matching_rule)
                        if files:  # 只有成功下载了附件才添加到列表
                            all_downloaded_files.extend(files)
                            print(f"已下载 {len(files)} 个附件")
                        else:
                            print("没有找到匹配的附件，邮件保持未读状态")
                    else:
                        print("邮件不匹配任何规则，保持未读状态")
                except Exception as e:
                    print(f"处理单个邮件时出错: {str(e)}")
                    continue

            if all_downloaded_files:
                print(f"\n总共下载了 {len(all_downloaded_files)} 个附件")
            else:
                print("\n没有下载任何附件")

            return all_downloaded_files
        except Exception as e:
            print(f"处理邮件时出错: {str(e)}")
            raise
        finally:
            self.email_service.disconnect() 