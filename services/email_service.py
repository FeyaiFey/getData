import imaplib
import email
from typing import List, Optional
from models.email_message import EmailMessage
from utils.email_decoder import EmailDecoder
from utils.file_handler import FileHandler
from services.rule_processor import RuleProcessor
from config import (
    EMAIL_ADDRESS, EMAIL_PASSWORD, EMAIL_SERVER,
    EMAIL_SERVER_PORT, EMAIL_USE_SSL
)
import re

class EmailService:
    def __init__(self):
        self.email = EMAIL_ADDRESS
        self.password = EMAIL_PASSWORD
        self.server = EMAIL_SERVER
        self.port = EMAIL_SERVER_PORT
        self.use_ssl = EMAIL_USE_SSL
        self.imap = None
        self.rule_processor = RuleProcessor()

    def connect(self):
        """连接到邮箱服务器"""
        if self.imap is not None:
            return

        if self.use_ssl:
            self.imap = imaplib.IMAP4_SSL(self.server, self.port)
        else:
            self.imap = imaplib.IMAP4(self.server, self.port)
        self.imap.login(self.email, self.password)

    def disconnect(self):
        """断开连接"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
            except:
                pass
            finally:
                self.imap = None

    def get_unread_emails(self) -> List[EmailMessage]:
        """获取所有未读邮件"""
        try:
            self.connect()
            self.imap.select('INBOX')
            _, messages = self.imap.search(None, 'UNSEEN')
            
            email_list = []
            for num in messages[0].split():
                email_msg = self._fetch_email_header(num)
                if email_msg:
                    email_list.append(email_msg)
            
            return email_list
        except Exception as e:
            print(f"获取未读邮件时出错: {str(e)}")
            return []

    def _fetch_email_header(self, uid: bytes) -> Optional[EmailMessage]:
        """获取邮件头信息"""
        try:
            _, msg_data = self.imap.fetch(uid, '(BODY.PEEK[HEADER])')
            email_header = email.message_from_bytes(msg_data[0][1])
            
            return EmailMessage(
                subject=EmailDecoder.decode_str(email_header['subject']),
                sender=email_header['from'],
                to=email_header['to'],
                uid=uid
            )
        except Exception as e:
            print(f"获取邮件头信息时出错: {str(e)}")
            return None

    def load_full_message(self, email_msg: EmailMessage) -> bool:
        """加载完整邮件内容"""
        if email_msg.has_full_content:
            return True

        try:
            _, msg_data = self.imap.fetch(email_msg.uid, '(BODY.PEEK[])')
            email_msg._full_message = email.message_from_bytes(msg_data[0][1])
            return True
        except Exception as e:
            print(f"加载邮件内容时出错: {str(e)}")
            return False

    def mark_as_read(self, email_msg: EmailMessage):
        """将邮件标记为已读"""
        try:
            self.imap.store(email_msg.uid, '+FLAGS', '\\Seen')
            print(f"邮件已标记为已读: {email_msg.subject}")
        except Exception as e:
            print(f"标记邮件为已读时出错: {str(e)}")

    def get_attachments(self, email_msg: EmailMessage, rule: dict) -> List[str]:
        """获取邮件附件"""
        if not self.load_full_message(email_msg):
            print(f"无法加载邮件完整内容: {email_msg.subject}")
            return []

        downloaded_files = []
        has_matching_attachment = False
        print(f"\n开始处理邮件附件 - 主题: {email_msg.subject}")

        for part in email_msg._full_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            
            content_disp = part.get('Content-Disposition')
            if content_disp is None:
                continue

            filename = EmailDecoder.decode_filename(part.get_filename())
            if not filename:
                continue

            print(f"发现附件: {filename}")

            if filename.lower().endswith(('.xlsx', '.xls')):
                print(f"Excel附件: {filename}")
                # 检查文件名是否匹配当前规则的模式
                filename_matched = False
                for pattern in rule['attachment_name_pattern']:
                    try:
                        print(f"尝试匹配模式: {pattern}")
                        if re.match(pattern, filename, re.IGNORECASE):
                            filename_matched = True
                            print(f"匹配成功: {pattern}")
                            break
                        else:
                            print(f"匹配失败: {pattern}")
                    except re.error as e:
                        print(f"警告: 无效的附件名正则表达式: {pattern}, 错误: {str(e)}")
                        continue
                
                if not filename_matched:
                    print(f"附件名称不匹配任何模式: {filename}")
                    continue

                has_matching_attachment = True
                download_path = rule.get('download_path', 'downloads')
                print(f"下载路径: {download_path}")
                FileHandler.ensure_dir(download_path)

                save_filename = FileHandler.generate_unique_filename(filename)
                file_path = FileHandler.join_paths(download_path, save_filename)
                print(f"保存路径: {file_path}")

                try:
                    content = part.get_payload(decode=True)
                    if content:
                        FileHandler.save_file(content, file_path)
                        downloaded_files.append(file_path)
                        print(f"成功下载附件: {save_filename} (规则: {rule['name']})")
                    else:
                        print(f"警告: 附件内容为空: {filename}")
                except Exception as e:
                    print(f"下载附件时出错: {filename}, 错误: {str(e)}")

        if has_matching_attachment:
            self.mark_as_read(email_msg)
            print(f"邮件已标记为已读: {email_msg.subject}")
        else:
            print(f"未找到匹配的附件: {email_msg.subject}")

        return downloaded_files 