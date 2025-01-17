import imaplib
import email
from typing import List, Optional, Dict, Any
from models.email_message import EmailMessage
from utils.email_decoder import EmailDecoder
from utils.file_handler import FileHandler
from utils.log_handler import LogHandler
from services.rule_processor import RuleProcessor
from config import (
    EMAIL_ADDRESS, EMAIL_PASSWORD, EMAIL_SERVER,
    EMAIL_SERVER_PORT, EMAIL_USE_SSL
)
import re
import os

class EmailService:
    """邮件服务类，负责邮件连接和操作
    
    主要功能：
    1. 连接和断开邮件服务器
    2. 获取未读邮件列表
    3. 下载和保存邮件附件
    4. 标记邮件已读状态
    5. 邮件内容解码和处理
    """

    def __init__(self, rule_processor: RuleProcessor):
        """初始化邮件服务
        
        初始化过程：
        1. 加载邮件配置
        2. 初始化日志记录器
        3. 初始化邮件解码器
        4. 设置规则处理器
        
        Args:
            rule_processor: 规则处理器实例
            
        Raises:
            Exception: 初始化失败时抛出
        """
        self.config = self._load_config()
        self.logger = LogHandler().get_logger('EmailService', file_level='DEBUG', console_level='INFO')
        self.decoder = EmailDecoder()
        self.rule_processor = rule_processor
        self._imap = None

    def _load_config(self) -> dict:
        """加载邮件配置
        
        从配置文件加载邮件服务器设置，包括服务器地址、端口、账号等。
        
        Returns:
            dict: 邮件配置字典
            
        Raises:
            FileNotFoundError: 配置文件不存在时抛出
        """
        self.email = EMAIL_ADDRESS
        self.password = EMAIL_PASSWORD
        self.server = EMAIL_SERVER
        self.port = EMAIL_SERVER_PORT
        self.use_ssl = EMAIL_USE_SSL
        
        return {
            'email': self.email,
            'password': self.password,
            'server': self.server,
            'port': self.port,
            'use_ssl': self.use_ssl
        }

    def connect(self) -> bool:
        """连接到邮件服务器
        
        连接过程：
        1. 创建IMAP4连接
        2. 登录邮件账号
        3. 选择收件箱
        
        Returns:
            bool: 连接成功返回True，否则返回False
            
        Raises:
            Exception: 连接过程中的错误
        """
        if self._imap is not None:
            return True

        try:
            if self.use_ssl:
                self._imap = imaplib.IMAP4_SSL(self.server, self.port)
            else:
                self._imap = imaplib.IMAP4(self.server, self.port)
            self._imap.login(self.email, self.password)
            self.logger.debug("已连接到邮箱服务器: %s", self.server)
            return True
        except Exception as e:
            self._imap = None
            self.logger.error("连接邮箱服务器失败: %s", LogHandler.format_error(e))
            raise

    def disconnect(self):
        """断开邮件服务器连接
        
        确保安全断开连接并清理资源。
        """
        if self._imap:
            try:
                self._imap.close()
                self._imap.logout()
                self.logger.debug("已断开邮箱连接")
            except Exception as e:
                self.logger.error("断开连接失败: %s", LogHandler.format_error(e))
            finally:
                self._imap = None

    def get_unread_emails(self) -> List[EmailMessage]:
        """获取所有未读邮件
        
        Returns:
            List[EmailMessage]: 未读邮件列表
            
        Raises:
            Exception: 获取邮件失败时抛出
        """
        try:
            if not self.connect():
                raise Exception("无法连接到邮箱服务器")
                
            self._imap.select('INBOX')
            _, messages = self._imap.search(None, 'UNSEEN')
            
            email_list = []
            for num in messages[0].split():
                email_msg = self._fetch_email_header(num)
                if email_msg:
                    email_list.append(email_msg)
            
            if email_list:
                self.logger.info("找到 %d 封未读邮件", len(email_list))
            return email_list
        except Exception as e:
            self.logger.error("获取未读邮件失败: %s", LogHandler.format_error(e))
            raise  # 重新抛出异常，保持原有行为

    def _fetch_email_header(self, uid: bytes) -> Optional[EmailMessage]:
        """获取邮件头信息"""
        try:
            _, msg_data = self._imap.fetch(uid, '(BODY.PEEK[HEADER])')
            email_header = email.message_from_bytes(msg_data[0][1])
            
            email_msg = EmailMessage(
                subject=EmailDecoder.decode_str(email_header['subject']),
                sender=email_header['from'],
                to=email_header['to'],
                uid=uid
            )
            self.logger.debug("获取邮件头信息: %s", email_msg.subject)
            return email_msg
        except Exception as e:
            self.logger.error("获取邮件头信息时出错: %s", LogHandler.format_error(e))
            return None

    def load_full_message(self, email_msg: EmailMessage) -> bool:
        """加载完整的邮件内容
        
        Args:
            email_msg: 邮件对象
            
        Returns:
            bool: 加载成功返回True，否则返回False
        """
        if email_msg.has_full_content:
            return True

        try:
            # 尝试不同的邮件获取命令
            msg_data = None
            fetch_methods = [
                ('RFC822', '(RFC822)'),
                ('BODY[]', '(BODY[])'),
                ('BODY.PEEK[]', '(BODY.PEEK[])')
            ]
            
            for method_name, method_cmd in fetch_methods:
                try:
                    self.logger.debug("尝试使用 %s 获取邮件", method_name)
                    _, msg_data = self._imap.fetch(email_msg.uid, method_cmd)
                    if msg_data and msg_data[0]:
                        break
                except Exception as e:
                    self.logger.debug("%s 获取失败: %s", method_name, LogHandler.format_error(e))
                    continue

            if not msg_data or not msg_data[0]:
                self.logger.error("获取邮件数据为空: %s", email_msg.subject)
                return False

            raw_email = msg_data[0][1] if isinstance(msg_data[0], tuple) else msg_data[0]
            
            # 处理数据类型
            if isinstance(raw_email, int):
                raw_email = str(raw_email).encode('utf-8')
            elif not isinstance(raw_email, bytes):
                raw_email = str(raw_email).encode('utf-8')

            full_message = email.message_from_bytes(raw_email)
            email_msg.set_full_message(full_message)
            self.logger.debug("已加载完整邮件内容: %s", email_msg.subject)
            return True
        except Exception as e:
            self.logger.error("加载邮件内容时出错: %s", LogHandler.format_error(e))
            return False

    def mark_as_read(self, email_msg: EmailMessage):
        """标记邮件为已读
        
        Args:
            email_msg: 邮件对象
            
        Raises:
            Exception: 标记失败时抛出
        """
        try:
            self._imap.store(email_msg.uid, '+FLAGS', '\\Seen')
            self.logger.info("邮件已标记为已读: %s", email_msg.subject)
        except Exception as e:
            self.logger.error("标记邮件为已读时出错: %s", LogHandler.format_error(e))

    def _debug_print_message_structure(self, message, level=0):
        """打印邮件结构的辅助方法"""
        prefix = "  " * level
        self.logger.debug("{}Type: {}", prefix, message.get_content_type())
        self.logger.debug("{}Disposition: {}", prefix, message.get('Content-Disposition'))
        
        filename = message.get_filename()
        if filename:
            decoded_filename = EmailDecoder.decode_filename(filename)
            self.logger.debug("{}Filename: {}", prefix, filename)
            self.logger.debug("{}Decoded filename: {}", prefix, decoded_filename)
        
        if message.is_multipart():
            for part in message.get_payload():
                self._debug_print_message_structure(part, level + 1)
        else:
            content_type = message.get_content_type()
            self.logger.debug("{}Content-Type: {}", prefix, content_type)
            if content_type == 'text/plain':
                try:
                    self.logger.debug("{}Text content length: {}", prefix, len(message.get_payload()))
                except:
                    self.logger.debug("{}Could not get text content length", prefix)

    def get_attachments(self, email_msg: EmailMessage, rule: Dict[str, Any]) -> List[str]:
        """获取匹配规则的附件
        
        Args:
            email_msg: 邮件对象
            rule: 匹配规则
            
        Returns:
            List[str]: 下载的附件文件路径列表
            
        Raises:
            Exception: 下载附件失败时抛出
        """
        try:
            downloaded_files = []
            
            # 遍历邮件结构
            for part in email_msg.message.walk():
                if not part.get('Content-Disposition'):
                    continue
                    
                filename = part.get_filename()
                if not filename:
                    continue
                    
                # 解码文件名
                filename = EmailDecoder.decode_filename(filename)
                
                # 检查是否为Excel文件
                if not filename.lower().endswith(('.xls', '.xlsx')):
                    self.logger.debug("跳过非Excel文件: %s", filename)
                    continue
                    
                # 检查文件名是否匹配规则
                if not self.rule_processor.match_attachment_name(rule, filename):
                    self.logger.debug("附件名称不匹配规则: %s", filename)
                    continue
                    
                # 获取保存路径
                save_dir = rule['download_path']
                save_path = os.path.join(save_dir, filename)
                
                # 确保目录存在
                FileHandler.ensure_dir(save_dir)
                
                # 保存附件
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        FileHandler.save_file(payload, save_path)
                        downloaded_files.append(save_path)
                        self.logger.info("已保存附件: %s", filename)
                    else:
                        self.logger.warning("附件内容为空: %s", filename)
                except Exception as e:
                    self.logger.error("保存附件失败 [%s]: %s", 
                                    filename, LogHandler.format_error(e))
                    
            return downloaded_files
            
        except Exception as e:
            self.logger.error("处理附件失败 [%s]: %s", 
                            email_msg.subject, LogHandler.format_error(e))
            return [] 