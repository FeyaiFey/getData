from email.message import Message
from typing import Optional

class EmailMessage:
    """邮件消息类，用于存储邮件信息"""
    
    def __init__(self, subject: str, sender: str, to: str, uid: bytes):
        """初始化邮件消息
        
        Args:
            subject: 邮件主题
            sender: 发件人
            to: 收件人
            uid: 邮件唯一标识
        """
        self.subject = subject
        self.sender = sender
        self.to = to
        self.uid = uid
        self._full_message: Optional[Message] = None
        
    @property
    def has_full_content(self) -> bool:
        """是否已加载完整邮件内容"""
        return self._full_message is not None
        
    @property
    def message(self) -> Optional[Message]:
        """获取完整邮件内容"""
        return self._full_message
        
    def set_full_message(self, message: Message):
        """设置完整邮件内容
        
        Args:
            message: 完整邮件内容
        """
        self._full_message = message 