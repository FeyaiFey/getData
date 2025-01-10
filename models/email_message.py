from dataclasses import dataclass
from typing import Optional, Any

@dataclass
class EmailMessage:
    subject: str
    sender: str
    to: str
    uid: bytes
    _full_message: Optional[Any] = None

    @property
    def has_full_content(self) -> bool:
        """检查是否已加载完整邮件内容"""
        return self._full_message is not None 