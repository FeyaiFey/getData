import email
from email.header import decode_header
from typing import Optional, Union
from utils.log_handler import LogHandler

class EmailDecoder:
    """邮件解码工具类，用于处理邮件编码相关的问题"""
    
    logger = LogHandler().get_logger('EmailDecoder')
    
    @classmethod
    def decode_str(cls, text: Optional[str]) -> str:
        """解码邮件字符串
        
        Args:
            text: 需要解码的文本
            
        Returns:
            str: 解码后的文本
        """
        if not text:
            return ''
            
        try:
            decoded_list = decode_header(text)
            result = []
            
            for content, charset in decoded_list:
                if isinstance(content, bytes):
                    try:
                        # 尝试使用指定的字符集解码
                        if charset:
                            decoded = content.decode(charset)
                        else:
                            # 如果没有指定字符集，尝试常用编码
                            for encoding in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                                try:
                                    decoded = content.decode(encoding)
                                    break
                                except UnicodeDecodeError:
                                    continue
                            else:
                                # 如果所有编码都失败，使用 errors='replace'
                                decoded = content.decode('utf-8', errors='replace')
                    except Exception as e:
                        cls.logger.warning("解码文本失败: {}, 错误: {}", 
                                         content, LogHandler.format_error(e))
                        decoded = str(content)
                else:
                    decoded = str(content)
                    
                result.append(decoded)
                
            return ''.join(result)
            
        except Exception as e:
            cls.logger.error("解码邮件字符串失败: {}, 错误: {}", 
                           text, LogHandler.format_error(e))
            return str(text)
    
    @classmethod
    def decode_filename(cls, filename: Optional[Union[str, bytes]]) -> str:
        """解码附件文件名
        
        Args:
            filename: 需要解码的文件名
            
        Returns:
            str: 解码后的文件名
        """
        if not filename:
            return ''
            
        if isinstance(filename, bytes):
            try:
                return filename.decode()
            except UnicodeDecodeError:
                try:
                    return filename.decode('gbk')
                except UnicodeDecodeError:
                    return filename.decode('utf-8', errors='replace')
                    
        return cls.decode_str(filename) 