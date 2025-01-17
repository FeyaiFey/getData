import email
from email.header import decode_header
from typing import Optional, Union
from utils.log_handler import LogHandler

class EmailDecoder:
    """邮件解码工具类，用于处理邮件编码相关的问题
    
    主要功能：
    1. 解码邮件主题、发件人、收件人等字符串
    2. 解码附件文件名
    3. 支持多种字符编码（utf-8, gbk, gb2312, iso-8859-1）
    4. 自动处理编码错误和替换字符
    """
    
    logger = LogHandler().get_logger('EmailDecoder', file_level='DEBUG', console_level='WARNING')
    
    @classmethod
    def decode_str(cls, text: Optional[str]) -> str:
        """解码邮件字符串
        
        支持多种编码格式的自动识别和转换，包括：
        - UTF-8
        - GBK
        - GB2312
        - ISO-8859-1
        
        Args:
            text: 需要解码的文本，可以是None
            
        Returns:
            str: 解码后的文本，如果输入为None则返回空字符串
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
                                cls.logger.warning("无法确定编码，使用替换字符: %s", content)
                    except Exception as e:
                        cls.logger.error("解码失败 [%s]: %s", content, LogHandler.format_error(e))
                        decoded = str(content)
                else:
                    decoded = str(content)
                    
                result.append(decoded)
                
            return ''.join(result)
            
        except Exception as e:
            cls.logger.error("字符串解码失败 [%s]: %s", text, LogHandler.format_error(e))
            return str(text)
    
    @classmethod
    def decode_filename(cls, filename: Optional[Union[str, bytes]]) -> str:
        """解码附件文件名
        
        支持对字节串和字符串两种格式的文件名进行解码。
        优先使用UTF-8编码，如果失败则尝试GBK编码。
        
        Args:
            filename: 需要解码的文件名，可以是字节串或字符串
            
        Returns:
            str: 解码后的文件名，如果输入为None则返回空字符串
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
                    cls.logger.warning("文件名解码失败，使用替换字符: %s", filename)
                    return filename.decode('utf-8', errors='replace')
                    
        return cls.decode_str(filename) 