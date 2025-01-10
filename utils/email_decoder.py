from email.header import decode_header

class EmailDecoder:
    @staticmethod
    def decode_str(text, default_encoding='utf-8'):
        """解码文本内容"""
        if text is None:
            return ""
        decoded_parts = []
        for part, encoding in decode_header(text):
            if isinstance(part, bytes):
                try:
                    decoded_parts.append(part.decode(encoding or default_encoding))
                except:
                    decoded_parts.append(part.decode(default_encoding, 'ignore'))
            else:
                decoded_parts.append(part)
        return ''.join(decoded_parts)

    @staticmethod
    def decode_filename(filename, default_encoding='utf-8'):
        """解码文件名"""
        if not filename:
            return None
        return EmailDecoder.decode_str(filename, default_encoding) 