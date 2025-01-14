from typing import List, Set, Dict, Any
from services.email_service import EmailService
from services.rule_processor import RuleProcessor
from utils.log_handler import LogHandler
from models.email_message import EmailMessage
from utils.excel_processor import ExcelProcessor

class EmailProcessor:
    """
    邮件处理器，负责处理邮件和附件
    
    主要功能：
    1. 获取未读邮件
    2. 根据规则处理邮件
    3. 下载并处理附件
    4. 合并处理结果并保存
    """
    
    def __init__(self):
        """
        初始化邮件处理器
        
        初始化：
        1. 日志处理器
        2. 规则处理器
        3. 邮件服务
        4. Excel处理器
        5. 已处理邮件主题集合
        """
        self.logger = LogHandler().get_logger('EmailProcessor')
        self.rule_processor = RuleProcessor()
        self.email_service = EmailService()
        self.excel_processor = ExcelProcessor()
        self._processed_subjects: Set[str] = set()  # 用于跟踪已处理的邮件主题
        
    def _is_processed(self, subject: str) -> bool:
        """
        检查邮件是否已处理
        
        Args:
            subject: 邮件主题
            
        Returns:
            bool: 如果邮件已处理返回True，否则返回False
        """
        return subject in self._processed_subjects
        
    def _mark_as_processed(self, subject: str):
        """
        将邮件标记为已处理
        
        Args:
            subject: 邮件主题
        """
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
        """
        处理所有未读邮件
        
        处理流程：
        1. 获取所有未读邮件
        2. 对每封邮件：
           - 检查是否匹配处理规则
           - 下载附件
           - 如果是送货单规则，处理对应目录下的所有Excel文件
        3. 返回处理结果
        """
        try:
            # 获取未读邮件列表
            email_list = self.email_service.get_unread_emails()
            if not email_list:
                self.logger.info("没有找到未读邮件")
                return

            # 处理每封未读邮件
            for email_msg in email_list:
                try:
                    # 获取匹配的规则
                    rule = self.rule_processor.get_matching_rule(email_msg)
                    if not rule:
                        continue

                    # 下载附件
                    attachments = self._process_single_email(email_msg)
                    if not attachments:
                        continue

                    # 如果是送货单规则，处理Excel文件
                    if "送货单" in rule["name"]:
                        excel_processor = ExcelProcessor()
                        try:
                            # 处理指定目录下的所有Excel文件
                            data_dict = excel_processor.process_excel(rule["download_path"], rule["name"])
                            if data_dict:
                                self.logger.info(f"成功处理 {rule['name']} 的Excel文件")
                            
                        except Exception as e:
                            self.logger.error(f"处理Excel文件失败: {str(e)}")
                            continue

                except Exception as e:
                    self.logger.error(f"处理邮件时出错: {str(e)}")
                    continue

        except Exception as e:
            self.logger.error(f"处理未读邮件时出错: {str(e)}")
            raise
        finally:
            # 确保在处理完成后断开邮件连接
            self.email_service.disconnect() 