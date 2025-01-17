from typing import List, Set, Dict, Any
from services.email_service import EmailService
from services.rule_processor import RuleProcessor
from utils.log_handler import LogHandler
from models.email_message import EmailMessage
from utils.excel_processor import ExcelProcessor
from workflows.erp_receipt import process_delivery_orders

class EmailProcessor:
    """
    邮件处理器，负责处理邮件及其附件
    
    主要功能：
    1. 处理单个邮件和批量邮件
    2. 下载和保存邮件附件
    3. 根据规则匹配邮件
    4. 处理Excel附件
    5. 错误处理和日志记录
    """
    
    def __init__(self, rule_processor: RuleProcessor, email_service: EmailService):
        """
        初始化邮件处理器
        
        Args:
            rule_processor: 规则处理器实例
            email_service: 邮件服务实例
            
        Raises:
            Exception: 初始化失败时抛出
        """
        self.logger = LogHandler().get_logger('EmailProcessor', file_level='DEBUG', console_level='INFO')
        self.rule_processor = rule_processor
        self.email_service = email_service
        self.excel_processor = ExcelProcessor()
        self._processed_subjects: Set[str] = set()
        
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
        
    def _process_single_email(self, email_msg: EmailMessage) -> bool:
        """
        处理单个邮件
        
        处理流程：
        1. 检查邮件是否已处理
        2. 匹配处理规则
        3. 下载匹配的附件
        4. 处理Excel文件
        
        Args:
            email_msg: 邮件对象
            
        Returns:
            bool: 处理成功返回True，否则返回False
        """
        try:
            # 检查是否已处理过
            if self._is_processed(email_msg.subject):
                self.logger.debug("跳过已处理邮件: %s", email_msg.subject)
                return False
                
            # 标记为已处理
            self._mark_as_processed(email_msg.subject)
            
            # 获取匹配的规则
            matching_rule = self.rule_processor.get_matching_rule(email_msg)
            if not matching_rule:
                self.logger.debug("邮件不匹配任何规则: %s", email_msg.subject)
                return False
                
            self.logger.info("处理邮件 [%s] - 匹配规则: %s", email_msg.subject, matching_rule['name'])
            
            # 加载完整邮件内容
            if not self.email_service.load_full_message(email_msg):
                self.logger.error("无法加载邮件内容: %s", email_msg.subject)
                return False
            
            # 获取附件
            downloaded_files = self.email_service.get_attachments(email_msg, matching_rule)
            if not downloaded_files:
                self.logger.info("邮件 [%s] 没有匹配的附件", email_msg.subject)
                return False
                
            # 标记为已读
            self.email_service.mark_as_read(email_msg)
            self.logger.info("完成处理邮件 [%s] - 下载附件数: %d", email_msg.subject, len(downloaded_files))
            return True
            
        except Exception as e:
            self.logger.error("处理邮件出错 [%s]: %s", 
                            email_msg.subject, LogHandler.format_error(e))
            return False
            
    def process_unread_emails(self) -> bool:
        """
        处理所有未读邮件
        
        处理流程：
        1. 获取所有未读邮件
        2. 逐个处理邮件
        3. 记录处理结果
        
        Returns:
            bool: 所有邮件处理成功返回True，否则返回False
        """
        try:
            # 获取未读邮件列表
            try:
                email_list = self.email_service.get_unread_emails()
            except Exception as e:
                self.logger.error("获取未读邮件失败: %s", LogHandler.format_error(e))
                return False

            if not email_list:
                self.logger.info("没有未读邮件")
                return True

            self.logger.info("开始处理 %d 封未读邮件", len(email_list))

            # 处理每封未读邮件
            for email_msg in email_list:
                try:
                    # 获取匹配的规则
                    rule = self.rule_processor.get_matching_rule(email_msg)
                    if not rule:
                        continue

                    # 下载附件
                    if not self._process_single_email(email_msg):
                        continue

                    # 如果是送货单规则，处理Excel文件
                    if "送货单" in rule["name"]:
                        excel_processor = ExcelProcessor()
                        try:
                            data_dict = excel_processor.process_excel(rule["download_path"], rule["name"])
                            if data_dict:
                                print(data_dict)
                                self.logger.info("成功处理 [%s] 的Excel文件", rule["name"])
                                
                                if process_delivery_orders(data_dict):
                                    self.logger.info("成功完成 [%s] 的送货单录入", rule["name"])
                                else:
                                    self.logger.error("送货单录入失败: %s", rule["name"])
                            
                        except Exception as e:
                            self.logger.error("处理Excel文件失败: %s", LogHandler.format_error(e))
                            continue

                except Exception as e:
                    self.logger.error("处理邮件失败: %s", LogHandler.format_error(e))
                    continue

            return True

        except Exception as e:
            self.logger.error("处理未读邮件时发生错误: %s", LogHandler.format_error(e))
            return False
        finally:
            try:
                self.email_service.disconnect()
            except Exception as e:
                self.logger.error("断开邮箱连接失败: %s", LogHandler.format_error(e)) 