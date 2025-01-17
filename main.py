import time
import schedule
from datetime import datetime
from services.email_processor import EmailProcessor
from services.email_service import EmailService
from services.rule_processor import RuleProcessor
from utils.log_handler import LogHandler
from utils.file_handler import FileHandler

logger = LogHandler().get_logger('Main', file_level='DEBUG', console_level='INFO')

def check_emails():
    """检查未读邮件并处理"""
    try:
        logger.info("开始检查未读邮件...")
        
        # 创建服务实例
        rule_processor = RuleProcessor()
        email_service = EmailService(rule_processor)
        
        # 创建邮件处理器
        email_processor = EmailProcessor(rule_processor, email_service)
        
        # 处理未读邮件
        email_processor.process_unread_emails()
        
    except Exception as e:
        logger.error("执行任务时出错: %s", str(e))
    finally:
        # 确保断开连接
        if 'email_service' in locals():
            email_service.disconnect()

def main():
    """主函数"""
    try:
        # 初始化目录结构
        if not FileHandler.init_project_directories():
            logger.error("初始化目录结构失败，程序退出")
            return
            
        logger.info("邮件自动下载程序已启动...")
        
        # 设置定时任务
        schedule.every(10).minutes.do(check_emails)
        logger.info("正在监控未读邮件...")
        
        # 立即执行一次
        check_emails()
        
        # 持续运行定时任务
        while True:
            try:
                schedule.run_pending()
                time.sleep(60)  # 每60秒检查一次是否需要执行任务
            except KeyboardInterrupt:
                logger.info("程序已停止")
                break
            except Exception as e:
                logger.error("运行时出错: %s", str(e))
                time.sleep(60)  # 发生错误时等待60秒后继续
    except Exception as e:
        logger.error("程序启动失败: %s", str(e))

if __name__ == "__main__":
    main() 