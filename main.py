import time
import schedule
from datetime import datetime
from services.email_processor import EmailProcessor

def job():
    """定时任务主函数"""
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 开始检查未读邮件...")
    
    try:
        processor = EmailProcessor()
        processor.process_unread_emails()
    except Exception as e:
        print(f"执行任务时出错: {str(e)}")

def main():
    print("邮件自动下载程序已启动...")
    print("正在监控未读邮件...")
    
    # 设置定时任务，每10分钟执行一次
    schedule.every(10).minutes.do(job)

    # 首次运行
    job()
    
    # 持续运行定时任务
    while True:
        schedule.run_pending()
        time.sleep(60)  # 每600秒检查一次是否需要执行任务

if __name__ == "__main__":
    main() 