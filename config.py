import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 邮箱配置
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_SERVER = os.getenv('EMAIL_SERVER')
EMAIL_SERVER_PORT = int(os.getenv('EMAIL_SERVER_PORT'))
EMAIL_USE_SSL = os.getenv('EMAIL_USE_SSL', 'True').lower() == 'true'

# 邮件过滤配置
DAYS_LOOK_BACK = int(os.getenv('DAYS_LOOK_BACK', 7))

# 文件路径配置
OUTPUT_PATH = os.getenv('OUTPUT_PATH')

# 创建基本目录
os.makedirs(OUTPUT_PATH, exist_ok=True)
os.makedirs('downloads', exist_ok=True)  # 确保下载目录存在 