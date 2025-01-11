import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 邮箱配置
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
EMAIL_SERVER = os.getenv('EMAIL_SERVER')
EMAIL_SERVER_PORT = int(os.getenv('EMAIL_SERVER_PORT', '993'))
EMAIL_USE_SSL = os.getenv('EMAIL_USE_SSL', 'True').lower() == 'true'

# 创建基本目录
os.makedirs('downloads', exist_ok=True)  # 确保下载目录存在 