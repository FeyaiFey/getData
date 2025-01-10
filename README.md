# Email Attachment Downloader

这是一个基于Python的邮件附件自动下载工具，可以根据配置的规则自动检查未读邮件，下载匹配的Excel附件。

## 功能特点

- 自动检查未读邮件
- 支持多种邮件匹配规则：
  - 邮件主题匹配（支持正则表达式）
  - 发件人匹配
  - 收件人匹配
  - 附件名称匹配（支持正则表达式）
- 自动下载匹配的Excel附件
- 支持自定义下载路径
- 下载完成后自动标记邮件为已读

## 安装步骤

1. 克隆代码仓库：
```bash
git clone [repository_url]
cd email-attachment-downloader
```

2. 创建并激活虚拟环境：
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux/Mac
source venv/bin/activate
```

3. 安装依赖包：
```bash
pip install -r requirements.txt
```

4. 配置环境变量：
   - 复制 `.env.example` 为 `.env`
   - 修改 `.env` 文件中的邮箱配置：
     ```
     EMAIL_ADDRESS=your-email@example.com
     EMAIL_PASSWORD=your-password
     EMAIL_SERVER=imap.example.com
     EMAIL_SERVER_PORT=993
     EMAIL_USE_SSL=True
     ```

5. 配置下载规则：
   - 修改 `email_rules.yaml` 文件，添加或修改邮件匹配规则

## 使用说明

1. 启动程序：
```bash
python main.py
```

2. 程序会自动：
   - 检查未读邮件
   - 根据规则匹配邮件
   - 下载匹配的Excel附件
   - 将处理过的邮件标记为已读

## 配置规则说明

在 `email_rules.yaml` 文件中配置邮件匹配规则，每条规则包含以下字段：

```yaml
- name: "规则名称"
  subject_contains: ["邮件主题正则表达式"]
  sender_contains: ["发件人邮箱"]
  receiver_contains: ["收件人邮箱1", "收件人邮箱2"]
  attachment_name_pattern: ["附件名称正则表达式"]
  download_path: "下载路径"
```

示例：
```yaml
- name: "示例规则"
  subject_contains: ["^重要文件 \\d{2}/\\d{2}$"]
  sender_contains: ["sender@example.com"]
  receiver_contains: ["receiver@example.com"]
  attachment_name_pattern: ["^Report-\\d{8}\\.xlsx?$"]
  download_path: "downloads/reports"
```

## 项目结构

```
email-attachment-downloader/
├── main.py                 # 主程序入口
├── config.py              # 配置加载
├── email_rules.yaml       # 邮件规则配置
├── requirements.txt       # 项目依赖
├── .env                  # 环境变量配置
├── services/             # 服务模块
│   ├── __init__.py
│   ├── email_processor.py  # 邮件处理器
│   ├── email_service.py    # 邮件服务
│   └── rule_processor.py   # 规则处理器
├── models/               # 数据模型
│   ├── __init__.py
│   └── email_message.py    # 邮件消息模型
└── utils/                # 工具类
    ├── __init__.py
    ├── email_decoder.py    # 邮件解码工具
    └── file_handler.py     # 文件处理工具
```

## 注意事项

1. 请确保邮箱开启了IMAP访问
2. 建议使用应用专用密码而不是邮箱密码
3. 下载路径必须具有写入权限
4. 正则表达式需要使用双反斜杠转义特殊字符 