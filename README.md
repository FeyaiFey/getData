# 自动邮件附件下载工具

这是一个基于 Python 的自动化工具，用于监控指定邮箱的未读邮件，并根据预设规则自动下载匹配的 Excel 附件。

## 主要功能

- 自动监控未读邮件
- 支持多种邮件匹配规则：
  - 邮件主题匹配（支持正则表达式）
  - 发件人匹配
  - 收件人匹配
  - 附件名称匹配（支持正则表达式）
- 自动下载匹配的 Excel 附件
- 支持自定义下载路径
- 下载完成后自动标记邮件为已读
- 支持定时任务，定期检查新邮件
- 详细的日志记录

## 系统要求

- Python 3.6+
- Windows/Linux/MacOS
- 邮箱需要开启 IMAP 访问

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
   - 修改 `.env` 文件中的配置：
     ```ini
     # 邮箱配置
     EMAIL_ADDRESS=your-email@example.com
     EMAIL_PASSWORD=your-password
     EMAIL_SERVER=imap.example.com
     EMAIL_SERVER_PORT=993
     EMAIL_USE_SSL=True

     # 邮件过滤配置
     DAYS_LOOK_BACK=7

     # 文件路径配置
     OUTPUT_PATH=downloads
     ```

5. 配置下载规则：
   修改 `email_rules.yaml` 文件，添加或修改邮件匹配规则

## 配置文件说明

### 邮件规则配置 (email_rules.yaml)

每条规则包含以下字段：

```yaml
- name: "规则名称"                        # 规则的描述性名称
  subject_contains: ["主题正则表达式"]     # 邮件主题匹配模式（支持正则）
  sender_contains: ["发件人邮箱"]         # 发件人邮箱匹配
  receiver_contains: ["收件人邮箱"]       # 收件人邮箱匹配
  attachment_name_pattern: ["附件正则"]   # 附件名称匹配模式（支持正则）
  download_path: "下载路径"              # 附件保存路径
```

规则示例：
```yaml
- name: "每日报表"
  subject_contains: ["^日报表 \\d{2}/\\d{2}$"]
  sender_contains: ["report@company.com"]
  receiver_contains: ["me@company.com"]
  attachment_name_pattern: ["^Report-\\d{8}\\.xlsx$"]
  download_path: "downloads/daily_reports"
```

### 环境变量配置 (.env)

必需的环境变量：
- `EMAIL_ADDRESS`: 邮箱地址
- `EMAIL_PASSWORD`: 邮箱密码（建议使用应用专用密码）
- `EMAIL_SERVER`: IMAP 服务器地址
- `EMAIL_SERVER_PORT`: IMAP 服务器端口
- `EMAIL_USE_SSL`: 是否使用 SSL 连接（True/False）

可选的环境变量：
- `DAYS_LOOK_BACK`: 检查几天内的邮件（默认7天）
- `OUTPUT_PATH`: 默认下载路径

## 使用说明

1. 启动程序：
```bash
python main.py
```

2. 程序会：
   - 每分钟自动检查一次未读邮件
   - 根据规则匹配邮件
   - 下载匹配的 Excel 附件
   - 将处理过的邮件标记为已读

## 项目结构

```
项目根目录/
├── main.py                 # 主程序入口
├── config.py              # 配置加载
├── email_rules.yaml       # 邮件规则配置
├── requirements.txt       # 项目依赖
├── .env                  # 环境变量配置
├── services/             # 服务模块
│   ├── email_processor.py  # 邮件处理器
│   ├── email_service.py    # 邮件服务
│   └── rule_processor.py   # 规则处理器
├── models/               # 数据模型
│   └── email_message.py    # 邮件消息模型
└── utils/                # 工具类
    ├── email_decoder.py    # 邮件解码工具
    └── file_handler.py     # 文件处理工具
```

## 核心模块说明

### EmailProcessor
- 处理未读邮件的主要逻辑
- 调用其他服务处理邮件和附件

### EmailService
- 处理邮件服务器连接
- 获取未读邮件
- 下载附件
- 标记邮件状态

### RuleProcessor
- 加载和解析规则配置
- 匹配邮件和规则
- 验证附件名称

### FileHandler
- 处理文件保存
- 生成唯一文件名
- 确保目录存在

## 注意事项

1. 邮箱安全：
   - 建议使用应用专用密码而不是邮箱密码
   - 确保邮箱开启了 IMAP 访问

2. 文件安全：
   - 确保下载路径具有写入权限
   - 定期清理下载目录，避免空间不足

3. 正则表达式：
   - 在 YAML 文件中使用正则表达式时需要双反斜杠转义
   - 测试正则表达式以确保匹配符合预期

4. 性能考虑：
   - 合理设置检查间隔时间
   - 适当配置 `DAYS_LOOK_BACK` 参数
   - 定期清理已处理的邮件

## 常见问题

1. 连接错误
   - 检查邮箱服务器配置
   - 确认网络连接
   - 验证邮箱凭据

2. 附件未下载
   - 检查规则配置是否正确
   - 确认附件名称匹配模式
   - 查看程序日志获取详细信息

3. 重复下载
   - 检查文件命名规则
   - 确认邮件标记为已读功能正常

## 开发计划

- [ ] 添加 Web 界面
- [ ] 支持更多文件格式
- [ ] 添加邮件过滤规则
- [ ] 优化性能
- [ ] 添加单元测试
- [ ] 支持多邮箱配置

## 贡献指南

欢迎提交 Issue 和 Pull Request。在提交代码前，请确保：

1. 代码符合 PEP 8 规范
2. 添加必要的注释和文档
3. 更新相关的配置示例
4. 测试代码功能

## 许可证

MIT License 