# PPTX MCP Server v3.0

基于 python-pptx 的 PowerPoint MCP Server，提供安全的 PPTX 文件操作能力。

## 功能

- ✅ 创建/打开/保存 PPTX 文件
- ✅ 添加幻灯片、文本、图片、表格
- ✅ 读取演示文稿内容
- ✅ 完整的安全验证
  - ZIP 炸弹防护
  - 宏/VBA 检测
  - 路径遍历防护
  - XML 实体注入防护
  - 资源限制
- ✅ 会话管理和超时清理
- ✅ 临时文件自动清理

## 安装

```bash
cd pptx-mcp-server
pip install -r requirements.txt
```

## 启动

```bash
python server.py --port 8010 --token your-secure-token
```

## 环境变量

- `PPTX_SERVER_PORT`: 服务端口（默认 8010）
- `PPTX_SERVER_TOKEN`: 认证令牌
- `PPTX_WORK_DIR`: 工作目录

## MCP 工具列表

| 工具 | 描述 |
|------|------|
| pptx_create | 创建空白演示文稿 |
| pptx_open | 打开现有文件 |
| pptx_save | 保存演示文稿 |
| pptx_close | 关闭会话 |
| pptx_info | 获取信息 |
| pptx_add_slide | 添加幻灯片 |
| pptx_add_text | 添加文本 |
| pptx_add_image | 添加图片 |
| pptx_add_table | 添加表格 |
| pptx_read_content | 读取内容 |
| pptx_list_slides | 列出幻灯片 |
| pptx_validate | 验证文件安全性 |

## 安全限制

- 最大文件大小：50MB
- 最大幻灯片数：500
- 会话超时：1小时
- 处理超时：60秒

## OpenClaw 集成

在 OpenClaw 配置中添加：

```json
{
  "mcpServers": {
    "pptx": {
      "url": "http://127.0.0.1:8010/mcp",
      "token": "your-secure-token"
    }
  }
}
```

## 项目结构

```
pptx-mcp-server/
├── server.py           # 主服务器
├── requirements.txt    # 依赖
├── security/           # 安全模块
│   ├── validator.py    # 文件验证
│   ├── session.py      # 会话管理
│   └── tempfile.py     # 临时文件管理
└── tools/              # 工具实现
    └── manager.py      # 工具管理器
```

## 评分

- Kiro 评估：7.8/10（无 P0 问题）
- 安全设计：8/10
- 生产就绪：7/10
