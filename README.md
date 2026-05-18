# 🔍 Google Batch Searcher — 批量 Google 搜索工具

从 Excel 读取搜索词，批量搜索 Google，按关键词过滤结果，导出链接。

## ✨ 功能

- 📂 **Excel 驱动** — 搜索词写在 Excel 里，工具自动逐条搜索
- 🔍 **批量 Google 搜索** — 每条搜索词可返回 1-50 个结果
- 🎯 **关键词过滤** — 支持"保留关键词"列，只保留链接包含指定关键词的结果
- 📋 **实时预览** — 搜索结果实时显示，保留/跳过分色标识
- 📥 **导出 Excel** — 两个 Sheet：全部结果 + 仅保留结果
- ⏹ **随时停止** — 可中途取消搜索
- ⏱️ **可调间隔** — 防止 Google 限流封 IP

## 📊 Excel 格式

| A 列 | B 列（可选） |
|------|-------------|
| 搜索词 | 保留关键词（逗号分隔） |

示例：

| 搜索词 | 保留关键词 |
|--------|-----------|
| Python tkinter Windows 清理工具 | github.com |
| 通义千问 开源论文 2025 | arxiv |
| Apple Design GUI toolkit | github, open-source |

- **B 列为空** = 该条搜索的所有结果都保留
- **B 列有关键词** = 只保留链接中包含任一关键词的结果
- 多关键词用逗号分隔（中英文逗号均可）

## 🚀 使用

```bash
# 安装依赖
pip install requests beautifulsoup4 openpyxl

# 运行
python searcher.py

# 或双击
启动搜索工具.bat
```

## 📋 系统要求

- Python 3.9+
- 网络能访问 Google（需要科学上网）

## 🛡️ 注意事项

- 搜索间隔建议 ≥ 2 秒，太频繁会被 Google 限流（429）
- 工具通过 Chrome 浏览器 UA 模拟正常访问
- 仅供个人研究学习使用

## 📦 依赖

```
requests
beautifulsoup4
openpyxl
```

## 📁 项目结构

```
google-batch-searcher/
├── searcher.py          # 主程序
├── 搜索词示例.xlsx       # Excel 模板
├── requirements.txt     # 依赖
└── 启动搜索工具.bat      # Windows 一键启动
```

## ⚠️ 免责声明

本工具仅供个人学习研究使用。请遵守 Google 服务条款，合理控制搜索频率。
