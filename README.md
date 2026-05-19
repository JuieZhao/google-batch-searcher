# 🔍 Google Batch Searcher — 批量 Google 搜索工具

从 Excel 读取搜索词，通过真实 Chrome 浏览器批量搜索 Google，按关键词过滤结果，导出链接。

## ✨ 功能

- 📂 **Excel 驱动** — 搜索词写在 Excel 里，工具自动逐条搜索
- 🦊 **真实浏览器搜索** — 使用 undetected-chromedriver 启动 Chrome，绕过 Google 反爬检测
- 🔍 **批量 Google 搜索** — 每条搜索词可返回 1-30 个结果
- 🎯 **关键词过滤** — 支持"保留关键词"列，只保留链接包含指定关键词的结果
- 📋 **实时预览** — 搜索结果实时显示，保留/跳过分色标识
- 📥 **导出 Excel** — 两个 Sheet：全部结果 + 仅保留结果
- ⏹ **随时停止** — 可中途取消搜索
- ⏱️ **可调间隔** — 防止 Google 限流

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

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行

```bash
python searcher.py

# 或双击
启动搜索工具.bat
```

首次运行会自动启动 Chrome 浏览器窗口，这是正常的，**请勿手动关闭**。

## 📋 系统要求

- Python 3.9+
- Google Chrome 浏览器
- 网络能访问 Google（需要科学上网）

## 🛡️ 注意事项

- 搜索时会弹出 Chrome 窗口，请勿手动关闭或操作该窗口内的搜索页面
- 搜索间隔建议 ≥ 3 秒，太频繁可能触发 Google CAPTCHA
- 如遇到 Google 同意页/CAPTCHA，可手动在 Chrome 窗口完成验证后自动继续
- 搜索完成后浏览器会自动关闭
- 仅供个人研究学习使用

## 📦 依赖

```
selenium
undetected-chromedriver
openpyxl
```

## 📁 项目结构

```
google-batch-searcher/
├── searcher.py          # 主程序（Tkinter GUI）
├── 搜索词示例.xlsx       # Excel 模板
├── requirements.txt     # 依赖
├── 启动搜索工具.bat      # Windows 一键启动
└── README.md
```

## ❓ 为什么用真实浏览器而不是直接 HTTP 请求？

Google 会检测并拦截来自 Python `requests`/`urllib` 等库的直接 HTTP 请求，返回验证页面（CAPTCHA 或 consent page），导致搜索结果始终为 0。`undetected-chromedriver` 通过启动真实的 Chrome 浏览器、修改底层检测特征，让 Google 认为这是正常用户访问，从而返回真实搜索结果。

## ⚠️ 免责声明

本工具仅供个人学习研究使用。请遵守 Google 服务条款，合理控制搜索频率。
