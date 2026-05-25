---
name: Official Document Template
description: >
  Generate Chinese government official documents (公文) compliant with GB/T 9704-2012.
  Supports notices, reports, requests, approvals, decisions, meeting minutes, and more.
  Produces perfectly formatted .docx with proper fonts, margins, headings, and line spacing.
version: 1.0.0
metadata:
  openclaw:
    requires:
      bins:
        - python
    emoji: "\U0001F4C4"
    homepage: https://github.com/user/official-doc-template
    install:
      - kind: uv
        package: python-docx
        bins: []
    envVars:
      - name: DOC_TEMPLATE_DIR
        required: false
        description: Custom directory for template assets. Defaults to the skill's built-in template.
---

# 📄 政府公文生成模板

> 🇨🇳 Chinese Government Document Generator — **GB/T 9704-2012 Compliant**

一键生成符合《党政机关公文格式》国家标准的 Word 文档。支持 **通知、报告、请示、批复、函、决定、意见、纪要、通报、公告、通告、议案** 等全部 15 种法定公文类型。

---

## ✨ 特性

- 📐 **严格遵循 GB/T 9704-2012** — 页边距、字体、字号、行距、缩进全部按国标
- 🎯 **标题层级智能排版** — 四级标题自动应用不同字体（方正小标宋 / 黑体 / 楷体 / 仿宋）
- 🔢 **数字 & 英文自动 Times New Roman** — 中西文字体分离设置
- 📎 **附件格式规范** — 悬挂缩进、全角序号
- 🖨️ **开箱即用** — 一条命令生成，无需手动调排版

---

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install python-docx
```

### 2. 准备正文内容

将正文写为 JSON 文件（`body.json`），每个段落指定类型：

```json
[
  {"type": "text",     "content": "当前，为贯彻落实……现将有关事项通知如下："},
  {"type": "heading1", "content": "一、提高思想认识"},
  {"type": "text",     "content": "各单位要充分认识……"},
  {"type": "heading2", "content": "（一）加强组织领导"},
  {"type": "text",     "content": "成立专项工作小组……"},
  {"type": "heading3", "content": "1．建立健全机制"},
  {"type": "heading4", "content": "（1）定期会商制度"},
  {"type": "text",     "content": "每月召开一次推进会……"}
]
```

### 3. 生成公文

```bash
python scripts/generate_doc.py \
  --title "关于进一步做好相关工作的通知" \
  --recipient "各有关单位：" \
  --body-file body.json \
  --sender "XX局办公室" \
  --date "2026年5月22日" \
  --attachments "附件一,附件二" \
  --output 通知.docx
```

### 参数说明

| 参数 | 必填 | 说明 |
|------|:--:|------|
| `--title` | ✅ | 公文标题 |
| `--recipient` | ✅ | 主送机关 |
| `--body-file` | ✅ | 正文 JSON 文件路径 |
| `--sender` | ✅ | 发文机关署名 |
| `--date` | ✅ | 成文日期（如 `2026年5月22日`）|
| `--attachments` | ❌ | 附件名称，逗号分隔 |
| `--output` | ❌ | 输出路径，默认 `output.docx` |

---

## 📋 格式规范速查

### 页面设置

| 属性 | 值 |
|------|-----|
| 纸张 | A4 (210mm × 297mm) |
| 上边距 | 37mm |
| 下边距 | 35mm |
| 左边距 | 28mm |
| 右边距 | 26mm |
| 版心 | 156mm × 225mm |

### 字号字体对照

| 元素 | 字体 | 字号 | 行距 | 对齐 |
|------|------|------|------|------|
| 🏷️ 公文标题 | **方正小标宋_GBK** | 22pt (二号) | 38pt | 居中 |
| 📌 一级标题 | **黑体** | 16pt (三号) | 28pt | 首行缩进2字符 |
| 📎 二级标题 | **楷体** | 16pt (三号) | 28pt | 首行缩进2字符 |
| 📝 三级标题 | 仿宋 + Times New Roman | 16pt | 28pt | 首行缩进2字符 |
| 📝 四级标题 | 仿宋 + Times New Roman | 16pt | 28pt | 首行缩进2字符 |
| 📄 正文 | 仿宋 + Times New Roman | 16pt | 28pt | 首行缩进2字符 |
| 👤 主送机关 | 仿宋 + Times New Roman | 16pt | 28pt | 顶格 |
| 📎 附件说明 | 仿宋 + Times New Roman | 16pt | 28pt | 悬挂缩进 |
| ✍️ 落款署名 | 仿宋 + Times New Roman | 16pt | 28pt | 右对齐偏移 |
| 📅 成文日期 | 仿宋 + Times New Roman | 16pt | 28pt | 右对齐偏移 |

### 标题层级编号规则

```
一、一级标题         ← 黑体，加粗
  （一）二级标题     ← 楷体，加粗
    1．三级标题      ← 仿宋，序号用全角点"．"
      （1）四级标题  ← 仿宋，括号用半角
```

---

## 🔧 在 AI 工具中使用

将此 skill 安装后，你可以直接用自然语言让 AI 生成公文：

> "帮我写一份关于做好节假日期间值班工作的通知，主送各直属单位，发文单位是办公室，日期今天。"

AI 会自动：
1. 理解公文类型与内容
2. 组织正文并标注标题层级
3. 调用 `scripts/generate_doc.py` 生成格式规范的 `.docx`
4. 返回可直接打印的文件

---

## ⚠️ 注意事项

- **字体依赖**：生成的 docx 使用了中文字体（方正小标宋_GBK、仿宋、黑体、楷体），请确保系统已安装这些字体。Windows 通常自带；macOS/Linux 可能需要额外安装。
- **版本兼容**：生成的 `.docx` 兼容 Word 2007+ / WPS / LibreOffice
- **三级标题序号**：必须用全角点 `．`（U+FF0E），不能用英文句点 `.`

---

## 📂 文件结构

```
official-doc-template/
├── SKILL.md                 # 本文件 — 技能定义与格式规范
├── scripts/
│   └── generate_doc.py      # Python 生成脚本
└── assets/
    └── 公文模板.docx         # 原始模板文件（格式参照）
```
