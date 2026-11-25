# doc2md-skill

[English](README.md) | 简体中文

**用 AI Agent 辅助读论文时使用的 Agent Skill。** 它把 PDF / Word 文档转成结构完整的 Markdown，并附上一份核对报告，让 Agent（Claude Code、Codex 等）按章节读论文：内容读全，也不浪费 token。

## 为什么做这个工具

这个工具是在日常用 Agent 辅助读论文的过程中做出来的。现在大家读论文基本都会让 Agent 帮忙，而 Agent 直接读 PDF 时，总是反复出现这几个问题：

- **整篇一次性灌进上下文。** 浪费 token，真正重要的内容反而被挤掉。
- **内容悄悄丢失。** 无框表格和三线表被压成一段文字，标题丢了，公式乱码，扫描页直接跳过，Word 批注根本看不到。
- **丢了什么没人知道。** 你和 Agent 都无法判断，回答到底是不是基于完整的阅读。

doc2md 的办法是在"读"之前加一道转换：先把文档转成干净的 Markdown，逐行和原文核对，再让 Agent 只读需要的章节。

## 它做什么

- **保留结构**：标题层级与编号（包括 Word 自定义样式、手工加粗放大的标题、用表格排版的章标题）、表格（有框、无框、三线表、跨页续表）、图片位置都保留。图片按阅读顺序提取为 `images/1.png`、`2.png`……
- **不丢内容**：Word 批注、修订、页眉页脚、图表数据、SmartArt、文本框、公式、嵌入文件一并输出。没有文字层的页（扫描件、纯图片幻灯片），以及没能逐字转出的行（比如公式），都会附上整页图，供 Agent 对照原貌。
- **自带核对**：原文逐行与输出比对，并核对标题、PDF 书签、表格题注，检查结构有没有丢。报告会明确告诉 Agent 该去哪里看。
- **省 token**：去掉重复的页眉页脚、页码和目录页；冗长的 HTML 表格转成紧凑的 Markdown 表格；Agent 先看大纲，再只读需要的章节。

报告长这样（节选，一篇 7 页的期刊论文）：

```
字符数：16986  预估 token：约 11226
── 转换报告 ──
  页数：7
  大字号标题(引擎漏判后补出)：2
  降级的伪标题(题注/整句/公式)：1
  补出的标题(按同级编号/按书签)：1/0
  逐行核对(原文行数/缺失行数)：713/0  覆盖率 100.00%
  缺失或乱序的行：无
  大字号标题核对(原文/未成为标题)：4/0
  题注提到却未输出的表格：无
```

## 安装

1. 克隆到 Agent 的 skills 目录：

   ```bash
   # Claude Code
   git clone https://github.com/searcyan/doc2md-skill ~/.claude/skills/doc2md
   # Codex
   git clone https://github.com/searcyan/doc2md-skill ~/.codex/skills/doc2md
   ```

2. 安装 Python 依赖（Python 3.10+）：

   ```bash
   python -m pip install -r ~/.claude/skills/doc2md/requirements.txt
   ```

3. 处理 Word 文档需要 [pandoc](https://pandoc.org/installing.html)（也可以 `python -m pip install pypandoc_binary`）。
4. 可选：[LibreOffice](https://www.libreoffice.org/)，用于处理旧版 `.doc` 和 Word 里的 EMF/WMF 图片。
5. 检查环境：

   ```bash
   python ~/.claude/skills/doc2md/scripts/doc2md.py --check
   ```

工具不在 `PATH` 中时，可用环境变量 `DOC2MD_PANDOC` / `DOC2MD_SOFFICE` 指定路径。

## 使用

装好后，直接让 Agent 读 PDF 或 Word 文件即可。skill 会让它先转换，再按需读大纲、章节和图片。

也可以手动运行：

```bash
python scripts/doc2md.py paper.pdf
```

输出在源文件旁的 `_md/paper/`：`paper.md`、`images/`（文中图片）、`pages/`（整页图，按需生成）、`attachments/`（Word 嵌入文件，有才生成）。参数：`-o <目录>`、`--password <密码>`（加密 PDF）、`--keep-toc`、`--keep-header-footer`、`--dpi 150`。

## 测试范围

在 Windows 11 + Python 3.14 上测试过：知网中文期刊论文、英文期刊论文、arXiv 双栏论文、Word 报告和毕业论文（自定义标题样式、批注、修订），以及纯图片的幻灯片 PDF。工具查找逻辑也覆盖了 macOS 和 Linux，但还没在这两个平台上实测，欢迎提 issue 反馈。

## 说明

- `pymupdf`、`pymupdf4llm`、`pymupdf-layout` 必须保持同一个固定版本（1.28.2），不要单独升级其中一个。
- 不内置 OCR：没有文字层的页会以整页图交给 Agent，多模态 Agent 可以直接识读。

## 许可证

[AGPL-3.0](LICENSE)，与核心依赖 [PyMuPDF](https://github.com/pymupdf/PyMuPDF)（AGPL-3.0 / Artifex 商业许可）保持一致。pandoc 为 GPL，以独立程序方式调用。
