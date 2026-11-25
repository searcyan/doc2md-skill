# doc2md-skill

English | [简体中文](README.zh-CN.md)

**An Agent Skill for reading papers with AI agents.** It converts PDF / Word documents into structure-preserving Markdown plus a verification report, so your agent (Claude Code, Codex, …) reads a paper section by section — completely, and without burning tokens.

## Why

This tool came out of everyday paper reading with AI agents. Most people now read papers with an agent's help, and when an agent reads a PDF directly, the same problems keep showing up:

- **The whole paper goes into the context at once.** Tokens are wasted, and the content that matters gets crowded out.
- **Content is silently lost.** Borderless and three-line tables are flattened into prose, headings disappear, formulas are garbled, scanned pages are skipped, and Word comments never show up.
- **Nobody knows what was lost.** Neither you nor the agent can tell whether the answer was based on a complete reading.

doc2md fixes this by putting a conversion step in front of reading: first turn the document into clean Markdown and check it line by line against the original, then let the agent read only the sections it needs.

## What it does

- **Keeps the structure.** Heading hierarchy and numbering (including custom Word styles, manually bolded/enlarged headings, and chapter titles laid out as tables), tables (ruled, borderless, three-line, cross-page continuations), and image positions are all preserved. Images are extracted as `images/1.png`, `2.png`, … in reading order.
- **Doesn't drop content.** Word comments, tracked changes, headers/footers, chart data, SmartArt, text boxes, equations, and embedded files are all exported. Pages without a text layer (scans, image-only slides) and lines that could not be converted verbatim (e.g. formulas) come with full-page images, so the agent can look at the original.
- **Verifies itself.** Every line of the original is checked against the output, and headings, PDF bookmarks, and table captions are checked for structural losses. The report tells the agent exactly where to look.
- **Saves tokens.** Repeated headers/footers, page numbers, and table-of-contents pages are removed; bulky HTML tables become compact Markdown tables; the agent reads the outline first and then only the sections it needs.

A report looks like this (excerpt, 7-page journal paper):

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

The report and messages are in Chinese; the conversion works for documents in any language (tested on Chinese and English papers).

## Install

1. Clone into your agent's skills directory:

   ```bash
   # Claude Code
   git clone https://github.com/searcyan/doc2md-skill ~/.claude/skills/doc2md
   # Codex
   git clone https://github.com/searcyan/doc2md-skill ~/.codex/skills/doc2md
   ```

2. Install Python dependencies (Python 3.10+):

   ```bash
   python -m pip install -r ~/.claude/skills/doc2md/requirements.txt
   ```

3. For Word documents, install [pandoc](https://pandoc.org/installing.html) (or `python -m pip install pypandoc_binary`).
4. Optional: [LibreOffice](https://www.libreoffice.org/), needed for legacy `.doc` files and EMF/WMF images in Word documents.
5. Check the environment:

   ```bash
   python ~/.claude/skills/doc2md/scripts/doc2md.py --check
   ```

If a tool is not on `PATH`, point to it with the `DOC2MD_PANDOC` / `DOC2MD_SOFFICE` environment variables.

## Usage

Once installed, just ask your agent to read a PDF or Word file — the skill tells it to convert first and then read the outline, sections, and images on demand.

You can also run it by hand:

```bash
python scripts/doc2md.py paper.pdf
```

Output goes to `_md/paper/` next to the source file: `paper.md`, `images/` (figures), `pages/` (full-page images, only when needed), and `attachments/` (files embedded in Word, only when present). Options: `-o <dir>`, `--password <pwd>` for encrypted PDFs, `--keep-toc`, `--keep-header-footer`, `--dpi 150`.

## Tested on

Windows 11 with Python 3.14: Chinese journal papers (CNKI), English journal articles, two-column arXiv papers, Word reports and theses (custom heading styles, comments, tracked changes), and image-only slide decks. Tool discovery also covers macOS and Linux, but those platforms have not been tested yet — issues and reports are welcome.

## Notes

- `pymupdf`, `pymupdf4llm`, and `pymupdf-layout` must stay on the same pinned version (1.28.2). Don't upgrade one of them alone.
- There is no built-in OCR: pages without a text layer are handed to the agent as full-page images, which multimodal agents can read directly.

## License

[AGPL-3.0](LICENSE), matching its core dependency [PyMuPDF](https://github.com/pymupdf/PyMuPDF) (AGPL-3.0 / Artifex commercial license). pandoc is GPL and is invoked as a separate program.
