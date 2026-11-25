#!/usr/bin/env python
"""doc2md：把 PDF / DOCX / DOC 转成 Markdown，并把图片按出现顺序编号(1、2、3…)提取到 images 目录。

用法：
    python doc2md.py <输入文件> [-o 输出根目录] [--keep-toc] [--keep-header-footer] [--dpi 150] [--password 密码]
    python doc2md.py --check        # 检查运行环境（依赖库版本、pandoc、LibreOffice）

输出：
    <输出根目录>/<文件名>/<文件名>.md
    <输出根目录>/<文件名>/images/1.png, 2.png ...      文中图片
    <输出根目录>/<文件名>/pages/p3.png ...             需要对照原貌的整页图（无文字层的页、有内容未能转成文字的页）
    <输出根目录>/<文件名>/attachments/...              Word 中嵌入的原始文件（如 Excel）

设计要点（目标：信息不丢失 + 少耗 token）：
    PDF ：pymupdf4llm(layout 模式) 做底
          + 用字号/坐标纠正引擎：漏判的标题、无框表格/三线表重建；按同级编号与 PDF 书签补全标题；误判标题降级
          + 跨页重复页眉页脚与页码清理、续表合并、题注归位；无文字层的页附整页图
    DOCX：pandoc 做底 + 补回 Word 自动编号 + 自定义样式/目录书签识别标题 + HTML 表格转 Markdown
          + 批注、页眉页脚、图表数据、SmartArt 文字、嵌入文件一并输出
    共用：图片按首次出现顺序编号、相同内容去重、表格去对齐空格、伪标题降级
    核对：原文逐行/逐段与 md 比对得出"文字覆盖率"，并核对标题与表格结构，输出转换报告
"""
import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote
import xml.etree.ElementTree as ET

PYMUPDF_PIN = "1.28.2"  # pymupdf / pymupdf4llm / pymupdf-layout 已验证的版本，三者必须一致


# ───────────────────────── 外部工具定位 ─────────────────────────
def _first_existing(paths):
    for p in paths:
        if p and os.path.isfile(p):
            return p
    return None


def find_pandoc():
    """依次查找：环境变量 DOC2MD_PANDOC → PATH → pypandoc_binary 自带的 pandoc → 常见安装位置。"""
    env = os.environ.get("DOC2MD_PANDOC")
    if env and os.path.isfile(env):
        return env
    hit = shutil.which("pandoc")
    if hit:
        return hit
    try:
        import pypandoc  # pip install pypandoc_binary 会附带 pandoc 可执行文件
        p = pypandoc.get_pandoc_path()
        if p and os.path.isfile(p):
            return p
    except Exception:
        pass
    local = os.environ.get("LOCALAPPDATA", "")
    return _first_existing([
        os.path.join(local, "Pandoc", "pandoc.exe"),
        r"C:\Program Files\Pandoc\pandoc.exe",
        "/opt/homebrew/bin/pandoc", "/usr/local/bin/pandoc", "/usr/bin/pandoc",
    ])


def find_soffice():
    """依次查找：环境变量 DOC2MD_SOFFICE → PATH(soffice/libreoffice) → 常见安装位置。找不到返回 None。"""
    env = os.environ.get("DOC2MD_SOFFICE")
    if env and os.path.isfile(env):
        return env
    for name in ("soffice", "libreoffice"):
        hit = shutil.which(name)
        if hit:
            return hit
    return _first_existing([
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice", "/snap/bin/libreoffice",
    ])


def run_soffice(args, workdir: Path, timeout=300):
    """以独立用户配置运行 LibreOffice，避免与用户已打开的 LibreOffice 冲突。返回是否找到 soffice。"""
    soffice = find_soffice()
    if not soffice:
        return False
    profile = (workdir / "_lo_profile").resolve().as_uri()
    subprocess.run([soffice, f"-env:UserInstallation={profile}", "--headless", *args],
                   capture_output=True, timeout=timeout)
    return True


def _pkg_version(name):
    try:
        from importlib.metadata import version
        return version(name)
    except Exception:
        return None


def check_env() -> bool:
    """打印环境检查结果。PDF 与 DOCX 任一可用即返回 True。"""
    print(f"Python：{sys.version.split()[0]}（{sys.executable}）")
    vers = {n: _pkg_version(n) for n in ("pymupdf", "pymupdf4llm", "pymupdf-layout")}
    pdf_ok = all(vers.values())
    for n, v in vers.items():
        flag = "未安装" if not v else (v if v == PYMUPDF_PIN else f"{v}（未验证，推荐 {PYMUPDF_PIN}）")
        print(f"  {n}：{flag}")
    if pdf_ok and len(set(vers.values())) > 1:
        print("  ⚠ 三个 pymupdf 包版本不一致，可能报错")
    pandoc = find_pandoc()
    soffice = find_soffice()
    print(f"PDF 转换：{'可用' if pdf_ok else '不可用'}")
    print(f"DOCX 转换（pandoc）：{pandoc or '不可用'}")
    print(f"DOC 转换与 EMF/WMF 图片转换（LibreOffice，可选）：{soffice or '未找到（.doc 文件与 Word 中的 emf/wmf 图片需要它）'}")
    if not pdf_ok:
        print(f'  安装：python -m pip install "pymupdf=={PYMUPDF_PIN}" "pymupdf4llm=={PYMUPDF_PIN}" "pymupdf-layout=={PYMUPDF_PIN}"')
    if not pandoc:
        print("  安装 pandoc：https://pandoc.org/installing.html ，或 python -m pip install pypandoc_binary")
    return pdf_ok or bool(pandoc)


CJK_IDEO = r"\u2e80-\u9fff\uf900-\ufaff"          # 汉字
CJK_PUNCT = r"\u3000-\u303f\uff00-\uffef"          # 全角标点/符号
CJK = CJK_IDEO + CJK_PUNCT


# ───────────────────────── 文本/表格工具 ─────────────────────────
# 只删「汉字与汉字」「任意字符与全角标点」之间的空格；
# 汉字与字母/数字之间的空格可能是真正的字段分隔符，予以保留。
_sp_a = re.compile(rf"(?<=[{CJK_IDEO}]) +(?=[{CJK_IDEO}])")
_sp_b = re.compile(rf"(?<=\S) +(?=[{CJK_PUNCT}])")
_sp_c = re.compile(rf"(?<=[{CJK_PUNCT}]) +(?=\S)")
_prefix_re = re.compile(r"^(\s*(?:#{1,6} |[-*+] |\d+[.)] |> ))?(.*)$")
SEP_RE = re.compile(r"\|?[\s:\-|]+\|?")


def squeeze_cjk_space(line: str) -> str:
    prefix, body = _prefix_re.match(line).groups()
    body = _sp_c.sub("", _sp_b.sub("", _sp_a.sub("", body)))
    return (prefix or "") + body


def join_text(a: str, b: str) -> str:
    """拼接两段文字：汉字之间不加空格，英文断词连字符去掉，其余加一个空格。"""
    if not a:
        return b
    if not b:
        return a
    if re.search(rf"[{CJK}]$", a) and re.match(rf"[{CJK}]", b):
        return a + b
    if re.search(r"[A-Za-z]-$", a) and re.match(r"[a-z]", b):
        return a[:-1] + b
    return a + " " + b


def is_table_line(line: str) -> bool:
    return line.startswith("|")


def is_sep_line(line: str) -> bool:
    return is_table_line(line) and "-" in line and SEP_RE.fullmatch(line) is not None


def split_row(line: str):
    """按未转义的 | 切分表格行。"""
    body = line.strip()
    body = body[1:] if body.startswith("|") else body
    if body.endswith("|") and not body.endswith("\\|"):
        body = body[:-1]
    return [c.strip() for c in re.split(r"(?<!\\)\|", body)]


def clean_cell(c: str) -> str:
    """规整单元格：汉字间/结尾的 <br> 去掉，其余 <br> 变空格；结尾反斜杠补空格以免吞掉分隔符。"""
    c = re.sub(rf"(?<=[{CJK}])(?:<br\s*/?>)+(?=[{CJK}])", "", c)
    c = re.sub(r"(?:<br\s*/?>)+\s*$", "", c)
    c = re.sub(r"\s*(?:<br\s*/?>)+\s*", " ", c).strip()
    return c + " " if c.endswith("\\") else c


def render_row(cells) -> str:
    return "|" + "|".join(clean_cell(c) for c in cells) + "|"


def render_table(grid):
    """二维文本 → Markdown 表格行列表（首行作表头）。"""
    n = max(len(r) for r in grid)
    grid = [[(c or "").replace("|", "\\|") for c in r] + [""] * (n - len(r)) for r in grid]
    return [render_row(grid[0]), "|" + "|".join("---" for _ in range(n)) + "|"] + [render_row(r) for r in grid[1:]]


def normalize_tables(lines):
    """重建所有表格行：去掉 pandoc 的对齐空格，分隔行统一成 |---|---|。
    行内出现 `\\|`(实为「内容为反斜杠的单元格 + 分隔符」)时，以分隔行的列数为准判断怎么切。"""
    out, ncols = [], None
    for l in lines:
        if not is_table_line(l):
            out.append(l); ncols = None
        elif is_sep_line(l):
            ncols = len(split_row(l))
            out.append("|" + "|".join("---" for _ in range(ncols)) + "|")
        else:
            cells = split_row(l)
            if ncols and len(cells) != ncols:
                naive = [c.strip() for c in l.strip().strip("|").split("|")]
                if len(naive) == ncols:
                    cells = naive
            out.append(render_row(cells))
    return out


# 页码行：- 12 -、第 3 页 共 9 页、Page 2 of 10、3/12
PAGENUM_RE = re.compile(r"^(?:[-–—]\s*\d{1,4}\s*[-–—]|第\s*\d{1,4}\s*页(?:\s*[/，,]?\s*共\s*\d{1,4}\s*页)?"
                        r"|(?:page|p\.)\s*\d{1,4}(?:\s*(?:of|/)\s*\d{1,4})?|\d{1,4}\s*/\s*\d{1,4})$", re.I)
CAPTION_RE = re.compile(r"^[（(]?(?:(?:续表|表|图)\s*[\dA-Za-z一二三四五六七八九十]|(?:Table|Figure|Fig\.?)\s*[\dIVX]+)", re.I)
CAP_ONLY_RE = re.compile(r"^(?:续表|表|图|Table|Figure|Fig\.?)\s*[\dA-Za-z]+(?:[.\-–]\d+)*$", re.I)
CONT_RE = re.compile(r"^[（(]?续表\s*[\dA-Za-z]*(?:[.\-–]\d+)*[）)]?$|^[（(]续[）)]$")


def _plain(line: str) -> str:
    return re.sub(r"^#{1,6}\s+|\*\*", "", line).strip()


def _is_gap_line(line: str) -> bool:
    """两段表格之间可忽略的行：空行、页码行、"续表 N"行。"""
    t = _plain(line)
    return not t or bool(PAGENUM_RE.match(t)) or bool(CONT_RE.match(t.replace(" ", "")))


def merge_split_tables(lines):
    """相邻(仅隔空行/页码/"续表"行)且列数相同的表格视为跨页被切开，合并为一张。
    后一段被误当表头的首行改回数据行；若与前一段表头相同(重复表头)则丢弃。返回 (行, 合并次数)。"""
    out, i, merged = [], 0, 0
    while i < len(lines):
        if not is_table_line(lines[i]):
            out.append(lines[i]); i += 1; continue
        j = i
        while j < len(lines) and is_table_line(lines[j]):
            j += 1
        block = lines[i:j]
        while True:
            k = j
            while k < len(lines) and not is_table_line(lines[k]) and _is_gap_line(lines[k]):
                k += 1
            if k >= len(lines) or not is_table_line(lines[k]):
                break
            m = k
            while m < len(lines) and is_table_line(lines[m]):
                m += 1
            nxt = lines[k:m]
            if len(split_row(block[0])) != len(split_row(nxt[0])):
                break
            if len(nxt) >= 2 and is_sep_line(nxt[1]):
                head, rest = nxt[0], nxt[2:]
                if split_row(head) != split_row(block[0]):
                    block.append(head)
                block.extend(rest)
            else:
                block.extend(nxt)
            j = m
            merged += 1
        out.extend(block)
        i = j
    return out, merged


def merge_wrapped_cells(lines):
    """单元格内换行被拆成独立一行(只有一格有字)的，并回上一行：仅在明显是续写时合并
    (上一格括号未闭合，或本格以小写字母开头)。返回 (行, 合并数)。"""
    out, n = [], 0
    for l in lines:
        if is_table_line(l) and not is_sep_line(l) and out and is_table_line(out[-1]) and not is_sep_line(out[-1]):
            cur, prev = split_row(l), split_row(out[-1])
            filled = [k for k, c in enumerate(cur) if c]
            if len(cur) == len(prev) and len(filled) == 1:
                k = filled[0]
                p, c = prev[k], cur[k]
                if p and (p.count("(") + p.count("（") > p.count(")") + p.count("）") or re.match(r"[a-z]", c)):
                    prev[k] = join_text(p, c)
                    out[-1] = render_row(prev)
                    n += 1
                    continue
        out.append(l)
    return out, n


def repair_caption_in_header(lines):
    """layout 引擎会把紧贴表格上方的题注碎片(加粗)塞进表头单元格；拼回题注并移到表前。"""
    out, fixed, i = [], 0, 0
    while i < len(lines):
        if (is_table_line(lines[i]) and (i == 0 or not is_table_line(lines[i - 1]))
                and i + 1 < len(lines) and is_sep_line(lines[i + 1])):
            cells, parts, new = split_row(lines[i]), [], []
            for c in cells:
                m = re.match(r"^\*\*(.+?)\*\*(?:<br>)?(.*)$", c)
                if m:
                    parts.append(m.group(1)); new.append(m.group(2))
                else:
                    new.append(c)
            cap = "".join(parts)
            if parts and CAPTION_RE.match(cap):
                out.extend([f"**{cap}**", "", render_row(new)])
                fixed += 1
                i += 1
                continue
        out.append(lines[i]); i += 1
    return out, fixed


NUM_HEAD_RE = re.compile(r"^(第[^\s]{1,6}?[章节篇部]|\d+(?:\.\d+)*\.?)\s+(.+)$")


ITALIC_RE = re.compile(r"(?<![\w*])_([^_\s][^_]*)_(?!\w)")


def _clean_mark(t: str) -> str:
    """去掉加粗/下划线/删除线/斜体标记。"""
    return ITALIC_RE.sub(r"\1", re.sub(r"\*\*|__|~~|</?u>", "", t)).strip()


def max_title_len(text: str) -> int:
    """标题长度上限：中文为主 60 字，英文为主 150 字符(英文标题天然更长)。"""
    cjk = len(re.findall(rf"[{CJK_IDEO}]", text))
    return 60 if cjk >= 0.3 * max(len(text.replace(" ", "")), 1) else 150


def tidy_heading_text(text: str) -> str:
    """标题文字规整：去加粗/斜体/下划线/删除线标记，压缩中文空格，但保留"第一章 xx""1.2 xx"编号后的空格。"""
    text = _clean_mark(text)
    m = NUM_HEAD_RE.match(text)
    if m:
        return f"{m.group(1)} {squeeze_cjk_space(m.group(2))}"
    return squeeze_cjk_space(text)


def looks_like_formula(raw: str) -> bool:
    """引擎把公式里的斜体变量标成 _x_：多段且每段都很短(如 _hi_ ( _l_ ) _P_)。整句斜体的英文小节标题不算。"""
    segs = ITALIC_RE.findall(raw)
    return len(segs) >= 2 and sum(len(s) for s in segs) / len(segs) <= 4


def _not_heading(raw: str, text: str) -> bool:
    """题注、整句、公式、续表等不应是标题。"""
    return bool(CAPTION_RE.match(text) or CONT_RE.match(text.replace(" ", ""))
                or re.search(r"[：:。；;]$", text) or "。" in text or len(text) > max_title_len(text)
                or looks_like_formula(raw) or not_title_text(text)
                or (re.search(r"[（(]\d{1,3}[）)]$", text) and re.search(r"[=+×∑∫<>≤≥]|_", raw))
                or len(re.sub(r"[\W_]", "", text)) < 2)


def fix_headings(lines):
    """标题规范化：去加粗/下划线/删除线标记；图片移出标题行；题注、整句、公式被误判为标题的降级为正文。返回 (行, 降级数)。"""
    out, demoted = [], 0
    for l in lines:
        m = re.match(r"^(#{1,6}) (.*)$", l)
        if not m:
            out.append(l); continue
        imgs = [x.group(0) for x in IMG_RE.finditer(m.group(2))]
        raw = IMG_RE.sub("", m.group(2))
        text = tidy_heading_text(raw)
        if not text:
            out.extend(imgs); continue
        if _not_heading(raw, text):
            out.append(text); demoted += 1
        else:
            out.append(f"{m.group(1)} {text}")
        for im in imgs:
            out.extend(["", im, ""])
    return out, demoted


def tighten_lists(lines):
    """相邻列表项之间的空行去掉。"""
    item = re.compile(r"^\s*(?:[-*+]|\d+[.)]) ")
    out = []
    for i, l in enumerate(lines):
        if l.strip() == "" and out and item.match(out[-1]):
            nxt = next((x for x in lines[i + 1:] if x.strip()), "")
            if item.match(nxt):
                continue
        out.append(l)
    return out


def finalize(md: str):
    """两种格式共用的收尾清理。"""
    lines = [l.rstrip() for l in md.replace("\r\n", "\n").split("\n")]
    lines = [re.sub(r"\\([\[\]])", r"\1", l) for l in lines]  # pandoc 多余的括号转义
    lines = normalize_tables(lines)
    lines = tighten_lists(lines)
    text = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip() + "\n"
    return text


# ───────────────────────── 公共：标题补全 ─────────────────────────
NUM_CLASSES = [
    ("章节", r"第[一二三四五六七八九十百零\d]+[章节篇部]"),
    ("中文顿号", r"[一二三四五六七八九十]+、"),
    ("中文括号", r"[（(][一二三四五六七八九十]+[）)]"),
    ("三级数字", r"\d+\.\d+\.\d+(?![.\d])"),
    ("二级数字", r"\d+\.\d+(?![.\d])"),
    ("一级数字", r"\d+[.．、](?!\d)"),
    ("数字括号", r"[（(]\d+[）)]"),
]
PLAIN_SKIP = re.compile(r"^\s*(?:#|\||!\[|<|>|[-*+] |\d+[.)] )")


def num_class(text: str):
    for name, pat in NUM_CLASSES:
        if re.match(pat, text):
            return name
    return None


def not_title_text(text: str) -> bool:
    """作者名单、"标签：值"式的封面/表单字段，不是标题。"""
    if len(re.findall(r"[，,]", text)) >= 2:
        return True
    if re.fullmatch(rf"(?:[{CJK_IDEO}·]{{2,4}}\s*\d{{1,2}}\s*[，,]?\s*){{2,}}", text):  # 张三 1 李四 2（作者名+上标单位号）
        return True
    return bool(FIELD_RE.match(text))


# 封面/表单字段标签（"标签：值"）；"主标题：副标题"式的真标题不在此列
FIELD_RE = re.compile(
    r"^\s*(?:姓\s*名|班\s*级|学\s*号|学\s*院|专\s*业|年\s*级|系\s*别|课程(?:名称)?|任课教师|指导教师|教\s*师|"
    r"(?:论文)?题\s*目|作\s*者|单\s*位|日\s*期|时\s*间|地\s*址|电\s*话|邮\s*箱|关键[词字]|摘\s*要|"
    r"基金项目|作者简介|收稿日期|中图分类号|文献标识码|文章编号|分类号|DOI|"
    r"keywords?|key words|authors?|date|name|e-?mail|affiliation)\s*[：:]\s*\S", re.I)


def heading_like(text: str, max_len=40) -> bool:
    return (2 <= len(text) <= max_len and "。" not in text and not re.search(r"[，,；;：:、]$", text)
            and not CAPTION_RE.match(text) and len(re.sub(r"[\W_\d]", "", text)) >= 2
            and not not_title_text(text))


def _standalone(lines, i) -> bool:
    return (i == 0 or not lines[i - 1].strip()) and (i + 1 >= len(lines) or not lines[i + 1].strip())


def _key(t: str) -> str:
    """标题比对用的键：只留字母数字汉字，不区分大小写。"""
    return _norm_for_cov(_clean_mark(t)).lower()


def _toc_key(t: str) -> str:
    """与书签比对用的键：再去掉开头的编号(VIII. / 3.1 / A. / 第二章)，书签里常不带编号。"""
    t = re.sub(r"^\s*(?:第.{1,4}[章节部篇]|[IVXLC]{1,6}\.|\d+(?:\.\d+)*\.?|[A-Z]\.)\s*", "", _clean_mark(t))
    return _key(t)


def promote_headings(lines, toc=None):
    """补全漏判的标题：
    ① 独立成行、编号格式与已有标题相同(如已有"#### 2.估计结果"，则"1.方法介绍"也是同级标题)；
    ② 独立成行、与 PDF 书签(目录)条目一致。返回 (行, 按编号补标数, 按书签补标数)。"""
    lv_by_cls = defaultdict(Counter)
    for l in lines:
        m = re.match(r"^(#{1,6}) (.+)$", l)
        if m:
            c = num_class(_clean_mark(m.group(2)))
            if c:
                lv_by_cls[c][len(m.group(1))] += 1
    toc_map, offset = {}, 0
    if toc:
        toc_map = {_toc_key(t): lv for lv, t in toc if len(_toc_key(t)) >= 2}
        diffs = Counter()
        for l in lines:
            m = re.match(r"^(#{1,6}) (.+)$", l)
            if m and _toc_key(m.group(2)) in toc_map:
                diffs[len(m.group(1)) - toc_map[_toc_key(m.group(2))]] += 1
        offset = diffs.most_common(1)[0][0] if diffs else 0
    n_sib = n_toc = 0
    for i, l in enumerate(lines):
        if not l.strip() or PLAIN_SKIP.match(l) or not _standalone(lines, i):
            continue
        t, k = _clean_mark(l), _toc_key(l)
        if not heading_like(t, max_title_len(t) if k in toc_map else 40):  # 与书签一致是强证据，放宽长度
            continue
        c = num_class(t)
        if k in toc_map:
            lv = lv_by_cls[c].most_common(1)[0][0] if c in lv_by_cls else toc_map[k] + offset
            lines[i] = "#" * min(max(lv, 1), 6) + " " + tidy_heading_text(t)
            n_toc += 1
        elif c in lv_by_cls:
            lines[i] = "#" * lv_by_cls[c].most_common(1)[0][0] + " " + tidy_heading_text(t)
            n_sib += 1
    return lines, n_sib, n_toc


def merge_caption_lines(lines):
    """"表 2"单独成行、题目在下一行(常被误判为标题)时，拼成一行题注。返回 (行, 合并数)。"""
    out, i, n = [], 0, 0
    while i < len(lines):
        t = _plain(lines[i])
        if CAP_ONLY_RE.match(t):
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines):
                body = re.sub(r"^#{1,6}\s+", "", lines[j])
                nxt = _plain(lines[j])
                if nxt and not PLAIN_SKIP.match(body) and len(nxt) <= 40 and "。" not in nxt:
                    out.append(f"{t} {nxt}")
                    i, n = j + 1, n + 1
                    continue
        out.append(lines[i]); i += 1
    return out, n


def check_captions(md: str):
    """题注写着"表 N"、附近却没有表格的，列出来(可能是表格被识别成了文字，或表格本身是图片)。"""
    lines, res = md.split("\n"), []
    for i, l in enumerate(lines):
        t = _plain(l)
        if (re.match(r"^(?:表|Table)\s*[\dA-Za-z]+(?:[.\-–]\d+)*(?:\s|$|[：:.])", t, re.I)
                and len(t) <= 50 and "。" not in t and not CONT_RE.match(t.replace(" ", ""))):
            win = lines[max(0, i - 8):i] + lines[i + 1:i + 9]
            if not any(is_table_line(x) for x in win):
                img = next((m.group(2) or m.group(3) for x in win for m in IMG_RE.finditer(x)), None)
                res.append(f"「{t[:30]}」" + (f"（表格为图片：{img}）" if img else ""))
    return res


# ───────────────────────── 公共：图片重编号 ─────────────────────────
IMG_RE = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)|<img\s[^>]*?src="([^"]+)"[^>]*?/?>')


def to_png_if_needed(path: Path, workdir: Path) -> Path:
    """emf/wmf 等格式用 LibreOffice 转 png，其余原样返回。"""
    if path.suffix.lower() not in (".emf", ".wmf", ".tif", ".tiff", ".bmp"):
        return path
    outdir = workdir / "_conv"
    outdir.mkdir(exist_ok=True)
    if not run_soffice(["--convert-to", "png", "--outdir", str(outdir), str(path)], workdir, timeout=120):
        return path
    png = outdir / (path.stem + ".png")
    return png if png.exists() else path


def renumber_images(md: str, resolve, out_dir: Path, workdir: Path):
    """按首次出现顺序把图片复制为 images/1.ext、2.ext…；相同内容复用同一编号。返回 (新md, 图片数)。"""
    img_dir = out_dir / "images"
    seen = {}  # md5 -> 文件名

    def repl(m):
        alt = m.group(1) or ""
        src = unquote(m.group(2) or m.group(3)).strip()
        p = resolve(src)
        if p is None or not p.exists():
            return m.group(0)
        p = to_png_if_needed(p, workdir)
        h = hashlib.md5(p.read_bytes()).hexdigest()
        if h not in seen:
            img_dir.mkdir(parents=True, exist_ok=True)
            name = f"{len(seen) + 1}{p.suffix.lower()}"
            shutil.copyfile(p, img_dir / name)
            seen[h] = name
        return f"![{alt}](images/{seen[h]})"

    return IMG_RE.sub(repl, md), len(seen)


# ───────────────────────── 公共：文字覆盖率核对 ─────────────────────────
def _norm_for_cov(s: str) -> str:
    s = re.sub(r"</?(?:sup|sub|br|u|b|i|em|strong|span|mark)\b[^>]*>|!\[[^\]]*\]\([^)]*\)", "", s)
    return re.sub(r"[\W_]+", "", s)  # 只留字母数字与汉字


def coverage(units, md_text: str, n=6, min_ratio=0.03):
    """units: [(标签, 原文)]。以 n-gram 检查原文每段是否都出现在 md 中。
    返回 (总覆盖率, [(标签, 缺失比例, 缺失片段示例)])。"""
    b = _norm_for_cov(md_text)
    sh = {b[i:i + n] for i in range(len(b) - n + 1)}
    total = miss_total = 0
    bad = []
    for label, text in units:
        a = _norm_for_cov(text)
        cnt = len(a) - n + 1
        if cnt <= 0:
            continue
        miss = [i for i in range(cnt) if a[i:i + n] not in sh]
        total += cnt
        miss_total += len(miss)
        if len(miss) / cnt >= min_ratio:
            bad.append((label, len(miss) / cnt, a[miss[0]:miss[0] + 24]))
    return (1 - miss_total / total if total else 1.0), bad


def missing_headings(expected, md: str):
    """expected: 期望成为标题的文字列表；返回在 md 标题中找不到的(不区分大小写)。"""
    heads = [h for h in (_key(m.group(1)) for m in re.finditer(r"(?m)^#{1,6} (.+)$", md)) if h]
    miss = []
    for e in expected:
        k = _key(e)
        if k and not any(k == h or k in h or (len(h) >= 4 and h in k) for h in heads):
            miss.append(e)
    return miss


# ───────────────────────── PDF ─────────────────────────
def open_pdf(src: Path, password):
    import pymupdf
    doc = pymupdf.open(str(src))
    if doc.needs_pass:
        if not password:
            sys.exit("该 PDF 已加密，需要打开密码。请向用户索取密码后用 --password 参数重新运行。")
        if not doc.authenticate(password):
            sys.exit("PDF 密码不正确，请向用户确认密码。")
    return doc


def strip_repeated_header_footer(doc, band=0.12, min_ratio=0.3):
    """检测页眉/页脚区域里跨页重复(数字归一化后相同、纵坐标相近)的文本块并涂掉；
    页码类文本(- 12 -、第3页、单独的数字)不论纵坐标是否一致都按重复处理。返回被移除的文本样例。"""
    import pymupdf
    n = len(doc)
    if n < 3:
        return []
    cand = defaultdict(list)
    for p in doc:
        H = p.rect.height
        for b in p.get_text("blocks"):
            if b[6] != 0:
                continue
            x0, y0, x1, y1, text = b[:5]
            if not (y1 < H * band or y0 > H * (1 - band)):
                continue
            flat = " ".join(text.split())
            key = re.sub(r"\d+", "#", flat)
            if not key:
                continue
            pagenum = bool(PAGENUM_RE.match(flat) or re.fullmatch(r"\d{1,4}", flat))
            cand[(key, "页码" if pagenum else round(y0 / 6))].append((p.number, pymupdf.Rect(x0, y0, x1, y1)))
    thresh = max(3, int(n * min_ratio))
    removed = []
    for (key, _), items in cand.items():
        if len({pn for pn, _ in items}) >= thresh:
            removed.append(key)
            for pn, r in items:
                doc[pn].add_redact_annot(r)
    if removed:
        for p in doc:
            p.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                               graphics=getattr(pymupdf, "PDF_REDACT_LINE_ART_NONE", 0))
    return removed


def pdf_typography(doc):
    """逐页取视觉行(同一水平线上相邻的文本行合并，避免换字体处被拆开)及其主字号；
    正文字号 = 在长行(≥15 字)中占比 ≥15% 的字号里最大的一个。
    (只取众数不可靠：小字号的表格/脚注内容多时，会把真正的正文误判为"大字号"。)"""
    import pymupdf
    pages, size_cnt, all_cnt = [], Counter(), Counter()
    for p in doc:
        raw = []
        for b in p.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                spans = [s for s in l["spans"] if s["text"].strip()]
                if spans:
                    raw.append(dict(bbox=pymupdf.Rect(l["bbox"]), spans=spans))
        # 换字体的字(如后备字体)可能落在别的文本块里：按坐标找同一水平线上首尾相接的片段拼回去
        for i, r in enumerate(raw):
            r["order"] = i
        merged = set()
        by_y = sorted(raw, key=lambda r: ((r["bbox"].y0 + r["bbox"].y1) / 2, r["bbox"].x0))
        for i, r in enumerate(by_y):
            if id(r) in merged:
                continue
            yc, h = (r["bbox"].y0 + r["bbox"].y1) / 2, r["bbox"].height
            row = [r]
            for q in by_y[i + 1:]:
                if (q["bbox"].y0 + q["bbox"].y1) / 2 - yc > 0.3 * h:
                    break
                if id(q) not in merged:
                    row.append(q)
            row.sort(key=lambda q: q["bbox"].x0)
            cur = None
            for q in row:
                size = max(s["size"] for s in q["spans"])
                if cur is not None and -1 <= q["bbox"].x0 - cur["bbox"].x1 <= size:
                    cur["bbox"] |= q["bbox"]; cur["spans"] += q["spans"]
                    merged.add(id(q)); q["dead"] = True
                else:
                    cur = q
        lines = [dict(bbox=pymupdf.Rect(r["bbox"]), spans=list(r["spans"])) for r in raw if not r.get("dead")]
        out = []
        for l in lines:
            text = "".join(s["text"] for s in l["spans"]).strip()
            sz = Counter()
            for s in l["spans"]:
                sz[(round(s["size"] * 2) / 2, bool(s["flags"] & 2))] += len(s["text"].strip())
            (size, italic), _ = sz.most_common(1)[0]
            for (s_, _), c in sz.items():
                all_cnt[s_] += c
                if len(text) >= 15:
                    size_cnt[s_] += c
            out.append(dict(bbox=l["bbox"], text=text, size=size, italic=italic))
        pages.append(out)
    cnt = size_cnt or all_cnt
    total = sum(cnt.values())
    major = [s for s, c in cnt.items() if c >= 0.15 * total]
    return dict(pages=pages, body=max(major) if major else 10.0)


def is_size_heading(text: str, size: float, body: float, italic=False) -> bool:
    """字号明显大于正文、文字较短且不像句子的行视为标题。斜体(多为公式变量)、作者名单、表单字段除外。"""
    letters = len(re.sub(rf"[^A-Za-z{CJK_IDEO}]", "", text))
    return (size >= body + 1.5 and not italic and 2 <= len(text) <= 60 and letters >= 2
            and letters >= 0.5 * len(text.replace(" ", "")) and not re.search(r"[。；;，,：:、]$", text)
            and "。" not in text and not not_title_text(text))


def adjacent_line_pairs(typo):
    """原文中上下紧邻、字号相同的行对(规范化文字)，用于把被拆成两个的跨行标题合并回去。"""
    pairs = set()
    for lines in typo["pages"]:
        for a, b in zip(lines, lines[1:]):
            h = a["bbox"].height
            if a["size"] == b["size"] and 0 <= b["bbox"].y0 - a["bbox"].y1 < 0.6 * h:
                pairs.add((_norm_for_cov(a["text"]), _norm_for_cov(b["text"])))
    return pairs


def merge_heading_continuations(lines, pairs):
    """相邻两个同级标题，若在原文中是上下紧邻的两行(跨行长标题)，合并为一个。返回 (行, 合并数)。"""
    out, n = [], 0
    for l in lines:
        m = re.match(r"^(#{1,6}) (.+)$", l)
        if m and out:
            k = len(out) - 1
            while k >= 0 and not out[k].strip():
                k -= 1
            pm = re.match(r"^(#{1,6}) (.+)$", out[k]) if k >= 0 else None
            if pm and re.fullmatch(r"\d{1,2}|第.{1,4}[章节部篇]|[IVX]{1,4}\.?", pm.group(2).strip()):
                out[k] = f"{pm.group(1)} {pm.group(2).strip()} {m.group(2)}"  # 只有编号的标题 + 下一个标题
                del out[k + 1:]
                n += 1
                continue
            if (pm and pm.group(1) == m.group(1) and not num_class(m.group(2))
                    and any(a.endswith(_norm_for_cov(pm.group(2))[-6:]) and b.startswith(_norm_for_cov(m.group(2))[:6])
                            for a, b in pairs if a and b)):
                out[k] = pm.group(1) + " " + join_text(pm.group(2), m.group(2))
                del out[k + 1:]
                n += 1
                continue
        out.append(l)
    return out, n


def heading_levels(chunks, typo):
    """字号 → 标题级别：优先沿用引擎对同字号标题给出的级别，其余字号按大小排位。"""
    votes = defaultdict(Counter)
    for ch, lines in zip(chunks, typo["pages"]):
        idx = {_norm_for_cov(l["text"]): l["size"] for l in lines}
        for m in re.finditer(r"(?m)^(#{1,6}) (.+)$", ch["text"]):
            k = _norm_for_cov(_clean_mark(m.group(2)))
            if k in idx:
                votes[idx[k]][len(m.group(1))] += 1
    level = {s: c.most_common(1)[0][0] for s, c in votes.items()}
    sizes = sorted({l["size"] for lines in typo["pages"] for l in lines
                    if is_size_heading(l["text"], l["size"], typo["body"], l["italic"])}, reverse=True)
    for s in sizes:
        if s in level:
            continue
        bigger = [lv for x, lv in level.items() if x > s]
        smaller = [lv for x, lv in level.items() if x < s]
        lv = max(bigger) + 1 if bigger else (min(smaller) - 1 if smaller else 1)
        level[s] = min(max(lv, 1), 6)
    return level


def box_rows(page, rect):
    """取版面块内的文字，按行聚类，行内按大间距切成"单元格"。"""
    import pymupdf
    toks = []
    d = page.get_text("rawdict", clip=pymupdf.Rect(rect) + (-1, -1, 1, 1))
    for b in d["blocks"]:
        for l in b.get("lines", []):
            if abs(l.get("dir", (1, 0))[0]) < 0.9:  # 竖排/旋转文字不参与
                continue
            for s in l["spans"]:
                cur = []
                for ch in s["chars"] + [{"c": " "}]:
                    if ch["c"].isspace():
                        if cur:
                            r = pymupdf.Rect(cur[0]["bbox"])
                            for c in cur[1:]:
                                r |= c["bbox"]
                            toks.append(dict(x0=r.x0, x1=r.x1, y0=r.y0, y1=r.y1, size=s["size"],
                                             italic=bool(s["flags"] & 2), text="".join(c["c"] for c in cur)))
                            cur = []
                    else:
                        cur.append(ch)
    toks.sort(key=lambda t: ((t["y0"] + t["y1"]) / 2, t["x0"]))
    rows = []
    for t in toks:
        yc, h = (t["y0"] + t["y1"]) / 2, t["y1"] - t["y0"]
        if rows and abs(yc - rows[-1]["yc"]) <= 0.45 * max(h, rows[-1]["h"]):
            rows[-1]["toks"].append(t)
        else:
            rows.append(dict(yc=yc, h=h, toks=[t]))
    for r in rows:
        cells = []
        for t in sorted(r["toks"], key=lambda t: t["x0"]):
            if cells and t["x0"] - cells[-1]["x1"] <= max(t["size"], cells[-1]["size"]):
                c = cells[-1]
                tight = t["x0"] - c["x1"] < 0.15 * t["size"]
                c["text"] = c["text"] + t["text"] if tight else join_text(c["text"], t["text"])
                c["x1"] = max(c["x1"], t["x1"])
            else:
                cells.append(dict(x0=t["x0"], x1=t["x1"], text=t["text"], size=t["size"]))
        sz, ita = Counter(), 0
        for t in r["toks"]:
            sz[round(t["size"] * 2) / 2] += len(t["text"])
            ita += len(t["text"]) if t["italic"] else -len(t["text"])
        r.update(cells=cells, size=sz.most_common(1)[0][0], text=" ".join(c["text"] for c in cells), italic=ita > 0,
                 x0=min(t["x0"] for t in r["toks"]), y0=min(t["y0"] for t in r["toks"]),
                 y1=max(t["y1"] for t in r["toks"]))
    return rows


def table_from_rows(rows):
    """无框表格/三线表：多行、每行多个单元格且列位置对齐 → 二维文本；不像表格返回 None。"""
    multi = [r for r in rows if len(r["cells"]) >= 2]
    if len(multi) < 3:
        return None
    comps = []
    for a, b in sorted((c["x0"], c["x1"]) for r in multi for c in r["cells"]):
        if comps and a <= comps[-1][1] + 2:
            comps[-1][1] = max(comps[-1][1], b)
        else:
            comps.append([a, b])
    nc = len(comps)
    if nc < 2 or nc > 20:
        return None

    def col(c):
        return next((k for k, (a, b) in enumerate(comps) if a - 2 <= c["x0"] <= b + 2), nc - 1)

    grid, bad = [], 0
    for r in rows:
        if len(r["cells"]) >= 2 or not grid:
            row, ks = [""] * nc, [col(c) for c in r["cells"]]
            bad += len(set(ks)) < len(ks)
            for c, k in zip(r["cells"], ks):
                row[k] = join_text(row[k], c["text"])
            grid.append(row)
        else:  # 单元格内换行的续行
            c = r["cells"][0]
            grid[-1][col(c)] = join_text(grid[-1][col(c)], c["text"])
    if bad > 0.2 * len(multi):
        return None
    cells = [c for r in multi for c in r["cells"]]

    def width(s):  # 中文按 2、其他按 1 计
        return len(s) + len(re.findall(rf"[{CJK}]", s))

    if max(width(c["text"]) for c in cells) > 100 or sum(width(c["text"]) for c in cells) / len(cells) > 40:
        return None  # 单元格太长：多栏正文/参考文献，不是无框表格
    full = sum(1 for c in cells if c["x1"] - c["x0"] >= 0.9 * (comps[col(c)][1] - comps[col(c)][0]))
    if full / len(cells) >= 0.8 and sum(len(c["text"]) for c in cells) / len(cells) >= 12:
        return None  # 多栏正文，不是表格
    marker = r"[\[(（]?\d{1,3}[\].)）、:：]?|[•·●○■◆▪\-–]"
    if sum(1 for r in multi if re.fullmatch(marker, r["cells"][0]["text"]) or re.match(r"\[\d{1,3}\]", r["cells"][0]["text"])) >= 0.5 * len(multi):
        return None  # 带编号的列表、参考文献、带行号的代码/算法
    return grid


def rebuild_text_boxes(page, chunk, skip, typo, levels):
    """引擎判为正文的块里若藏着标题(大字号)或无框表格，按字号/坐标重建该块。
    返回 (替换表 {块号: 文本}, 重建表格数, 补出标题数, 本页期望成为标题的行)。"""
    repl, ntab, nhead, expected = {}, 0, 0, []
    body = typo["body"]
    for b in chunk["page_boxes"]:
        if b["class"] in ("section-header", "title", "text", "list-item"):
            prev, prev_ok = None, False
            for l in typo["pages"][page.number]:
                if _cover(l["bbox"], b["bbox"]) < 0.6:
                    continue
                cont = (prev is not None and prev["size"] == l["size"]
                        and 0 <= l["bbox"].y0 - prev["bbox"].y1 < 0.6 * l["bbox"].height)
                if cont:  # 同字号紧邻的下一行：跟随上一行(上一行是标题则并入，不是则也不算)
                    if prev_ok:
                        expected[-1] = join_text(expected[-1], l["text"])
                elif l["size"] >= body + 1.5:
                    prev_ok = is_size_heading(l["text"], l["size"], body, l["italic"])
                    if prev_ok:
                        expected.append(l["text"])
                else:
                    prev_ok = False
                prev = l if l["size"] >= body + 1.5 else None
        if b["index"] in skip or b["class"] not in ("text", "list-item"):
            continue
        rows = box_rows(page, b["bbox"])
        if not rows:
            continue
        is_head = [is_size_heading(r["text"], r["size"], body, r["italic"]) for r in rows]
        big = [r["size"] >= body + 1.5 for r in rows]
        for k in range(1, len(rows)):  # 被排除的大字号行(如封面字段)的同字号续行，也不是标题
            if (is_head[k] and big[k - 1] and not is_head[k - 1] and rows[k - 1]["size"] == rows[k]["size"]
                    and rows[k]["y0"] - rows[k - 1]["y1"] < 0.6 * rows[k]["h"]):
                is_head[k] = False
        # 找连续的表格行段
        runs, i, n = [], 0, len(rows)
        while i < n:
            if len(rows[i]["cells"]) >= 2 and not is_head[i]:
                j = i + 1
                while j < n and not is_head[j]:
                    if len(rows[j]["cells"]) >= 2:
                        j += 1
                    elif (len(rows[j]["cells"]) == 1 and rows[j]["cells"][0]["x0"] > rows[i]["cells"][0]["x1"]
                          and len(rows[j]["text"]) <= 40):
                        j += 1
                    else:
                        break
                grid = table_from_rows(rows[i:j])
                if grid:
                    runs.append((i, j, grid))
                i = j
            else:
                i += 1
        if not runs and not any(is_head):
            continue
        out, para, prev, i = [], "", None, 0
        run_at = {s: (e, g) for s, e, g in runs}

        def flush():
            nonlocal para
            if para:
                out.append(para)
            para = ""

        while i < n:
            r = rows[i]
            if i in run_at:
                flush()
                e, g = run_at[i]
                out.append("\n".join(render_table(g)))
                ntab += 1
                i, prev = e, None
                continue
            if is_head[i]:
                flush()
                if (i > 0 and is_head[i - 1] and out and out[-1].startswith("#") and rows[i - 1]["size"] == r["size"]
                        and r["y0"] - rows[i - 1]["y1"] < 0.6 * r["h"]):
                    out[-1] = join_text(out[-1], r["text"])  # 跨行的长标题
                else:
                    out.append("#" * levels.get(r["size"], 2) + " " + r["text"])
                    nhead += 1
                i, prev = i + 1, None
                continue
            if prev is not None and (r["y0"] - prev["y1"] > 0.8 * r["h"] or r["x0"] > prev["x0"] + 1.5 * r["size"]):
                flush()
            para = join_text(para, r["text"])
            prev = r
            i += 1
        flush()
        repl[b["index"]] = "\n\n".join(out) + "\n\n"
    return repl, ntab, nhead, expected


def cell_text(page, rect) -> str:
    """按坐标取单元格内的词并重组：同一行按 x 排序，汉字之间不加空格，其余按间距判断。"""
    import pymupdf
    r = pymupdf.Rect(rect) + (1, 1, -1, -1)
    if r.is_empty:
        return ""
    words = page.get_text("words", clip=r)
    if not words:
        return ""
    words.sort(key=lambda w: ((w[1] + w[3]) / 2, w[0]))
    lines = []
    for w in words:
        yc, h = (w[1] + w[3]) / 2, w[3] - w[1]
        if lines and abs(yc - lines[-1][0]) <= 0.5 * max(h, lines[-1][1]):
            lines[-1][2].append(w)
        else:
            lines.append([yc, h, [w]])
    out = []
    for _, _, ws in lines:
        ws.sort(key=lambda w: w[0])
        s = ws[0][4]
        for a, b in zip(ws, ws[1:]):
            s += ("" if b[0] - a[2] < 0.15 * (b[3] - b[1]) else " ") + b[4]
        out.append(s)
    res = out[0]
    for s in out[1:]:
        join_tight = re.search(rf"[{CJK}]$", res) and re.match(rf"[{CJK}]", s)
        res += ("" if join_tight else " ") + s
    return res.replace("|", "\\|")


def _column_edges(t, min_ratio=0.6, tol=3.0):
    """取「在多数行里都出现」的竖线 x 坐标作为真正的列边界(忽略表头两侧的小边距格、合并格造成的杂线)。"""
    per_row = []
    for r in t.rows:
        xs = []
        for c in r.cells:
            if c:
                xs += [c[0], c[2]]
        per_row.append(xs)
    allx = sorted(x for xs in per_row for x in xs)
    clusters = []
    for x in allx:
        if clusters and x - clusters[-1][-1] <= tol:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    n = len(per_row)
    edges = []
    for cl in clusters:
        lo, hi = cl[0] - tol, cl[-1] + tol
        cnt = sum(1 for xs in per_row if any(lo <= x <= hi for x in xs))
        if cnt >= max(2, min_ratio * n) or n == 1:
            edges.append(sum(cl) / len(cl))
    return edges


def table_rows(page, t):
    """find_tables 的表格 → 二维文本。按「多数行共有的竖线」重新划分列，单元格按左边界归列，
    同列多个碎格文字合并；整列为空的列去掉。"""
    edges = _column_edges(t)
    if len(edges) < 3:  # 少于 2 列，退回逐格读取
        return []
    rows = []
    for r in t.rows:
        row = [""] * (len(edges) - 1)
        for c in r.cells:
            if not c:
                continue
            j = max(k for k in range(len(edges) - 1) if edges[k] <= c[0] + 3) if c[0] + 3 >= edges[0] else 0
            txt = cell_text(page, c)
            if txt:
                row[j] = (row[j] + " " + txt).strip()
        rows.append(row)
    keep = [j for j in range(len(edges) - 1) if any(r[j] for r in rows)]
    rows = [[r[j] for j in keep] for r in rows]
    return [r for r in rows if any(r)]


def find_ruled_tables(doc):
    """临时关闭 pymupdf 的 layout 钩子后再抽表：layout 一旦激活(pymupdf4llm 导入即激活)，
    find_tables 会被模型结果干扰(例如把上方正文段落并进表格)，规则式结果才是稳定的。"""
    import pymupdf
    saved = getattr(pymupdf, "_get_layout", None)
    pymupdf._get_layout = None
    try:
        return _find_ruled_tables(doc)
    finally:
        pymupdf._get_layout = saved


def _find_ruled_tables(doc):
    """用规则式 find_tables 抽取有边框的表格。返回 {页号: [(Rect, Markdown行列表, 单元格填充率)]}。
    单行表仅当紧邻的下一张表列数相同(跨页孤儿表头)时才收；其余单行框可能是矢量示意图，不接管。"""
    import pymupdf
    found = []
    for p in doc:
        for t in p.find_tables().tables:
            rows = table_rows(p, t)
            if rows and len(rows[0]) >= 2:
                found.append((p.number, pymupdf.Rect(t.bbox), rows))
    result = defaultdict(list)
    for i, (pn, bbox, rows) in enumerate(found):
        ok = len(rows) >= 2
        if not ok and i + 1 < len(found):
            nxt = found[i + 1]
            ok = nxt[0] - pn <= 1 and len(nxt[2][0]) == len(rows[0])
        if not ok:
            continue
        n = len(rows[0])
        md = [render_row(rows[0]), "|" + "|".join("---" for _ in range(n)) + "|"]
        md += [render_row(r + [""] * (n - len(r))) for r in rows[1:]]
        fill = sum(1 for r in rows for c in r if c) / (len(rows) * n)
        result[pn].append((bbox, md, fill))
    return result


def _cover(inner, outer) -> float:
    """inner 有多大比例落在 outer 内。"""
    import pymupdf
    a = pymupdf.Rect(inner)
    if a.is_empty:
        return 0.0
    i = a & pymupdf.Rect(outer)
    return 0.0 if i.is_empty else i.get_area() / a.get_area()


def splice_tables(page, chunk, tabs):
    """把引擎某页文本中与规则式表格重合的版面块，替换成规则式表格。
    共用同一引擎块的表格/被引擎一并框进去的题注、说明文字，按 y 坐标统一排序输出。
    返回 (替换表 {块号: 文本}, 接管的表格数)。"""
    boxes = chunk["page_boxes"]
    img_boxes = [i["bbox"] for i in page.get_image_info()]
    # 1. 每张表 → 它覆盖的引擎块
    assign = {}
    for ti, (rect, md, fill) in enumerate(tabs):
        cov = [b for b in boxes
               if _cover(b["bbox"], rect) >= 0.5 or (b["class"] != "picture" and _cover(rect, b["bbox"]) >= 0.5)]
        if not cov:
            continue
        # 引擎判为 picture 的区域：含嵌入位图，或表格内容稀疏(多半是矢量示意图)时保留为图；
        # 否则只是被误判成图的表格，由规则式结果接管
        if any(b["class"] == "picture" for b in cov):
            if any(_cover(rect, ib) >= 0.3 or _cover(ib, rect) >= 0.3 for ib in img_boxes) or fill < 0.6:
                continue
        assign[ti] = [b["index"] for b in cov]
    if not assign:
        return {}, 0
    # 2. 共用引擎块的表格并成一组
    groups = []
    for ti, idxs in assign.items():
        merged = [g for g in groups if g["boxes"] & set(idxs)]
        g = {"tables": [ti], "boxes": set(idxs)}
        for m in merged:
            g["tables"] += m["tables"]; g["boxes"] |= m["boxes"]; groups.remove(m)
        groups.append(g)
    by_index = {b["index"]: b for b in boxes}
    all_rects = [t[0] for t in tabs]
    repl = {}
    for g in groups:
        region = [by_index[i]["bbox"] for i in g["boxes"]]
        items = [(tabs[ti][0].y0, "\n".join(tabs[ti][1])) for ti in g["tables"]]
        for blk in page.get_text("blocks"):
            if blk[6] != 0:
                continue
            cx, cy = (blk[0] + blk[2]) / 2, (blk[1] + blk[3]) / 2
            inside_box = any(r[0] <= cx <= r[2] and r[1] <= cy <= r[3] for r in region)
            in_table = any(rc.x0 - 2 <= cx <= rc.x1 + 2 and rc.y0 - 2 <= cy <= rc.y1 + 2 for rc in all_rects)
            if inside_box and not in_table:
                items.append((blk[1], " ".join(blk[4].split())))
        items.sort(key=lambda x: x[0])
        first = min(g["boxes"])
        for i in g["boxes"]:
            repl[i] = "\n\n".join(t for _, t in items) + "\n\n" if i == first else ""
    return repl, len(assign)


def apply_repl(chunk, repl) -> str:
    """按版面块在页文本中的位置做替换；引擎个别块(如 title)结尾漏了换行，会和下一块粘连，这里补上。"""
    text, out, pos = chunk["text"], [], 0
    for b in sorted(chunk["page_boxes"], key=lambda b: b["pos"][0]):
        s, e = b["pos"]
        if b["index"] in repl:
            out.append(text[pos:s]); out.append(repl[b["index"]]); pos = e
        elif e > s and not text[s:e].endswith("\n"):
            out.append(text[pos:e]); out.append("\n\n"); pos = e
    out.append(text[pos:])
    return "".join(out)


PIC_TXT = re.compile(r"<!-- Start of picture text -->\s*(.*?)\s*<!-- End of picture text -->", re.S)


def tidy_picture_text(md: str) -> str:
    """引擎给每张图附带「图内文字」块：含表格的是被误框成图的表格，直接展开；
    真插图里的文字压成单行注释(图片本身已保留，这里只为可检索)。"""
    def repl(m):
        body = m.group(1).strip()
        if not body:
            return ""
        if re.search(r"(?m)^\|", body):
            return "\n\n" + body + "\n\n"
        flat = " ".join(re.sub(r"<br\s*/?>", " ", body).split())
        return f"<!-- 图内文字：{flat} -->"
    return PIC_TXT.sub(repl, md)


def clean_pdf_md(md: str, keep_toc: bool, toc, pairs=frozenset(), big=frozenset()):
    report = {}
    md = tidy_picture_text(md)
    lines = normalize_tables(md.replace("\r\n", "\n").split("\n"))
    # 1. 目录：带点线的表格行整体移除（标题结构已保留同样的信息）
    if not keep_toc:
        toc_rows = [i for i, l in enumerate(lines) if is_table_line(l) and re.search(r"[.…]{5,}", l)]
        if toc_rows:
            drop = set(toc_rows)
            drop |= {i + 1 for i in toc_rows if i + 1 < len(lines) and is_sep_line(lines[i + 1])}
            lines = [l for i, l in enumerate(lines) if i not in drop]
            for i, l in enumerate(lines):  # 目录标题若后面已无内容也移除
                if re.match(r"^#+\s*目录\s*$", l):
                    rest = [x for x in lines[i + 1:i + 6] if x.strip()]
                    if not rest or re.match(r"^#+ ", rest[0]):
                        lines[i] = ""
            report["移除目录行"] = len(toc_rows)
    # 2. 标题规范化
    lines, demoted = fix_headings(lines)
    report["降级的伪标题(题注/整句/公式)"] = demoted
    lines, nmerge = merge_heading_continuations(lines, pairs)
    report["跨行标题合并"] = nmerge
    # 3. 其他行：去 <u>、伪删除线(分数线/表格线压在文字上被误判)、中文空格、页码行
    strike, pagenum = 0, 0
    for i, l in enumerate(lines):
        if re.match(r"^#{1,6} ", l):
            continue
        if PAGENUM_RE.match(_plain(l)):
            lines[i] = ""; pagenum += 1; continue
        strike += len(re.findall(r"~~[^~]+~~", l))
        lines[i] = squeeze_cjk_space(re.sub(r"</?u>|~~", "", l))
    report["去除的伪删除线"] = strike
    report["清除的页码行"] = pagenum
    # 4. 题注归位 + 跨页表格合并
    lines, cap = merge_caption_lines(lines)
    lines, fixed = repair_caption_in_header(lines)
    report["题注归位(拆开的题注/表头里的题注)"] = cap + fixed
    lines, _ = fix_empty_table_headers(lines)
    lines, ndec = decorated_headings(lines, big)
    report["装饰性章标题(单行表格排版)还原"] = ndec
    lines, merged = merge_split_tables(lines)
    report["跨页表格/续表合并"] = merged
    lines, wrapped = merge_wrapped_cells(lines)
    report["单元格换行被拆成独立行(已并回)"] = wrapped
    # 5. 补全漏判的标题
    lines, n_sib, n_toc = promote_headings(lines, toc)
    report["补出的标题(按同级编号/按书签)"] = f"{n_sib}/{n_toc}"
    return "\n".join(lines), report


def decorated_headings(lines, big=None):
    """单行"表格"[编号 | 标题]实为装饰性章标题(Word 里常用一行两格的表格排版)：还原为标题。
    标题格的判据：PDF 为大字号(big 集合)，DOCX 为整格加粗(big=None)。返回 (行, 还原数)。"""
    lv = 2 if any(re.match(r"^# ", l) for l in lines) else 1
    out, i, n = [], 0, 0
    while i < len(lines):
        if not is_table_line(lines[i]):
            out.append(lines[i]); i += 1; continue
        j = i
        while j < len(lines) and is_table_line(lines[j]):
            j += 1
        rows = [split_row(l) for l in lines[i:j] if not is_sep_line(l)]
        if len(rows) == 1 and 2 <= len(rows[0]) <= 3:
            cells = [c for c in rows[0] if c]
            nums = [c for c in cells if re.fullmatch(r"\d{1,2}|第.{1,4}[章节部篇]|[IVX]{1,4}", _clean_mark(c))]
            rest = " ".join(c for c in cells if c not in nums)
            strong = (_key(rest) in big) if big is not None else bool(re.fullmatch(r"\*\*[^*]+\*\*", rest.strip()))
            if len(nums) == 1 and heading_like(_clean_mark(rest)) and strong:
                out.append(f"{'#' * lv} {_clean_mark(nums[0])} {tidy_heading_text(rest)}")
                i, n = j, n + 1
                continue
        out.extend(lines[i:j]); i = j
    return out, n


def render_pages(doc, pns, out_dir: Path, dpi: int):
    """把指定页(1 起)渲染为 pages/pN.png，返回相对路径表。"""
    d = out_dir / "pages"
    d.mkdir(parents=True, exist_ok=True)
    res = {}
    for pn in sorted(set(pns)):
        f = d / f"p{pn}.png"
        if not f.exists():
            doc[pn - 1].get_pixmap(dpi=dpi).save(str(f))
        res[pn] = f"pages/p{pn}.png"
    return res


def convert_pdf(src: Path, out_dir: Path, args):
    try:
        import pymupdf  # noqa: F401
        import pymupdf4llm
    except ImportError as e:
        sys.exit(f"缺少 PDF 依赖（{e.name}）。请安装：\n"
                 f'  python -m pip install "pymupdf=={PYMUPDF_PIN}" "pymupdf4llm=={PYMUPDF_PIN}" "pymupdf-layout=={PYMUPDF_PIN}"')
    v = _pkg_version("pymupdf4llm") or ""
    if v != PYMUPDF_PIN:
        print(f"⚠ pymupdf4llm 版本为 {v}，本脚本只在 {PYMUPDF_PIN} 上验证过；若报错请安装该版本。", file=sys.stderr)
    doc = open_pdf(src, args.password)
    orig = open_pdf(src, args.password)  # 不做任何改动的副本：表格识别、整页渲染用
    report = {"页数": len(doc)}

    # 无文字层的页(扫描件、整页图片、文字转曲)：直接给整页图
    notext = [p.number + 1 for p in doc
              if len(p.get_text().strip()) < 20 and (p.get_images() or len(p.get_drawings()) > 30)]

    removed = [] if args.keep_header_footer else strip_repeated_header_footer(doc)
    report["移除的重复页眉页脚"] = removed or "无"

    # 先记下原文(供最后核对)，再让引擎转换；有边框的表格改用规则式结果覆盖引擎的表格块
    src_lines = [(p.number + 1, l) for p in doc for l in p.get_text().split("\n")]
    # 表格必须在「未抹页眉页脚」的原始文档上识别：抹除会改变矢量线几何，导致表格被切开或吞进正文
    ruled = find_ruled_tables(orig)
    typo = pdf_typography(doc)
    toc = [(lv, t) for lv, t, _ in orig.get_toc()]

    work = Path(tempfile.mkdtemp(prefix="doc2md_"))
    chunks = pymupdf4llm.to_markdown(doc, page_chunks=True, write_images=True,
                                     image_path=str(work / "raw"), image_format="png", dpi=args.dpi)
    levels = heading_levels(chunks, typo)
    texts, spliced, rebuilt_tab, rebuilt_head, expected = [], 0, 0, 0, []
    for ch, page in zip(chunks, doc):
        if page.number + 1 in notext:
            texts.append(f"〔DOC2MD_PAGE_{page.number + 1}〕")
            continue
        repl, n = splice_tables(page, ch, ruled.get(page.number, []))
        repl2, nt, nh, exp = rebuild_text_boxes(page, ch, set(repl), typo, levels)
        spliced, rebuilt_tab, rebuilt_head = spliced + n, rebuilt_tab + nt, rebuilt_head + nh
        expected += exp
        texts.append(apply_repl(ch, {**repl2, **repl}))
    md = "\n\n".join(texts)
    report["有框表格(找到/接管)"] = f"{sum(len(v) for v in ruled.values())}/{spliced}"
    report["无框表格/三线表(引擎漏判后重建)"] = rebuilt_tab
    report["大字号标题(引擎漏判后补出)"] = rebuilt_head
    big = {_key(l["text"]) for ls in typo["pages"] for l in ls if l["size"] >= typo["body"] + 1.5}
    md, r2 = clean_pdf_md(md, args.keep_toc, toc, adjacent_line_pairs(typo), big)
    report.update(r2)
    md, nimg = renumber_images(md, lambda s: Path(s), out_dir, work)
    if notext:
        pages = render_pages(orig, notext, out_dir, args.dpi)
        md = re.sub(r"〔DOC2MD_PAGE_(\d+)〕",
                    lambda m: f"![第{m.group(1)}页原貌（本页无文字层，内容见图）]({pages[int(m.group(1))]})", md)
    md = finalize(md)

    # 逐行精确核对：原文每一行(去标点空白后)都应能在 md 中找到
    b = _norm_for_cov(md)
    total, missing = 0, []
    for pg, line in src_lines:
        if pg in notext or (not args.keep_toc and re.search(r"[.…]{5,}", line)):
            continue
        a = _norm_for_cov(line)
        if len(a) < 5:
            continue
        total += 1
        if a not in b:
            missing.append((pg, line.strip()))
    cov = (f"{total}/{len(missing)}  覆盖率 {100 * (1 - len(missing) / max(total, 1)):.2f}%" if total
           else "0/0（没有可核对的文字层）")
    if notext:
        cov += f"；另有 {len(notext)} 页无文字层，不在核对范围内，已附整页图"
    report["逐行核对(原文行数/缺失行数)"] = cov
    if missing:
        pages = render_pages(orig, [pg for pg, _ in missing], out_dir, args.dpi)
        byp = defaultdict(list)
        for pg, line in missing:
            byp[pg].append(line)
        report["缺失或乱序的行(已附整页图，涉及时请查看)"] = [
            f"第{pg}页 {len(ls)} 行 → {pages[pg]}  例：「{ls[0][:30]}」" for pg, ls in sorted(byp.items())][:15]
    else:
        report["缺失或乱序的行"] = "无"
    report["无文字层的页(已附整页图)"] = notext or "无"

    # 结构核对
    miss_big = missing_headings(expected, md)
    report["大字号标题核对(原文/未成为标题)"] = f"{len(expected)}/{len(miss_big)}" + (f"  {miss_big[:8]}" if miss_big else "")
    if toc:
        miss_toc = missing_headings([t for _, t in toc], md)
        report["PDF书签核对(书签数/未对应到标题)"] = f"{len(toc)}/{len(miss_toc)}" + (f"  {miss_toc[:8]}" if miss_toc else "")
    caps = check_captions(md)
    report["题注提到却未输出的表格"] = caps[:10] or "无"
    shutil.rmtree(work, ignore_errors=True)
    return md, nimg, report


# ───────────────────────── DOCX ─────────────────────────
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS_C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_DGM = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _attr(el, name):
    return None if el is None else el.get(W + name)


def _cn_num(n: int) -> str:
    d = "零一二三四五六七八九"
    if n < 10:
        return d[n]
    if n < 20:
        return "十" + (d[n % 10] if n % 10 else "")
    if n < 100:
        return d[n // 10] + "十" + (d[n % 10] if n % 10 else "")
    return str(n)


def _fmt_num(n: int, fmt: str) -> str:
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "upperLetter":
        return chr(64 + (n - 1) % 26 + 1)
    if fmt == "lowerLetter":
        return chr(96 + (n - 1) % 26 + 1)
    if fmt in ("upperRoman", "lowerRoman"):
        vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
        s, x = "", n
        for v, r in vals:
            while x >= v:
                s += r; x -= v
        return s if fmt == "upperRoman" else s.lower()
    if fmt.startswith("chinese") or fmt.startswith("ideograph"):
        return _cn_num(n)
    return str(n)


CN_DIGIT = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}


def docx_heading_labels(docx_path: Path):
    """按文档顺序返回标题段落 [dict(text, label, level, builtin)]：
    builtin=True 为 Word 内置标题样式(pandoc 已识别)；其余来自"X级"样式名或目录(TOC)书签，需要补标。
    label 为 Word 自动编号(pandoc 会丢掉)。"""
    z = zipfile.ZipFile(docx_path)
    styles = ET.fromstring(z.read("word/styles.xml"))
    doc = ET.fromstring(z.read("word/document.xml"))
    numbering = ET.fromstring(z.read("word/numbering.xml")) if "word/numbering.xml" in z.namelist() else None

    sty = {}  # styleId -> 属性
    for s in styles.iter(W + "style"):
        ppr = s.find(W + "pPr")
        num = ppr.find(W + "numPr") if ppr is not None else None
        sty[s.get(W + "styleId")] = dict(
            name=(_attr(s.find(W + "name"), "val") or ""),
            based=_attr(s.find(W + "basedOn"), "val"),
            outline=_attr(ppr.find(W + "outlineLvl"), "val") if ppr is not None else None,
            numId=_attr(num.find(W + "numId"), "val") if num is not None else None,
            ilvl=_attr(num.find(W + "ilvl"), "val") if num is not None else None,
        )

    def resolve_style(sid, key):
        seen = set()
        while sid and sid in sty and sid not in seen:
            seen.add(sid)
            if sty[sid][key] is not None:
                return sty[sid][key]
            sid = sty[sid]["based"]
        return None

    def builtin_level(sid):
        """内置标题样式(Heading N / 标题 N，或大纲级别 0-8)的级别；不是返回 None。"""
        s0, seen = sid, set()
        while sid and sid in sty and sid not in seen:
            seen.add(sid)
            m = re.match(r"^(?:heading|标题)\s*(\d)", sty[sid]["name"], re.I)
            if m:
                return int(m.group(1))
            sid = sty[sid]["based"]
        ol = resolve_style(s0, "outline")
        return int(ol) + 1 if ol is not None and ol.isdigit() and int(ol) <= 8 else None

    def custom_level(sid):
        """自定义样式名里的级别：如"论文一级""2级标题""标题二"。"""
        name = sty.get(sid, {}).get("name", "") if sid else ""
        m = re.search(r"([一二三四五六1-6])\s*级", name) or re.search(r"(?:标题|heading)\s*([一二三四五六])", name, re.I)
        if not m:
            return None
        g = m.group(1)
        return CN_DIGIT.get(g) or int(g)

    abstract, num2abs, num_start = {}, {}, {}
    if numbering is not None:
        for an in numbering.iter(W + "abstractNum"):
            lv = {}
            for l in an.iter(W + "lvl"):
                lv[int(l.get(W + "ilvl"))] = dict(
                    fmt=_attr(l.find(W + "numFmt"), "val") or "decimal",
                    text=_attr(l.find(W + "lvlText"), "val") or "",
                    start=int(_attr(l.find(W + "start"), "val") or 1))
            abstract[an.get(W + "abstractNumId")] = lv
        for n in numbering.iter(W + "num"):
            num2abs[n.get(W + "numId")] = _attr(n.find(W + "abstractNumId"), "val")
            for ov in n.iter(W + "lvlOverride"):
                so = ov.find(W + "startOverride")
                if so is not None:
                    num_start[(n.get(W + "numId"), int(ov.get(W + "ilvl")))] = int(so.get(W + "val"))

    body = doc.find(W + "body")
    # 目录段落(样式 toc N)里的书签名 → 级别
    toc_lv = {}
    for p in body.iter(W + "p"):
        ppr = p.find(W + "pPr")
        sid = _attr(ppr.find(W + "pStyle"), "val") if ppr is not None else None
        m = re.match(r"^toc\s*(\d)$", sty.get(sid, {}).get("name", "") if sid else "", re.I)
        if m:
            for a in re.findall(r"_Toc\d+", ET.tostring(p, encoding="unicode"))[:1]:
                toc_lv.setdefault(a, int(m.group(1)))

    # 手工排版的标题(不用样式，只加粗/放大字号)：作为没有任何样式标题时的兜底
    def style_rpr(sid, tag):
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            s = styles.find(f"{W}style[@{W}styleId='{sid}']")
            if s is None:
                break
            el = s.find(f"{W}rPr/{W}{tag}")
            if el is not None:
                return el
            sid = sty.get(sid, {}).get("based")
        return None

    dflt = styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}sz")
    default_sz = int(dflt.get(W + "val")) if dflt is not None else 21
    in_box = {id(p) for t in list(body.iter(W + "tbl")) + list(body.iter(W + "txbxContent")) for p in t.iter(W + "p")}

    def para_format(p, sid):
        """(字号半磅, 是否整段加粗, 字数)"""
        sz_cnt, bold_chars, n = Counter(), 0, 0
        for r in p.iter(W + "r"):
            t = "".join(x.text or "" for x in r.iter(W + "t")).strip()
            if not t:
                continue
            rpr = r.find(W + "rPr")
            sz = rpr.find(W + "sz") if rpr is not None else None
            sz = sz if sz is not None else style_rpr(sid, "sz")
            b = rpr.find(W + "b") if rpr is not None else None
            b = b if b is not None else style_rpr(sid, "b")
            sz_cnt[int(sz.get(W + "val")) if sz is not None else default_sz] += len(t)
            if b is not None and b.get(W + "val") not in ("0", "false"):
                bold_chars += len(t)
            n += len(t)
        return (sz_cnt.most_common(1)[0][0] if sz_cnt else default_sz), (n > 0 and bold_chars == n), n

    counters = defaultdict(dict)  # abstractId -> {ilvl: 当前值}
    started = set()
    result, fmt_cand, body_sz = [], [], Counter()
    for p in body.iter(W + "p"):
        ppr = p.find(W + "pPr")
        sid = _attr(ppr.find(W + "pStyle"), "val") if ppr is not None else None
        blv = builtin_level(sid) if sid else None
        level = blv or custom_level(sid)
        if level is None:
            marks = [b.get(W + "name") for b in p.iter(W + "bookmarkStart")]
            level = next((toc_lv[a] for a in marks if a in toc_lv), None)
        if level is None:
            if id(p) not in in_box:
                text = "".join(t.text or "" for t in p.iter(W + "t")).strip()
                sz, bold, n = para_format(p, sid)
                body_sz[sz] += n if n >= 15 else 0
                if heading_like(text) and not re.search(r"[：:]", text):
                    fmt_cand.append(dict(text=text, sz=sz, bold=bold))
            continue
        text = "".join(t.text or "" for t in p.iter(W + "t")).strip()
        if not text:
            continue
        numpr = ppr.find(W + "numPr") if ppr is not None else None
        numId = _attr(numpr.find(W + "numId"), "val") if numpr is not None else resolve_style(sid, "numId")
        ilvl = _attr(numpr.find(W + "ilvl"), "val") if numpr is not None else resolve_style(sid, "ilvl")
        label = None
        if numId and numId != "0" and numId in num2abs and num2abs[numId] in abstract:
            aid, lv = num2abs[numId], abstract[num2abs[numId]]
            il = int(ilvl or 0)
            if il in lv and lv[il]["fmt"] != "bullet":
                cnt = counters[aid]
                if (numId, il) in num_start and (numId, il) not in started:
                    cnt[il] = num_start[(numId, il)] - 1
                started.add((numId, il))
                cnt[il] = cnt.get(il, lv[il]["start"] - 1) + 1
                for deeper in [k for k in cnt if k > il]:
                    del cnt[deeper]
                label = lv[il]["text"]
                for k in range(il + 1):
                    if k in lv:
                        label = label.replace(f"%{k + 1}", _fmt_num(cnt.get(k, lv[k]["start"]), lv[k]["fmt"]))
                label = label.strip()
        result.append(dict(text=text, label=label, level=min(level, 6),
                           builtin=bool(re.match(r"^(?:heading|标题)\s*\d", sty.get(sid, {}).get("name", ""), re.I))))
    if not result and fmt_cand and body_sz:
        body = body_sz.most_common(1)[0][0]
        # 字号比正文大(≥1 磅)，或与正文同字号但整段加粗且带编号
        cands = [c for c in fmt_cand if c["sz"] >= body + 2 or (c["bold"] and c["sz"] >= body and num_class(c["text"]))]
        sizes = sorted({c["sz"] for c in cands}, reverse=True)
        for c in cands:
            result.append(dict(text=c["text"], label=None, level=min(sizes.index(c["sz"]) + 1, 6), builtin=False))
    return result


def _norm(s: str) -> str:
    return re.sub(r"[\s*_\\`]|</?u>|<[^>]+>", "", s)


ANCHOR_RE = re.compile(r'<span id="[^"]*" class="anchor"></span>')
TOC_LINE_RE = re.compile(r"^\[.*\]\(#_Toc\d+\)\s*$")


def strip_docx_toc(md: str, keep_toc: bool):
    """pandoc 把 Word 目录输出为指向书签的链接行：默认删除(标题结构已有同样信息)，--keep-toc 时转为纯文字。返回 (md, 删除行数)。"""
    lines, n = md.split("\n"), 0
    toc_idx = [i for i, l in enumerate(lines) if TOC_LINE_RE.match(l.strip())]
    if not toc_idx:
        return md, 0
    if keep_toc:
        for i in toc_idx:
            lines[i] = re.sub(r"\[([^\[\]]*)\]\(#_Toc\d+\)", r"\1", lines[i])
            lines[i] = re.sub(r"\[([^\[\]]*)\]\(#_Toc\d+\)", r"\1", lines[i])
        return "\n".join(lines), 0
    drop = set(toc_idx)
    first = toc_idx[0]
    for i in range(first - 1, max(-1, first - 6), -1):  # 目录前的"目 录"标题一并删除
        if lines[i].strip():
            if _norm_for_cov(_clean_mark(re.sub(r"^#+\s*", "", ANCHOR_RE.sub("", lines[i])))).lower() in ("目录", "contents", "tableofcontents"):
                drop.add(i)
            break
    return "\n".join(l for i, l in enumerate(lines) if i not in drop), len(toc_idx)


def promote_docx_headings(md: str, items):
    """自定义样式/目录书签识别出的标题段落，在 pandoc 输出中是普通段落：按顺序 + 文本匹配补成标题。返回 (md, 补标数)。"""
    lines, j, n = md.split("\n"), 0, 0
    for i, l in enumerate(lines):
        is_head = re.match(r"^#{1,6} ", l)
        if not is_head and (not l.strip() or not _standalone(lines, i) or is_table_line(l) or IMG_RE.search(l)):
            continue
        plain = _norm(re.sub(r"^#{1,6} ", "", ANCHOR_RE.sub("", l)))
        if not plain:
            continue
        for k in range(j, min(len(items), j + 60)):
            if _norm(items[k]["text"]) == plain:
                j = k + 1
                if not is_head:
                    text = _clean_mark(re.sub(r"\\$", "", ANCHOR_RE.sub("", l)))
                    lines[i] = "#" * items[k]["level"] + " " + text
                    n += 1
                break
    return "\n".join(lines), n


def restore_heading_numbers(md: str, labels):
    """把自动编号补到 pandoc 输出的标题前；按顺序 + 文本匹配，找不到就跳过(不会乱加)。"""
    lines = md.split("\n")
    j, added = 0, 0
    for i, l in enumerate(lines):
        m = re.match(r"^(#{1,6}) (.*)$", l)
        if not m:
            continue
        plain = _norm(IMG_RE.sub("", ANCHOR_RE.sub("", m.group(2))))
        for k in range(j, len(labels)):
            if plain and _norm(labels[k][0]) == plain:
                j = k + 1
                label = labels[k][1]
                if label and not plain.startswith(_norm(label)):
                    sep = "" if re.search(rf"[{CJK}]$", label) else " "
                    lines[i] = f"{m.group(1)} {label}{sep}{m.group(2)}"
                    added += 1
                break
    return "\n".join(lines), added


def fix_empty_table_headers(lines):
    """pandoc 对无表头表格会生成全空表头行；丢掉它，把第一行数据提升为表头(去加粗)。"""
    out, i, fixed = [], 0, 0
    while i < len(lines):
        if (is_table_line(lines[i]) and i + 2 < len(lines)
                and all(c == "" for c in split_row(lines[i]))
                and is_sep_line(lines[i + 1]) and is_table_line(lines[i + 2])):
            cells = [re.sub(r"^\*\*(.*)\*\*$", r"\1", c) for c in split_row(lines[i + 2])]
            out.extend([render_row(cells), lines[i + 1]])
            i += 3
            fixed += 1
            continue
        out.append(lines[i]); i += 1
    return out, fixed


class _TableParser(HTMLParser):
    """解析 pandoc 输出的 HTML 表格：得到 [[(文本, colspan, rowspan)]]。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows, self.cell, self.depth = [], None, 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "table":
            self.depth += 1
        elif tag == "tr":
            self.rows.append([])
        elif tag in ("td", "th") and self.rows is not None:
            self.cell = dict(text="", cs=int(a.get("colspan") or 1), rs=int(a.get("rowspan") or 1))
        elif self.cell is not None:
            if tag == "img":
                self.cell["text"] += f" ![]({a.get('src', '')}) "
            elif tag in ("p", "br", "li", "div"):
                self.cell["text"] += "\n"
            elif tag in ("strong", "b"):
                self.cell["text"] += "**"

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.cell is not None and self.rows:
            parts = [" ".join(x.split()) for x in self.cell["text"].split("\n")]
            text = ""
            for p in parts:
                text = join_text(text, p)
            text = re.sub(r"\*\*\s*\*\*", "", text).strip()
            self.rows[-1].append((text, self.cell["cs"], self.cell["rs"]))
            self.cell = None
        elif self.cell is not None and tag in ("strong", "b"):
            self.cell["text"] += "**"

    def handle_data(self, data):
        if self.cell is not None:
            self.cell["text"] += data


def html_tables_to_md(md: str):
    """pandoc 遇到合并单元格/单元格内多段时输出冗长的 HTML 表格；转成 Markdown 表格(合并格只在首格写内容)。返回 (md, 转换数)。"""
    n = 0

    def repl(m):
        nonlocal n
        block = m.group(0)
        if block.count("<table") > 1:  # 嵌套表格保留 HTML
            return block
        p = _TableParser()
        p.feed(block)
        rows = [r for r in p.rows if r]
        if not rows:
            return block
        grid, pending = [], {}  # pending: (行, 列) 被上方跨行格占用
        for ri, r in enumerate(rows):
            row, ci = [], 0
            for text, cs, rs in r:
                while (ri, ci) in pending:
                    row.append(""); ci += 1
                row.append(text)
                for dc in range(cs):
                    for dr in range(rs):
                        if dr or dc:
                            pending[(ri + dr, ci + dc)] = True
                row.extend([""] * (cs - 1))
                ci += cs
            while (ri, ci) in pending:
                row.append(""); ci += 1
            grid.append(row)
        n += 1
        return "\n\n" + "\n".join(render_table(grid)) + "\n\n"

    return re.sub(r"<table\b.*?</table>", repl, md, flags=re.S), n


def _xml_text(el) -> str:
    return "".join(t.text or "" for t in el.iter() if t.tag in (W + "t", NS_A + "t"))


def _rels(z, part):
    """读取某部件的关系表 {rId: 部件全名}。"""
    d, base = {}, part.rsplit("/", 1)[0]
    rel = f"{base}/_rels/{part.rsplit('/', 1)[1]}.rels"
    if rel not in z.namelist():
        return d
    for r in ET.fromstring(z.read(rel)).iter(NS_REL + "Relationship"):
        t = r.get("Target", "")
        d[r.get("Id")] = t.lstrip("/") if t.startswith("/") else os.path.normpath(f"{base}/{t}").replace("\\", "/")
    return d


CHART_TYPES = {"bar": "柱形/条形图", "bar3D": "柱形/条形图", "line": "折线图", "line3D": "折线图", "pie": "饼图",
               "pie3D": "饼图", "doughnut": "环形图", "area": "面积图", "area3D": "面积图", "scatter": "散点图",
               "radar": "雷达图", "bubble": "气泡图", "ofPie": "复合饼图", "stock": "股价图", "surface": "曲面图"}


def chart_md(z, part) -> str:
    """图表部件 → 题目 + 数据表(取自图表内缓存的数据)。"""
    root = ET.fromstring(z.read(part))
    t = root.find(f"{NS_C}chart/{NS_C}title")
    title = _xml_text(t).strip() if t is not None else ""
    plot = root.find(f".//{NS_C}plotArea")
    kinds, series = [], []

    def pts(el):
        return {} if el is None else {int(p.get("idx")): (p.findtext(NS_C + "v") or "") for p in el.iter(NS_C + "pt")}

    for ch in (list(plot) if plot is not None else []):
        tag = ch.tag.split("}")[1]
        if not tag.endswith("Chart"):
            continue
        kinds.append(CHART_TYPES.get(tag[:-5], tag[:-5]))
        for ser in ch.findall(NS_C + "ser"):
            tx = ser.find(NS_C + "tx")
            name = "".join(v.text or "" for v in tx.iter(NS_C + "v")) or _xml_text(tx) if tx is not None else ""
            cat = ser.find(NS_C + "cat")
            if cat is None:
                cat = ser.find(NS_C + "xVal")
            val = ser.find(NS_C + "val")
            if val is None:
                val = ser.find(NS_C + "yVal")
            series.append((name or f"系列{len(series) + 1}", pts(cat), pts(val)))
    head = f"**图表：{title or '（无标题）'}**（{'、'.join(dict.fromkeys(kinds)) or '图表'}，数据取自图表内嵌数据）"
    if not series:
        return head
    idx = sorted({i for _, c, v in series for i in list(c) + list(v)})
    cats = next((c for _, c, _ in series if c), {})
    grid = [["类别"] + [s[0] for s in series]]
    for i in idx:
        grid.append([cats.get(i, str(i + 1))] + [s[2].get(i, "") for s in series])
    return head + "\n\n" + "\n".join(render_table(grid))


def smartart_md(z, part) -> str:
    root = ET.fromstring(z.read(part))
    items = []
    for pt in root.iter(NS_DGM + "pt"):
        if pt.get("type", "node") in ("node", "asst"):
            t = _xml_text(pt).strip()
            if t:
                items.append(t)
    return "**SmartArt 图示：**\n\n" + "\n".join(f"- {t}" for t in items) if items else "**SmartArt 图示**（无文字）"


def preprocess_docx(src: Path, work: Path):
    """pandoc 读不出图表数据和 SmartArt 文字：先把它们替换成占位符，转换后再填入。
    返回 (送给 pandoc 的 docx 路径, {占位符: markdown}, 图表数, SmartArt 数)。"""
    z = zipfile.ZipFile(src)
    names = z.namelist()
    if not any(n.startswith(("word/charts/", "word/diagrams/")) for n in names):
        return src, {}, 0, 0
    rels = _rels(z, "word/document.xml")
    xml = z.read("word/document.xml").decode("utf8")
    blocks, counts = {}, Counter()

    def handle(m):
        s = m.group(0)
        cm = re.search(r'<c:chart\b[^>]*?r:id="([^"]+)"', s)
        dm = re.search(r'<dgm:relIds\b[^>]*?r:dm="([^"]+)"', s)
        try:
            if cm and rels.get(cm.group(1)) in names:
                md, kind = chart_md(z, rels[cm.group(1)]), "CHART"
            elif dm and rels.get(dm.group(1)) in names:
                md, kind = smartart_md(z, rels[dm.group(1)]), "SMART"
            else:
                return s
        except ET.ParseError:
            return s
        counts[kind] += 1
        key = f"DOC2MD{kind}{len(blocks) + 1}Z"
        blocks[key] = md
        return f"<w:t>{key}</w:t>"

    new = re.sub(r"<mc:AlternateContent>.*?</mc:AlternateContent>",
                 lambda m: handle(m) if ("<c:chart" in m.group(0) or "<dgm:relIds" in m.group(0)) else m.group(0),
                 xml, flags=re.S)
    new = re.sub(r"<w:drawing>.*?</w:drawing>", handle, new, flags=re.S)
    try:
        ET.fromstring(new)
    except ET.ParseError:  # 替换破坏了结构(嵌套绘图等罕见情况)：不替换，图表数据追加到文末
        return src, {"__APPEND__": "\n\n".join(blocks.values())}, counts["CHART"], counts["SMART"]
    out = work / "pre.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            zo.writestr(n, new.encode("utf8") if n == "word/document.xml" else z.read(n))
    return out, blocks, counts["CHART"], counts["SMART"]


def docx_header_footer(z):
    texts = []
    for n in sorted(z.namelist()):
        m = re.match(r"word/(header|footer)\d*\.xml$", n)
        if not m:
            continue
        root = ET.fromstring(z.read(n))
        for p in root.iter(W + "p"):
            t = " ".join("".join(x.text or "" for x in p.iter(W + "t")).split())
            if t and not re.fullmatch(r"[\d\W_]*", t) and (m.group(1), t) not in texts:
                texts.append((m.group(1), t))
    return texts


def docx_comments(z):
    """批注 → [(作者, 内容, 批注所指的原文)]。"""
    if "word/comments.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("word/comments.xml"))
    docxml = z.read("word/document.xml").decode("utf8", "ignore")
    res = []
    for c in root.iter(W + "comment"):
        cid = c.get(W + "id")
        body = " ".join(" ".join("".join(t.text or "" for t in p.iter(W + "t")).split()) for p in c.iter(W + "p")).strip()
        m = re.search(rf'<w:commentRangeStart w:id="{cid}"\s*/>(.*?)<w:commentRangeEnd w:id="{cid}"\s*/>', docxml, re.S)
        anchor = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", m.group(1))) if m else ""
        res.append((c.get(W + "author") or "", body, " ".join(anchor.split())))
    return res


def docx_units_and_risks(docx_path: Path, md: str, out_dir: Path, keep_toc: bool):
    """返回 (逐段原文 units, 报告项)。报告项说明各类特殊内容的处理结果。目录段落(已按设计删除)不计入核对。"""
    z = zipfile.ZipFile(docx_path)
    names = z.namelist()
    docxml = z.read("word/document.xml").decode("utf8", "ignore")
    root = ET.fromstring(docxml)
    toc_ids = set()
    if not keep_toc:
        for s in ET.fromstring(z.read("word/styles.xml")).iter(W + "style"):
            if re.match(r"^toc\s*\d$", _attr(s.find(W + "name"), "val") or "", re.I):
                toc_ids.add(s.get(W + "styleId"))
    units = []
    for i, p in enumerate(root.iter(W + "p")):
        ps = p.find(f"{W}pPr/{W}pStyle")
        if ps is not None and ps.get(W + "val") in toc_ids:
            continue
        t = "".join(x.text or "" for x in p.iter(W + "t"))
        if t.strip():
            units.append((f"第{i + 1}段", t))
    rep = {}
    ins, dels = docxml.count("<w:ins "), docxml.count("<w:del ")
    rep["修订痕迹"] = f"插入 {ins} 处/删除 {dels} 处（已按「接受全部修订」输出）" if ins or dels else "无"
    tb = [p for t in root.iter(W + "txbxContent") for p in t.iter(W + "p")]
    if tb:
        cov, _ = coverage([("文本框", "".join(x.text or "" for x in p.iter(W + "t"))) for p in tb], md)
        rep["文本框"] = f"{docxml.count('<w:txbxContent')} 个（文字{'已转出' if cov >= 0.97 else f'转出 {cov * 100:.0f}%，请核对'}）"
    else:
        rep["文本框"] = "无"
    emb = [n for n in names if n.startswith("word/embeddings/")]
    if emb:
        d = out_dir / "attachments"
        d.mkdir(parents=True, exist_ok=True)
        for n in emb:
            (d / n.rsplit("/", 1)[1]).write_bytes(z.read(n))
        rep["嵌入文件"] = f"{len(emb)} 个（原文件已导出到 attachments/：{', '.join(n.rsplit('/', 1)[1] for n in emb[:5])}）"
    else:
        rep["嵌入文件"] = "无"
    nmath = docxml.count("<m:oMath>") + docxml.count("<m:oMath ")
    rep["公式"] = f"{nmath} 个（已转为 LaTeX）" if nmath else "无"
    return units, rep


def convert_docx(src: Path, out_dir: Path, args):
    pandoc = find_pandoc()
    if not pandoc:
        sys.exit("未找到 pandoc，无法转换 Word 文档。请安装 pandoc（https://pandoc.org/installing.html），"
                 "或执行 python -m pip install pypandoc_binary；也可用环境变量 DOC2MD_PANDOC 指定路径。")
    work = Path(tempfile.mkdtemp(prefix="doc2md_"))
    report = {}
    feed, blocks, nchart, nsmart = preprocess_docx(src, work)
    md_file = work / "raw.md"
    subprocess.run([pandoc, str(feed), "-t", "gfm", "--wrap=none", "--track-changes=accept",
                    "--extract-media=media", "-o", str(md_file)], cwd=str(work), check=True)
    md = md_file.read_text(encoding="utf-8")

    for k, v in blocks.items():
        md = md.replace(k, "\n\n" + v + "\n\n") if k != "__APPEND__" else md + "\n\n" + v + "\n"
    report["图表(数据已转为表格)/SmartArt(文字已转出)"] = f"{nchart}/{nsmart}"

    items = docx_heading_labels(src)
    md, ntoc = strip_docx_toc(md, args.keep_toc)
    report["移除目录行"] = ntoc
    md, promoted = promote_docx_headings(md, [i for i in items])
    report["标题(Word标题样式/按自定义样式、目录书签或加粗放大字号补出)"] = f"{sum(i['builtin'] for i in items)}/{promoted}"
    md = ANCHOR_RE.sub("", md)
    md, added = restore_heading_numbers(md, [(i["text"], i["label"]) for i in items])
    report["补回的标题自动编号"] = added
    md = md.replace("\\\n", "\n")  # pandoc 的硬换行 → 普通换行
    md, nhtml = html_tables_to_md(md)
    report["复杂表格转为 Markdown 表格"] = nhtml
    lines, demoted = fix_headings(md.split("\n"))
    report["降级的伪标题(整句被设成标题样式)"] = demoted
    lines = [re.sub(r"</?u>", "", l) for l in lines]
    lines, fixed = fix_empty_table_headers(lines)
    report["修复的空表头"] = fixed
    lines, ndec = decorated_headings(lines)
    report["装饰性章标题(单行表格排版)还原"] = ndec
    md = "\n".join(lines)

    z = zipfile.ZipFile(src)
    hf = docx_header_footer(z)
    if hf:
        md = "<!-- " + "；".join(f"{'页眉' if k == 'header' else '页脚'}：{t}" for k, t in hf) + " -->\n\n" + md
    report["页眉页脚"] = f"{len(hf)} 条（已写入 md 开头的注释）" if hf else "无"
    cmts = docx_comments(z)
    if cmts:
        md += "\n\n---\n\n**批注**（共 {} 条）\n\n".format(len(cmts)) + "\n".join(
            f"{i}. {('[' + a + '] ') if a else ''}{b}" + (f" ——批注位置：「{s[:60]}」" if s else "")
            for i, (a, b, s) in enumerate(cmts, 1)) + "\n"
    report["批注"] = f"{len(cmts)} 条（已附在文末）" if cmts else "无"

    md, nimg = renumber_images(md, lambda s: (work / s), out_dir, work)
    md = finalize(md)

    units, risks = docx_units_and_risks(src, md, out_dir, args.keep_toc)
    cov, bad = coverage(units, md, min_ratio=0.3)
    report["文字覆盖率"] = f"{cov * 100:.2f}%"
    report["覆盖率偏低的段落"] = [f"{l} 缺{r * 100:.0f}% 例:「{s}」" for l, r, s in bad[:10]] or "无"
    report.update(risks)
    report["题注提到却未输出的表格"] = check_captions(md)[:10] or "无"
    shutil.rmtree(work, ignore_errors=True)
    return md, nimg, report


def doc_to_docx(src: Path, work: Path) -> Path:
    """旧版 .doc 用 LibreOffice 转成 .docx。"""
    if not run_soffice(["--convert-to", "docx", "--outdir", str(work), str(src)], work):
        sys.exit("旧版 .doc 需要 LibreOffice 才能转换，但未找到。请安装 LibreOffice（https://www.libreoffice.org/），"
                 "或用环境变量 DOC2MD_SOFFICE 指定 soffice 路径；也可以在 Word 中另存为 .docx 后再处理。")
    out = work / (src.stem + ".docx")
    if not out.exists():
        sys.exit("LibreOffice 未能把 .doc 转成 .docx（文件可能损坏或加密），请在 Word 中另存为 .docx 后再处理。")
    return out


# ───────────────────────── 入口 ─────────────────────────
def estimate_tokens(text: str) -> int:
    cjk = len(re.findall(rf"[{CJK}]", text))
    return int(cjk * 1.0 + (len(text) - cjk) / 3.5)


def main():
    ap = argparse.ArgumentParser(description="PDF/DOCX/DOC → Markdown + 顺序编号图片")
    ap.add_argument("input", nargs="?")
    ap.add_argument("-o", "--out", default=None, help="输出根目录，默认为输入文件同目录下的 _md 文件夹")
    ap.add_argument("--keep-toc", action="store_true", help="保留目录(默认移除)")
    ap.add_argument("--keep-header-footer", action="store_true", help="保留PDF页眉页脚")
    ap.add_argument("--dpi", type=int, default=150, help="PDF 图片与整页图的渲染分辨率")
    ap.add_argument("--password", default=None, help="加密 PDF 的打开密码")
    ap.add_argument("--check", action="store_true", help="只检查运行环境，不转换")
    args = ap.parse_args()

    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    if args.check:
        sys.exit(0 if check_env() else 1)
    if not args.input:
        ap.error("请提供输入文件路径（或使用 --check 检查环境）")
    src = Path(args.input).resolve()
    if not src.is_file():
        sys.exit(f"找不到文件：{src}")
    root = Path(args.out) if args.out else src.parent / "_md"
    out_dir = root / src.stem
    # 重跑时只清理本工具生成的产物，不整目录删除
    for sub in ("images", "pages", "attachments"):
        if (out_dir / sub).exists():
            shutil.rmtree(out_dir / sub)
    (out_dir / f"{src.stem}.md").unlink(missing_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    ext = src.suffix.lower()
    if ext == ".pdf":
        md, nimg, report = convert_pdf(src, out_dir, args)
    elif ext == ".docx":
        md, nimg, report = convert_docx(src, out_dir, args)
    elif ext == ".doc":
        tmp = Path(tempfile.mkdtemp(prefix="doc2md_doc_"))
        try:
            md, nimg, report = convert_docx(doc_to_docx(src, tmp), out_dir, args)
            report = {"来源": "旧版 .doc，已用 LibreOffice 转为 .docx 后处理", **report}
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    else:
        sys.exit(f"不支持的格式：{ext}（支持 .pdf / .docx / .doc）")

    md_path = out_dir / f"{src.stem}.md"
    md_path.write_text(md, encoding="utf-8", newline="\n")

    print(f"输出：{md_path}")
    print(f"图片：{nimg} 张 → {out_dir / 'images'}")
    if (out_dir / "pages").exists():
        print(f"整页图：{len(list((out_dir / 'pages').iterdir()))} 张 → {out_dir / 'pages'}")
    print(f"字符数：{len(md)}  预估 token：约 {estimate_tokens(md)}")
    print("── 转换报告 ──")
    for k, v in report.items():
        print(f"  {k}：{v}")


if __name__ == "__main__":
    main()
