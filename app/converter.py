# -*- coding: utf-8 -*-
"""
PDF 转 Word 核心转换模块
-------------------------
技术要点(解决"改字号后异常换行/缩进"问题):
1. 段落重建:PDF 中同一逻辑段落的分散行会被合并为真正的 Word 段落,
   而不是每行一个孤立段落,避免改字号后 Word 重新排版时乱换行。
2. 格式统一:所有正文统一使用 Word 的"正文"样式,字号由样式控制,
   不保留 PDF 的逐行硬编码格式(行距/缩进),改字号时全文自动重排。
3. 合理缩进:首行缩进用段落属性实现(而不是手动空格),改字号不会错位。

依赖:PyMuPDF(fitz)、python-docx
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

@dataclass
class TextLine:
    """PDF 中提取的一行文本(可能只是段落的一部分)"""
    text: str
    bbox: Tuple[float, float, float, float]  # x0, y0, x1, y1
    size: float          # 字号(pt)
    font: str
    bold: bool
    align: Optional[str]  # 'left' | 'center' | 'right' | 'justify' | None
    block_id: int = 0     # 所属 PDF 文本块 ID(辅助段落判断)


@dataclass
class Paragraph:
    """重建后的逻辑段落"""
    lines: List[TextLine] = field(default_factory=list)
    style: str = "body"      # 'title' | 'heading' | 'body' | 'caption'
    align: str = "left"
    indent_first: float = 0.0   # 首行缩进(pt)
    indent_left: float = 0.0    # 左缩进(pt)

    @property
    def text(self) -> str:
        # 合并行内文本:英文单词间需要空格,中文直接拼接
        parts = []
        for line in self.lines:
            parts.append(line.text)
        return "".join(parts) if _is_cjk_text("".join(parts)) else " ".join(p.strip() for p in parts)

    @property
    def size(self) -> float:
        if not self.lines:
            return 12.0
        return max(l.size for l in self.lines)


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]")


def _is_cjk_text(text: str) -> bool:
    """判断文本是否以中文为主"""
    if not text:
        return False
    cjk = len(_CJK_RE.findall(text))
    return cjk / max(len(text), 1) > 0.3


def _clean_text(s: str) -> str:
    """清理文本:去空白、修正常见 OCR/排版残留、剔除非法 XML 控制字符"""
    # 剔除 XML 1.0 不允许的控制字符(NULL、0x01-0x08、0x0B、0x0C、0x0E-0x1F 等)
    s = re.sub(
        r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]",
        "",
        s,
    )
    s = s.replace("\u3000", " ").replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = s.strip()
    return s


def _is_blank_line(s: str) -> bool:
    return not s.strip() or len(s.strip()) <= 1


# ---------------------------------------------------------------------------
# 段落重建
# ---------------------------------------------------------------------------

def _extract_lines(page: fitz.Page) -> List[TextLine]:
    """从页面提取文本行(使用 dict 模式获得字体/字号/位置信息)"""
    lines: List[TextLine] = []
    blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)["blocks"]
    for block in blocks:
        if block.get("type") != 0:  # 只处理文本块,跳过图片
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            text = _clean_text("".join(s.get("text", "") for s in spans))
            if not text:
                continue
            bbox = line.get("bbox", (0, 0, 0, 0))
            sizes = [s.get("size", 12.0) for s in spans]
            fonts = [s.get("font", "") for s in spans]
            flags = [s.get("flags", 0) for s in spans]
            # flags bit 4 (16) = bold
            bold = any(f & 16 for f in flags)
            lines.append(TextLine(
                text=text,
                bbox=bbox,
                size=max(sizes),
                font=max(fonts, key=fonts.count),
                bold=bold,
                align=_detect_align(line, page.rect.width),
                block_id=block.get("number", 0),
            ))
    return lines


def _detect_align(line, page_width: float) -> str:
    x0, _, x1, _ = line.get("bbox", (0, 0, 0, 0))
    line_width = x1 - x0
    left_margin = x0
    right_margin = page_width - x1
    if left_margin < 5 and right_margin < 5:
        return "justify" if line_width > page_width * 0.8 else "left"
    if left_margin > right_margin * 2.5 and left_margin > 20:
        return "right"
    if right_margin > left_margin * 2.5 and right_margin > 20:
        return "left"  # 有左缩进,不算居中
    if left_margin > 10 and right_margin > 10:
        return "center"
    return "left"


def _infer_style(line: TextLine, is_first: bool, para_lines: List[TextLine]) -> str:
    """推断段落样式:标题 / 小标题 / 正文(基于字号+加粗+对齐,不强制 bold 标志)"""
    if line.size >= 20:
        return "title"
    if line.size >= 14:
        # 大字号通常是小标题;居中且短文本也可能是标题
        if line.bold or line.align == "center" or len(line.text) < 30:
            return "heading"
    return "body"


def _same_paragraph(prev: TextLine, cur: TextLine, para_gap: float, line_gap: float) -> bool:
    """判断两行是否属于同一逻辑段落(核心算法)
    :param para_gap: 段落间距阈值(动态计算,pt)
    :param line_gap: 行内间距中位数(pt)
    """
    # 垂直距离:当前行顶部 - 上一行底部
    gap = cur.bbox[1] - prev.bbox[3]

    # 若行距大于段落阈值 → 新段落
    if gap > para_gap:
        return False

    # 若当前行有显著左缩进(>12pt),通常是新段落(如正文首行缩进)
    if cur.bbox[0] - prev.bbox[0] > 12:
        return False

    # 不同文本块且行距明显大于典型行距 → 新段
    if prev.block_id != cur.block_id and gap > line_gap:
        return False

    return True


def _estimate_para_gap(lines: List[TextLine]) -> Tuple[float, float]:
    """动态估算间距阈值
    :return: (段落阈值, 行内间距基准)
    行内间距取低分位值(20%),段落阈值 = 行内间距 * 2.0(至少 10pt),
    避免被标题/图片间的大间距拉高阈值。
    """
    gaps = []
    for i in range(1, len(lines)):
        g = lines[i].bbox[1] - lines[i - 1].bbox[3]
        if g > 0:
            gaps.append(g)
    if not gaps:
        return 20.0, 4.0
    gaps.sort()
    # 20% 分位作为行内典型间距
    idx = min(len(gaps) - 1, max(0, int(len(gaps) * 0.2)))
    line_gap = gaps[idx]
    return max(line_gap * 2.0, 10.0), line_gap


def _align_from_lines(para: Paragraph, page_width: float) -> str:
    """根据行位置综合判断对齐方式"""
    if not para.lines:
        return "left"
    aligns = [l.align for l in para.lines]
    from collections import Counter
    return Counter(aligns).most_common(1)[0][0] or "left"


def rebuild_paragraphs(lines: List[TextLine], page_width: float, page_height: float) -> List[Paragraph]:
    """把提取的行重建为逻辑段落"""
    paragraphs: List[Paragraph] = []
    current: Optional[Paragraph] = None
    para_gap, line_gap = _estimate_para_gap(lines)

    for line in lines:
        if _is_blank_line(line.text):
            if current is not None:
                paragraphs.append(current)
                current = None
            continue

        if current is None:
            current = Paragraph(lines=[line])
            current.align = line.align
            continue

        if _same_paragraph(current.lines[-1], line, para_gap, line_gap):
            current.lines.append(line)
        else:
            paragraphs.append(current)
            current = Paragraph(lines=[line])
            current.align = line.align

    if current is not None:
        paragraphs.append(current)

    # 后处理:样式推断 + 对齐修正
    for para in paragraphs:
        para.align = _align_from_lines(para, page_width)
        para.style = _infer_style(para.lines[0], True, para.lines)
        # 首行缩进检测:第二行比第一行靠左,且第一行有缩进 → 段落缩进
        if len(para.lines) > 1:
            first_x0 = para.lines[0].bbox[0]
            rest_x0 = min(l.bbox[0] for l in para.lines[1:])
            indent = first_x0 - rest_x0
            if indent > 8:
                para.indent_first = indent
            para.indent_left = rest_x0

    return paragraphs


# ---------------------------------------------------------------------------
# DOCX 生成
# ---------------------------------------------------------------------------

FONT_SIZE_MAP = {
    "小四": 12,
    "四号": 14,
    "小三": 15,
    "三号": 16,
    "小二": 18,
    "二号": 22,
}


def _setup_style(doc: Document, font_name: str, font_size_pt: int):
    """设置文档默认样式,统一控制字体字号"""
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(font_size_pt)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_after = Pt(6)
    pf.space_before = Pt(0)


def _set_run_font(run, font_name: str, size_pt: float, bold: bool = False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def add_paragraph(doc: Document, para: Paragraph, font_name: str, base_size: float):
    """把重建段落写入 docx"""
    p = doc.add_paragraph()
    pf = p.paragraph_format

    # 对齐
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    pf.alignment = align_map.get(para.align, WD_ALIGN_PARAGRAPH.LEFT)

    # 缩进用段落属性(不是空格)→ 改字号不错位
    if para.indent_first > 0:
        pf.first_line_indent = Pt(para.indent_first)
    if para.indent_left > 0:
        pf.left_indent = Pt(para.indent_left)

    # 样式相关字号
    if para.style == "title":
        size = max(base_size * 1.8, 26)
        bold = True
    elif para.style == "heading":
        size = max(base_size * 1.4, 18)
        bold = True
    else:
        size = base_size
        bold = False

    text = para.text
    run = p.add_run(text)
    _set_run_font(run, font_name, size, bold)

    return p


def convert_pdf_to_docx(
    pdf_path: str,
    docx_path: str,
    font_size_label: str = "三号",
    font_name: str = "微软雅黑",
    progress_cb=None,
) -> dict:
    """
    转换主入口
    :param pdf_path: 输入 PDF
    :param docx_path: 输出 DOCX
    :param font_size_label: 字号('小四'/'四号'/'小三'/'三号'/'小二'/'二号')
    :param font_name: 中文字体名
    :param progress_cb: 进度回调(page_index, total_pages)
    """
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")

    base_size = FONT_SIZE_MAP.get(font_size_label, 16)
    doc = Document()
    _setup_style(doc, font_name, base_size)

    pdf = fitz.open(pdf_path)
    total = pdf.page_count
    para_count = 0

    try:
        for i, page in enumerate(pdf):
            if progress_cb:
                progress_cb(i + 1, total)
            lines = _extract_lines(page)
            paragraphs = rebuild_paragraphs(lines, page.rect.width, page.rect.height)
            for para in paragraphs:
                add_paragraph(doc, para, font_name, base_size)
                para_count += 1
    finally:
        pdf.close()

    doc.save(docx_path)

    return {
        "pages": total,
        "paragraphs": para_count,
        "output": docx_path,
        "font": font_name,
        "font_size": font_size_label,
    }
