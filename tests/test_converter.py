# -*- coding: utf-8 -*-
"""测试:生成中文测试 PDF 并转换,验证段落重建效果"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import fitz
from docx import Document

from app.converter import convert_pdf_to_docx, rebuild_paragraphs, _extract_lines

TEST_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_PDF = os.path.join(TEST_DIR, "sample.pdf")
SAMPLE_DOCX = os.path.join(TEST_DIR, "sample.docx")


def make_sample_pdf(path: str):
    """手工生成一个含标题、多段正文、首行缩进、居中文本的 PDF"""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4

    # 标题(居中、大号、加粗)
    page.insert_text(
        (297 - 80, 80), "老人健康饮食指南", fontsize=24, fontname="china-s",
        color=(0, 0, 0), render_mode=0,
    )

    # 副标题
    page.insert_text((297 - 60, 120), "—— 给家人的一份贴心建议", fontsize=14, fontname="china-s")

    # 正文段落 1(首行缩进模拟)
    body1 = "随着年龄增长,身体机能逐渐下降,合理的饮食搭配显得尤为重要。老年人应当注重蛋白质的摄入,同时控制油脂和盐分的用量,保持营养均衡。"
    y = 170
    for i in range(0, len(body1), 24):
        seg = body1[i:i + 24]
        page.insert_text((60 + (12 if i == 0 else 0), y), seg, fontsize=12, fontname="china-s")
        y += 18

    # 正文段落 2(正常左对齐,无缩进)
    body2 = "每天坚持适量运动,如散步、打太极等,有助于增强体质。多喝水,多吃新鲜蔬菜水果,少吃辛辣刺激的食物,保证充足的睡眠。"
    y += 10
    for i in range(0, len(body2), 24):
        seg = body2[i:i + 24]
        page.insert_text((60, y), seg, fontsize=12, fontname="china-s")
        y += 18

    # 居中文本(模拟引用/说明)
    page.insert_text((297 - 90, y + 30), "—— 以上建议仅供参考 ——", fontsize=12, fontname="china-s")

    doc.save(path)
    doc.close()
    print(f"[OK] 测试 PDF 已生成: {path}")


def verify_docx(path: str):
    """检查 docx 段落结构:验证是否重建为逻辑段落"""
    doc = Document(path)
    paras = [p for p in doc.paragraphs if p.text.strip()]
    print(f"\n[INFO] docx 共 {len(paras)} 个非空段落:")
    for i, p in enumerate(paras):
        style = p.style.name if p.style else "?"
        align = p.alignment
        indent = p.paragraph_format.first_line_indent
        print(f"  段{i}: style={style} align={align} indent={indent} text={p.text[:30]}...")
    return paras


def main():
    make_sample_pdf(SAMPLE_PDF)

    # 先看提取的原始行数
    pdf = fitz.open(SAMPLE_PDF)
    lines = _extract_lines(pdf[0])
    pdf.close()
    print(f"[INFO] 提取原始行数: {len(lines)}")
    for l in lines:
        print(f"  行: size={l.size:.0f} bbox={[round(v) for v in l.bbox]} align={l.align} text={l.text[:25]}")

    # 段落重建
    paras = rebuild_paragraphs(lines, 595, 842)
    print(f"\n[INFO] 重建段落数: {len(paras)} (期望约 4 段)")
    for i, p in enumerate(paras):
        print(f"  段{i}: style={p.style} align={p.align} indent_first={p.indent_first:.0f} lines={len(p.lines)} text={p.text[:30]}")

    # 完整转换
    result = convert_pdf_to_docx(SAMPLE_PDF, SAMPLE_DOCX, font_size_label="三号")
    print(f"\n[INFO] 转换结果: {result}")
    verify_docx(SAMPLE_DOCX)

    # 基本断言
    assert len(paras) >= 3, f"段落重建异常: 只有 {len(paras)} 段"
    assert any(p.style == "title" for p in paras), "未识别标题"
    print("\n✅ 测试通过!")


if __name__ == "__main__":
    main()
