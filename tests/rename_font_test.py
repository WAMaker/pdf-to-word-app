#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""构造'其他字体'测试:直接修改 PDF 对象,把字体名改为不含 Bold 的自定义名,
验证 _analyze_fonts 的多信号链(FontWeight 优先)不受字体名影响。
"""
import fitz
import os
import re
import shutil

# 1. 用现有经络 PDF 复制一份
src = "samples/抱朴-經絡輔導課20.pdf"
dst = "/tmp/font_renamed.pdf"
shutil.copy(src, dst)

# 2. 修改 PDF 中字体名:去掉 Bold 后缀(模拟命名不含 Bold 的字体)
pdf = fitz.open(dst)
# 找到所有 FontDescriptor 里的字体名并重写
renamed = {}
for xref in range(1, pdf.xref_length()):
    try:
        obj = pdf.xref_object(xref, compressed=False)
    except Exception:
        continue
    if "MicrosoftYaHei-Bold" in obj:
        new_obj = obj.replace("MicrosoftYaHei-Bold", "MyCustomFont-Bd")
        pdf.update_object(xref, new_obj)
        renamed[xref] = "MyCustomFont-Bd"
    elif "MicrosoftYaHei" in obj and "/BaseFont" in obj:
        # 只改 BaseFont/字体资源里的常规名,保留 FontDescriptor 的 FontName
        pass
print("改写的粗体字体对象:", list(renamed.keys()))
pdf.save("/tmp/font_renamed2.pdf", incremental=False, encryption=fitz.PDF_ENCRYPT_KEEP)
pdf.close()

# 3. 验证 _analyze_fonts 能识别(通过 FontWeight=700 而非名字)
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.converter import _analyze_fonts
pdf2 = fitz.open("/tmp/font_renamed2.pdf")
fm = _analyze_fonts(pdf2)
print("重命名后 font_map:", fm)
pdf2.close()
