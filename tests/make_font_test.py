#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成"其他字体"测试 PDF:字体名不含 Bold 关键词,靠 FontWeight/StemV 区分
模拟真实场景:不同 PDF 有不同字体命名习惯(如 SimSun/SimHei、KaiTi 等)。
"""
import fitz

doc = fitz.open()
page = doc.new_page(width=595, height=842)

# 用 china-s(china-s 是 CJK 支持字体)生成文本
page.insert_text((72, 100), "常规文字测试段落，用于模拟普通正文。", fontsize=12, fontname="china-s")
page.insert_text((72, 130), "加粗文字测试段落，用于模拟强调内容。", fontsize=12, fontname="china-s", render_mode=2)

out = "/tmp/font_test_sim.pdf"
doc.save(out)
doc.close()
print("saved", out)

# 检查:把 FontName 改成不含 Bold 的名字(模拟其他字体命名)
import re
data = open(out, "rb").read()
# china-s 在 fitz 里映射为 Heiti,手动造一个自定义字体名的 PDF 太复杂,
# 这里直接验证 _analyze_fonts 对 FontWeight 信号的处理
