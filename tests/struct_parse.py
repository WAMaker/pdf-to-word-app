#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 StructElem 的 ActualText 提取段落(只遍历两层:Document→P→Span)"""
import re
from typing import List, Optional

import fitz


def _decode_actualtext(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("<FEFF") and raw.endswith(">"):
        try:
            return bytes.fromhex(raw[5:-1]).decode("utf-16-be", errors="replace")
        except Exception:
            return ""
    if raw.startswith("<") and raw.endswith(">"):
        try:
            hexs = raw[1:-1]
            if len(hexs) % 4 == 0:
                return bytes.fromhex(hexs).decode("utf-16-be", errors="replace")
            return bytes.fromhex(hexs).decode("latin1", errors="replace")
        except Exception:
            return ""
    # 字面量:过滤孤立的括号残留(Word导出PDF时空格的错误表示)
    t = raw
    if t in ("(", ")", "（", "）", "( ", " )"):
        return " "
    return t


def parse_struct_paragraphs(pdf: fitz.Document) -> List[dict]:
    # StructTreeRoot
    root = None
    for xref in range(1, pdf.xref_length()):
        obj = pdf.xref_object(xref, compressed=False)
        if "/Type /Catalog" in obj and "/StructTreeRoot" in obj:
            root = int(re.search(r"/StructTreeRoot\s+(\d+)\s+0\s+R", obj).group(1))
            break
    if not root:
        return []
    root_obj = pdf.xref_object(root, compressed=False)
    m = re.search(r"/K\s*\[\s*(\d+)\s+0\s+R", root_obj)
    if not m:
        return []
    doc_elem = int(m.group(1))

    doc_obj = pdf.xref_object(doc_elem, compressed=False)
    kid_xrefs = [int(x) for x in re.findall(r"(\d+)\s+0\s+R", doc_obj)]

    paras: List[dict] = []
    for k in kid_xrefs:
        if k == doc_elem:
            continue
        try:
            obj = pdf.xref_object(k, compressed=False)
        except Exception:
            continue
        if "/StructElem" not in obj:
            continue
        sm = re.search(r"/S\s+/(\w+)", obj)
        stype = sm.group(1) if sm else "?"
        pgm = re.search(r"/Pg\s+(\d+)\s+0\s+R", obj)
        page_xref = int(pgm.group(1)) if pgm else None

        para = {"type": stype, "text": "", "page_xref": page_xref}
        # P 的 K 子节点 = Span 元素(带 ActualText)
        span_xrefs = [int(x) for x in re.findall(r"(\d+)\s+0\s+R", obj)]
        for sx in span_xrefs:
            if sx == k:
                continue
            try:
                sobj = pdf.xref_object(sx, compressed=False)
            except Exception:
                continue
            if "/StructElem" not in sobj:
                continue
            ssm = re.search(r"/S\s+/(\w+)", sobj)
            stype2 = ssm.group(1) if ssm else "?"
            atext_m = re.search(r"/ActualText\s+([^\s/]+)", sobj)
            if stype2 == "Span" and atext_m:
                para["text"] += _decode_actualtext(atext_m.group(1))
            # 处理嵌套(如 Span 内还有 Span/LI)
            if stype2 in ("Span", "L", "LI"):
                sub_xrefs = [int(x) for x in re.findall(r"(\d+)\s+0\s+R", sobj)]
                for ssx in sub_xrefs:
                    if ssx == sx:
                        continue
                    try:
                        ssobj = pdf.xref_object(ssx, compressed=False)
                    except Exception:
                        continue
                    if "/StructElem" not in ssobj:
                        continue
                    sat_m = re.search(r"/ActualText\s+([^\s/]+)", ssobj)
                    if sat_m:
                        para["text"] += _decode_actualtext(sat_m.group(1))
        if para["text"].strip():
            paras.append(para)
    return paras


if __name__ == "__main__":
    import sys
    pdf = fitz.open(sys.argv[1])
    paras = parse_struct_paragraphs(pdf)
    print(f"段落数: {len(paras)}")
    for p in paras[:25]:
        pno = "?"
        if p["page_xref"]:
            for pi in range(pdf.page_count):
                if pdf[pi].xref == p["page_xref"]:
                    pno = pi + 1
                    break
        print(f"[{p['type']}] p{pno}: {p['text'][:65]}")
    pdf.close()
