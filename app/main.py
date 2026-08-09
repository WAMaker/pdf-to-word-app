# -*- coding: utf-8 -*-
"""
极简 PDF 转 Word 界面(老人友好设计)
-------------------------------------
- 全流程:选文件 → 选字号 → 转换 → 预览效果
- 大按钮、大字,避免复杂选项
- 支持预览:转换后可切换字号,实时预览调整后的效果(图片保留)
"""
from __future__ import annotations

import html
import os
import sys
import tempfile
import threading

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from app.converter import convert_pdf_to_docx

APP_TITLE = "PDF 转 Word"
FONT_SIZES = ["小四", "四号", "小三", "三号", "小二", "二号"]
FONT_SIZE_DESC = {
    "小四": "12 号 · 最小",
    "四号": "14 号",
    "小三": "15 号",
    "三号": "16 号 · 常用",
    "小二": "18 号",
    "二号": "22 号 · 最大",
}
# 字体选项(Windows 内置中文字体)
FONTS = ["微软雅黑", "宋体", "黑体", "楷体", "仿宋"]
FONT_DESC = {
    "微软雅黑": "清晰现代 · 推荐",
    "宋体": "正式文档风格",
    "黑体": "粗壮醒目",
    "楷体": "柔和手写感",
    "仿宋": "公文风格",
}
OUT_DIR_NAME = "PDF转Word结果"

# 预览样式对应关系
_STYLE_TAG = {
    "title": "h1",
    "heading": "h3",
    "body": "p",
}


class ConvertThread(QThread):
    """后台转换线程,避免界面卡死"""
    progress = Signal(int, int)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, pdf_path, docx_path, font_size, font_name, image_dir):
        super().__init__()
        self.pdf_path = pdf_path
        self.docx_path = docx_path
        self.font_size = font_size
        self.font_name = font_name
        self.image_dir = image_dir

    def run(self):
        try:
            result = convert_pdf_to_docx(
                self.pdf_path,
                self.docx_path,
                font_size_label=self.font_size,
                font_name=self.font_name,
                progress_cb=lambda cur, total: self.progress.emit(cur, total),
                image_dir=self.image_dir,
            )
            self.finished_ok.emit(result)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(860, 760)
        self.selected_pdf: str | None = None
        self.thread: ConvertThread | None = None
        self.last_result: dict | None = None
        self.preview_temp_dir = tempfile.mkdtemp(prefix="pdf2word_")

        self._build_ui()
        self._apply_fonts()

    # ------------------------------------------------------------------
    # UI 构建
    # ------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(30, 24, 30, 24)
        root.setSpacing(14)

        # 标题
        title = QLabel("PDF 转 Word")
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("title")
        root.addWidget(title)

        # 第一步:选择文件
        root.addWidget(self._section_label("① 选择 PDF 文件"))
        self.file_btn = QPushButton("📄  点击选择 PDF 文件")
        self.file_btn.setObjectName("fileBtn")
        self.file_btn.setMinimumHeight(56)
        self.file_btn.clicked.connect(self.choose_file)
        root.addWidget(self.file_btn)

        self.file_label = QLabel("尚未选择文件")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setWordWrap(True)
        self.file_label.setObjectName("fileLabel")
        root.addWidget(self.file_label)

        # 第二步:选择字号(切换即刷新预览)
        root.addWidget(self._section_label("② 选择文字大小"))
        size_row = QHBoxLayout()
        size_row.setSpacing(10)
        self.size_buttons: dict[str, QPushButton] = {}
        self.selected_size = "三号"
        for label in FONT_SIZES:
            btn = QPushButton(label)
            btn.setObjectName("sizeBtn")
            btn.setMinimumHeight(52)
            btn.setToolTip(FONT_SIZE_DESC[label])
            btn.clicked.connect(lambda _=False, s=label: self.select_size(s))
            size_row.addWidget(btn)
            self.size_buttons[label] = btn
        root.addLayout(size_row)

        self.size_hint = QLabel(f"当前选择:{FONT_SIZE_DESC[self.selected_size]}")
        self.size_hint.setAlignment(Qt.AlignCenter)
        self.size_hint.setObjectName("hint")
        root.addWidget(self.size_hint)

        # 第三步:选择字体(切换即刷新预览)
        root.addWidget(self._section_label("③ 选择字体"))
        font_row = QHBoxLayout()
        font_row.setSpacing(10)
        self.font_buttons: dict[str, QPushButton] = {}
        self.selected_font = "微软雅黑"
        for name in FONTS:
            btn = QPushButton(name)
            btn.setObjectName("fontBtn")
            btn.setMinimumHeight(48)
            btn.setToolTip(FONT_DESC[name])
            btn.clicked.connect(lambda _=False, n=name: self.select_font(n))
            font_row.addWidget(btn)
            self.font_buttons[name] = btn
        root.addLayout(font_row)

        self.font_hint = QLabel(f"当前字体:{FONT_DESC[self.selected_font]}")
        self.font_hint.setAlignment(Qt.AlignCenter)
        self.font_hint.setObjectName("hint")
        root.addWidget(self.font_hint)

        # 第四步:转换
        root.addWidget(self._section_label("④ 开始转换"))
        self.convert_btn = QPushButton("🚀  开始转换")
        self.convert_btn.setObjectName("convertBtn")
        self.convert_btn.setMinimumHeight(64)
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.start_convert)
        root.addWidget(self.convert_btn)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setMinimumHeight(24)
        root.addWidget(self.progress)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setObjectName("status")
        root.addWidget(self.status_label)

        # 预览区
        self.preview_title = QLabel("效果预览(切换字号/字体即时更新)")
        self.preview_title.setAlignment(Qt.AlignCenter)
        self.preview_title.setObjectName("previewTitle")
        self.preview_title.setVisible(False)
        root.addWidget(self.preview_title)

        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_scroll.setVisible(False)
        self.preview_scroll.setObjectName("previewScroll")
        self.preview_browser = QTextBrowser()
        self.preview_browser.setOpenExternalLinks(False)
        self.preview_browser.setObjectName("previewBrowser")
        self.preview_scroll.setWidget(self.preview_browser)
        self.preview_scroll.setMinimumHeight(240)
        root.addWidget(self.preview_scroll, stretch=1)

    def _section_label(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("section")
        return lbl

    def _apply_fonts(self):
        """全局放大字体,适配老人阅读"""
        font = QFont()
        font.setPointSize(14)
        QApplication.instance().setFont(font)
        self.setStyleSheet("""
            QLabel#title {
                font-size: 32px; font-weight: bold;
                color: #2c3e50; padding: 4px;
            }
            QLabel#section {
                font-size: 18px; font-weight: bold; color: #7f8c8d;
                padding-top: 6px;
            }
            QLabel#fileLabel { font-size: 15px; color: #34495e; }
            QLabel#hint { font-size: 15px; color: #16a085; }
            QLabel#status { font-size: 16px; color: #2c3e50; }
            QLabel#previewTitle {
                font-size: 17px; font-weight: bold; color: #16a085;
                padding-top: 8px;
            }
            QPushButton#fileBtn {
                font-size: 20px; background: #ecf0f1; color: #2c3e50;
                border-radius: 12px; border: 2px dashed #bdc3c7;
            }
            QPushButton#fileBtn:hover { background: #e8f8f5; border-color: #16a085; }
            QPushButton#sizeBtn {
                font-size: 18px; background: #ecf0f1; color: #2c3e50;
                border-radius: 10px; border: 2px solid #bdc3c7;
            }
            QPushButton#sizeBtn:hover { background: #e8f8f5; }
            QPushButton#sizeBtn[selected="true"],
            QPushButton#fontBtn[selected="true"] {
                background: #16a085; color: white; border-color: #16a085;
                font-weight: bold;
            }
            QPushButton#fontBtn {
                font-size: 17px; background: #ecf0f1; color: #2c3e50;
                border-radius: 10px; border: 2px solid #bdc3c7;
            }
            QPushButton#fontBtn:hover { background: #e8f8f5; }
            QPushButton#convertBtn {
                font-size: 22px; font-weight: bold;
                background: #16a085; color: white; border-radius: 14px;
                border: none;
            }
            QPushButton#convertBtn:hover { background: #1abc9c; }
            QPushButton#convertBtn:disabled { background: #bdc3c7; }
            QProgressBar {
                font-size: 14px; border-radius: 10px; text-align: center;
                background: #ecf0f1;
            }
            QProgressBar::chunk { background: #16a085; border-radius: 10px; }
            QScrollArea#previewScroll {
                border: 2px solid #bdc3c7; border-radius: 10px;
                background: white;
            }
            QTextBrowser#previewBrowser {
                border: none; background: white;
                font-size: 15px;
            }
        """)
        self._refresh_size_buttons()
        self._refresh_font_buttons()

    # ------------------------------------------------------------------
    # 交互逻辑
    # ------------------------------------------------------------------
    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 PDF 文件", "",
            "PDF 文件 (*.pdf);;所有文件 (*)",
        )
        if path:
            self.selected_pdf = path
            self.file_label.setText(os.path.basename(path))
            self.convert_btn.setEnabled(True)
            self.status_label.setText("")
            self._clear_preview()

    def select_size(self, label: str):
        self.selected_size = label
        self._refresh_size_buttons()
        self.size_hint.setText(f"当前选择:{FONT_SIZE_DESC[label]}")
        # 已有转换结果 → 立即用新字号重新转换并刷新预览
        if self.last_result is not None and self.selected_pdf:
            self.status_label.setText("正在按新字号重新转换…")
            self.start_convert(refresh=True)

    def select_font(self, name: str):
        self.selected_font = name
        self._refresh_font_buttons()
        self.font_hint.setText(f"当前字体:{FONT_DESC[name]}")
        # 已有转换结果 → 立即用新字体重新转换并刷新预览
        if self.last_result is not None and self.selected_pdf:
            self.status_label.setText("正在按新字体重新转换…")
            self.start_convert(refresh=True)

    def _refresh_size_buttons(self):
        for label, btn in self.size_buttons.items():
            btn.setProperty("selected", "true" if label == self.selected_size else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    def _refresh_font_buttons(self):
        for name, btn in self.font_buttons.items():
            btn.setProperty("selected", "true" if name == self.selected_font else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    def start_convert(self, refresh: bool = False):
        if not self.selected_pdf:
            return
        src = self.selected_pdf
        out_dir = os.path.join(os.path.dirname(src), OUT_DIR_NAME)
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(src))[0]
        dst = os.path.join(out_dir, f"{base}.docx")

        self._set_busy(True, refresh=refresh)
        self.progress.setVisible(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.status_label.setText("正在转换,请稍候…")

        self.thread = ConvertThread(src, dst, self.selected_size, self.selected_font, self.preview_temp_dir)
        self.thread.progress.connect(self.on_progress)
        self.thread.finished_ok.connect(self.on_done)
        self.thread.failed.connect(self.on_fail)
        self.thread.start()

    def on_progress(self, cur: int, total: int):
        self.progress.setValue(int(cur / max(total, 1) * 100))
        self.status_label.setText(f"正在转换… 第 {cur}/{total} 页")

    def on_done(self, result: dict):
        self.last_result = result
        self._set_busy(False)
        self.progress.setValue(100)
        self.progress.setVisible(False)
        img_note = f",图片 {result['images']} 张" if result["images"] else ""
        self.status_label.setText(
            f"✅ 转换完成!共 {result['pages']} 页{img_note},已保存到「{OUT_DIR_NAME}」文件夹"
        )
        self._render_preview(result)
        out_dir = os.path.dirname(result["output"])
        if sys.platform == "win32":
            os.startfile(out_dir)  # noqa: S606
        else:
            os.system(f'open "{out_dir}"')  # noqa: S605

    def on_fail(self, err: str):
        self._set_busy(False)
        self.progress.setVisible(False)
        QMessageBox.critical(self, "转换失败", f"出错了:\n{err}\n\n请确认 PDF 文件没有密码保护。")

    def _set_busy(self, busy: bool, refresh: bool = False):
        if refresh:
            # 刷新模式:字号/字体按钮保持可用,只禁转换按钮
            self.convert_btn.setEnabled(not busy and self.selected_pdf is not None)
        else:
            self.convert_btn.setEnabled(not busy and self.selected_pdf is not None)
            self.file_btn.setEnabled(not busy)
            for btn in self.size_buttons.values():
                btn.setEnabled(not busy)
            for btn in self.font_buttons.values():
                btn.setEnabled(not busy)

    # ------------------------------------------------------------------
    # 预览渲染
    # ------------------------------------------------------------------
    def _clear_preview(self):
        self.preview_browser.clear()
        self.preview_title.setVisible(False)
        self.preview_scroll.setVisible(False)

    def _render_preview(self, result: dict):
        """根据预览结构化数据渲染 HTML(模拟 Word 排版,图片保留)"""
        self._clear_preview()
        blocks = result.get("preview", [])
        if not blocks:
            return

        # 字号档位 → 像素(预览用)
        size_px = {
            "小四": 16, "四号": 18, "小三": 19, "三号": 21, "小二": 24, "二号": 28,
        }.get(result["font_size"], 21)
        font_name = result.get("font", "微软雅黑")

        parts: list[str] = []
        for block in blocks:
            if block["type"] == "image":
                path = block.get("path")
                if path and os.path.isfile(path):
                    parts.append(
                        f'<div style="text-align:center;padding:6px;">'
                        f'<img src="{path}" style="max-width:100%;"/></div>'
                    )
                continue

            text = block.get("text", "")
            if not text:
                continue
            style = block.get("style", "body")
            align = block.get("align", "left")
            para_bold = block.get("bold", False)
            spans = block.get("spans", [])

            if style == "title":
                px = size_px + 14
                weight = "bold"
            elif style == "heading":
                px = size_px + 6
                weight = "bold"
            else:
                px = size_px
                weight = "normal"

            halign = {"center": "center", "right": "right", "justify": "justify"}.get(align, "left")
            esc = html.escape(text)
            # span 级加粗:按 (文本, 加粗) 分段渲染,行内部分加粗精确显示
            if spans:
                span_html = []
                for st, sb in spans:
                    w = "bold" if sb else "normal"
                    span_html.append(
                        f'<span style="font-weight:{w}">{html.escape(st)}</span>'
                    )
                parts.append(
                    f'<div style="font-size:{px}px;font-family:{font_name};'
                    f'color:#2c3e50;text-align:{halign};padding:2px 0;">'
                    + "".join(span_html) + "</div>"
                )
            else:
                parts.append(
                    f'<div style="font-size:{px}px;font-weight:{weight};font-family:{font_name};'
                    f'color:#2c3e50;text-align:{halign};padding:2px 0;">{esc}</div>'
                )

        css = f"body {{ background: white; padding: 16px; font-family: {font_name}; }}"
        self.preview_browser.setHtml(
            f'<html><head><style>{css}</style></head><body>{"".join(parts)}</body></html>'
        )
        self.preview_title.setVisible(True)
        self.preview_scroll.setVisible(True)
        self.preview_scroll.verticalScrollBar().setValue(0)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
