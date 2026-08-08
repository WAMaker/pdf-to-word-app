# -*- coding: utf-8 -*-
"""
极简 PDF 转 Word 界面(老人友好设计)
-------------------------------------
- 全流程 3 步:选文件 → 选字号 → 开始转换
- 大按钮、大字,避免复杂选项
"""
from __future__ import annotations

import os
import sys
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
OUT_DIR_NAME = "PDF转Word结果"


class ConvertThread(QThread):
    """后台转换线程,避免界面卡死"""
    progress = Signal(int, int)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, pdf_path, docx_path, font_size):
        super().__init__()
        self.pdf_path = pdf_path
        self.docx_path = docx_path
        self.font_size = font_size

    def run(self):
        try:
            result = convert_pdf_to_docx(
                self.pdf_path,
                self.docx_path,
                font_size_label=self.font_size,
                progress_cb=lambda cur, total: self.progress.emit(cur, total),
            )
            self.finished_ok.emit(result)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(720, 640)
        self.selected_pdf: str | None = None
        self.thread: ConvertThread | None = None

        self._build_ui()
        self._apply_fonts()

    # ------------------------------------------------------------------
    # UI 构建
    # ------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(40, 40, 40, 40)
        root.setSpacing(20)

        # 标题
        title = QLabel("PDF 转 Word")
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("title")
        root.addWidget(title)

        # 第一步:选择文件
        root.addWidget(self._section_label("① 选择 PDF 文件"))
        self.file_btn = QPushButton("📄  点击选择 PDF 文件")
        self.file_btn.setObjectName("fileBtn")
        self.file_btn.setMinimumHeight(64)
        self.file_btn.clicked.connect(self.choose_file)
        root.addWidget(self.file_btn)

        self.file_label = QLabel("尚未选择文件")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setWordWrap(True)
        self.file_label.setObjectName("fileLabel")
        root.addWidget(self.file_label)

        # 第二步:选择字号
        root.addWidget(self._section_label("② 选择文字大小"))
        size_row = QHBoxLayout()
        size_row.setSpacing(12)
        self.size_buttons: dict[str, QPushButton] = {}
        self.selected_size = "三号"
        for label in FONT_SIZES:
            btn = QPushButton(label)
            btn.setObjectName("sizeBtn")
            btn.setMinimumHeight(56)
            btn.setToolTip(FONT_SIZE_DESC[label])
            btn.clicked.connect(lambda _=False, s=label: self.select_size(s))
            size_row.addWidget(btn)
            self.size_buttons[label] = btn
        root.addLayout(size_row)

        self.size_hint = QLabel(f"当前选择:{FONT_SIZE_DESC[self.selected_size]}")
        self.size_hint.setAlignment(Qt.AlignCenter)
        self.size_hint.setObjectName("hint")
        root.addWidget(self.size_hint)

        # 第三步:转换
        root.addWidget(self._section_label("③ 开始转换"))
        self.convert_btn = QPushButton("🚀  开始转换")
        self.convert_btn.setObjectName("convertBtn")
        self.convert_btn.setMinimumHeight(72)
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.start_convert)
        root.addWidget(self.convert_btn)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setMinimumHeight(28)
        root.addWidget(self.progress)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setObjectName("status")
        root.addWidget(self.status_label)

        root.addStretch(1)

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
                font-size: 34px; font-weight: bold;
                color: #2c3e50; padding: 8px;
            }
            QLabel#section {
                font-size: 18px; font-weight: bold; color: #7f8c8d;
                padding-top: 8px;
            }
            QLabel#fileLabel { font-size: 15px; color: #34495e; }
            QLabel#hint { font-size: 15px; color: #16a085; }
            QLabel#status { font-size: 16px; color: #2c3e50; }
            QPushButton#fileBtn {
                font-size: 22px; background: #ecf0f1; color: #2c3e50;
                border-radius: 12px; border: 2px dashed #bdc3c7;
            }
            QPushButton#fileBtn:hover { background: #e8f8f5; border-color: #16a085; }
            QPushButton#sizeBtn {
                font-size: 20px; background: #ecf0f1; color: #2c3e50;
                border-radius: 10px; border: 2px solid #bdc3c7;
            }
            QPushButton#sizeBtn:hover { background: #e8f8f5; }
            QPushButton#sizeBtn[selected="true"] {
                background: #16a085; color: white; border-color: #16a085;
                font-weight: bold;
            }
            QPushButton#convertBtn {
                font-size: 24px; font-weight: bold;
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
        """)
        self._refresh_size_buttons()

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

    def select_size(self, label: str):
        self.selected_size = label
        self._refresh_size_buttons()
        self.size_hint.setText(f"当前选择:{FONT_SIZE_DESC[label]}")

    def _refresh_size_buttons(self):
        for label, btn in self.size_buttons.items():
            btn.setProperty("selected", "true" if label == self.selected_size else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    def start_convert(self):
        if not self.selected_pdf:
            return
        src = self.selected_pdf
        out_dir = os.path.join(os.path.dirname(src), OUT_DIR_NAME)
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(src))[0]
        dst = os.path.join(out_dir, f"{base}.docx")

        self._set_busy(True)
        self.progress.setVisible(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.status_label.setText("正在转换,请稍候…")

        self.thread = ConvertThread(src, dst, self.selected_size)
        self.thread.progress.connect(self.on_progress)
        self.thread.finished_ok.connect(self.on_done)
        self.thread.failed.connect(self.on_fail)
        self.thread.start()

    def on_progress(self, cur: int, total: int):
        self.progress.setValue(int(cur / max(total, 1) * 100))
        self.status_label.setText(f"正在转换… 第 {cur}/{total} 页")

    def on_done(self, result: dict):
        self._set_busy(False)
        self.progress.setValue(100)
        self.status_label.setText(
            f"✅ 转换完成!共 {result['pages']} 页,已保存到「{OUT_DIR_NAME}」文件夹"
        )
        out_dir = os.path.dirname(result["output"])
        # 转换完成后自动打开输出文件夹
        if sys.platform == "win32":
            os.startfile(out_dir)  # noqa: S606
        else:
            os.system(f'open "{out_dir}"')  # noqa: S605

    def on_fail(self, err: str):
        self._set_busy(False)
        self.progress.setVisible(False)
        QMessageBox.critical(self, "转换失败", f"出错了:\n{err}\n\n请确认 PDF 文件没有密码保护。")

    def _set_busy(self, busy: bool):
        self.convert_btn.setEnabled(not busy and self.selected_pdf is not None)
        self.file_btn.setEnabled(not busy)
        for btn in self.size_buttons.values():
            btn.setEnabled(not busy)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
