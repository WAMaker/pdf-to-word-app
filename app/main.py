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
import urllib.request
import json

from PySide6.QtCore import Qt, QThread, Signal, QSettings, QLockFile, QUrl
from PySide6.QtGui import QFont, QDesktopServices
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

# 单实例锁 + 设置持久化
APP_ORG = "PDF2Word"
APP_SETTINGS_KEY = "pdf2word"
GITHUB_REPO = "WAMaker/pdf-to-word-app"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
# 当前版本(打包 workflow 会替换此行保持同步)
APP_VERSION = "v1.4.0"
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


class UpdateThread(QThread):
    """后台检查 GitHub 最新版本"""
    ok = Signal(str, str, str, int)  # (最新版本号, 下载页URL, exe下载URL, 文件大小bytes)
    failed = Signal(str)             # 错误信息(网络/解析等)

    def __init__(self, current_version: str):
        super().__init__()
        self.current_version = current_version

    def run(self):
        try:
            req = urllib.request.Request(
                GITHUB_API_LATEST,
                headers={"User-Agent": "PDF2Word-Updater", "Accept": "application/vnd.github+json"},
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            latest = (data.get("tag_name") or "").strip()
            if not latest:
                self.failed.emit("无法获取版本信息,请稍后再试")
                return
            url = data.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases"
            # 找 exe 资产(下载更新用)
            asset_url = ""
            asset_size = 0
            for asset in (data.get("assets") or []):
                name = (asset.get("name") or "").lower()
                if name.endswith(".exe") and not name.endswith(".exe.asc"):
                    asset_url = asset.get("browser_download_url") or ""
                    asset_size = int(asset.get("size") or 0)
                    break
            self.ok.emit(latest, url, asset_url, asset_size)
        except Exception as e:  # noqa: BLE001
            self.failed.emit(f"检查更新失败(网络或服务器问题),请稍后再试\n{type(e).__name__}")


class DownloadThread(QThread):
    """后台下载更新包,带进度"""
    progress = Signal(int, int)  # (已下载bytes, 总bytes)
    finished_ok = Signal(str)    # 下载完成,返回本地临时文件路径
    failed = Signal(str)

    def __init__(self, url: str, dest_path: str):
        super().__init__()
        self.url = url
        self.dest_path = dest_path

    def run(self):
        try:
            req = urllib.request.Request(
                self.url,
                headers={"User-Agent": "PDF2Word-Updater"},
            )
            with urllib.request.urlopen(req, timeout=60) as resp:
                total = int(resp.headers.get("Content-Length") or 0)
                downloaded = 0
                with open(self.dest_path, "wb") as f:
                    while True:
                        chunk = resp.read(65536)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        self.progress.emit(downloaded, total)
            self.finished_ok.emit(self.dest_path)
        except Exception as e:  # noqa: BLE001
            # 清理半成品
            try:
                if os.path.exists(self.dest_path):
                    os.remove(self.dest_path)
            except Exception:  # noqa: BLE001
                pass
            self.failed.emit(f"下载失败(网络或服务器问题),请稍后再试\n{type(e).__name__}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(860, 760)
        self.selected_pdf: str | None = None
        self.thread: ConvertThread | None = None
        self.update_thread: UpdateThread | None = None
        self.download_thread: DownloadThread | None = None
        self.last_result: dict | None = None
        self.preview_temp_dir = tempfile.mkdtemp(prefix="pdf2word_")
        # 记住用户的字体/字号选项(QSettings 持久化)
        self.settings = QSettings(APP_ORG, APP_SETTINGS_KEY)
        saved_size = self.settings.value("font_size", "三号")
        saved_font = self.settings.value("font_name", "微软雅黑")
        self.selected_size = saved_size if saved_size in FONT_SIZES else "三号"
        self.selected_font = saved_font if saved_font in FONTS else "微软雅黑"
        # 当前版本(更新检查用;打包时由构建脚本写入)
        self.app_version = APP_VERSION

        self._build_ui()
        self._apply_fonts()
        self.setAcceptDrops(True)
        self._refresh_size_buttons()
        self._refresh_font_buttons()

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

        # 更新检查(小按钮,放转换按钮下方)
        self.update_btn = QPushButton("🔄  检查更新")
        self.update_btn.setObjectName("updateBtn")
        self.update_btn.setMinimumHeight(40)
        self.update_btn.setToolTip("去 GitHub 检查是否有新版本")
        self.update_btn.clicked.connect(self.check_update)
        root.addWidget(self.update_btn)

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
            QPushButton#updateBtn {
                font-size: 15px; background: #f8f9fa; color: #7f8c8d;
                border-radius: 8px; border: 1px solid #bdc3c7;
                padding: 4px;
            }
            QPushButton#updateBtn:hover { background: #e8f8f5; color: #16a085; }
            QPushButton#updateBtn:disabled { color: #bdc3c7; }
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
            self._set_pdf(path)

    def select_size(self, label: str):
        self.selected_size = label
        self.settings.setValue("font_size", label)  # 记住选项
        self._refresh_size_buttons()
        self.size_hint.setText(f"当前选择:{FONT_SIZE_DESC[label]}")
        # 已有转换结果 → 立即用新字号重新转换并刷新预览
        if self.last_result is not None and self.selected_pdf:
            self.status_label.setText("正在按新字号重新转换…")
            self.start_convert(refresh=True)

    def select_font(self, name: str):
        self.selected_font = name
        self.settings.setValue("font_name", name)  # 记住选项
        self._refresh_font_buttons()
        self.font_hint.setText(f"当前字体:{FONT_DESC[name]}")
        # 已有转换结果 → 立即用新字体重新转换并刷新预览
        if self.last_result is not None and self.selected_pdf:
            self.status_label.setText("正在按新字体重新转换…")
            self.start_convert(refresh=True)

    # ------------------------------------------------------------------
    # 拖放支持:把 PDF 文件直接拖进窗口
    # ------------------------------------------------------------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return
        path = urls[0].toLocalFile()
        if path.lower().endswith(".pdf") and os.path.isfile(path):
            self._set_pdf(path)
        else:
            QMessageBox.warning(self, "不支持的文件", "请拖入 PDF 文件。")

    def _set_pdf(self, path: str):
        self.selected_pdf = path
        self.file_label.setText(os.path.basename(path))
        self.convert_btn.setEnabled(True)
        self.status_label.setText("")
        self._clear_preview()

    # ------------------------------------------------------------------
    # 检查更新:去 GitHub 查最新版本,下载并替换当前 exe
    # ------------------------------------------------------------------
    def check_update(self):
        if self.update_thread and self.update_thread.isRunning():
            return
        self.update_btn.setEnabled(False)
        self.update_btn.setText("检查中…")
        self.update_thread = UpdateThread(self.app_version)
        self.update_thread.ok.connect(self.on_update_ok)
        self.update_thread.failed.connect(self.on_update_fail)
        self.update_thread.finished.connect(self.on_update_done)
        self.update_thread.start()

    def on_update_ok(self, latest: str, url: str, asset_url: str, asset_size: int):
        cur = self.app_version
        if latest.lower().lstrip("v") == cur.lower().lstrip("v"):
            QMessageBox.information(self, "检查更新", f"当前已是最新版本 ({cur}) ✓")
            return
        if not asset_url:
            QMessageBox.information(
                self, "发现新版本",
                f"发现新版本 {latest}\n当前版本 {cur}\n\n该版本未提供 exe 安装包,可打开下载页获取。",
            )
            return
        size_txt = f"{asset_size / 1024 / 1024:.1f} MB" if asset_size else ""
        box = QMessageBox(self)
        box.setWindowTitle("发现新版本")
        box.setText(
            f"发现新版本 {latest}\n当前版本 {cur}\n安装包大小:{size_txt}\n\n是否现在下载并更新?"
        )
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.button(QMessageBox.Yes).setText("下载更新")
        box.button(QMessageBox.No).setText("稍后再说")
        if box.exec() == QMessageBox.Yes:
            self._start_download(latest, asset_url)

    def _start_download(self, latest: str, asset_url: str):
        """开始下载更新包(带进度条)"""
        dest = os.path.join(tempfile.gettempdir(), f"pdf2word_update_{latest}.exe")
        self.status_label.setText(f"正在下载更新 {latest} …")
        self.progress.setVisible(True)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.update_btn.setEnabled(False)
        self.update_btn.setText("下载中…")
        self.download_thread = DownloadThread(asset_url, dest)
        self.download_thread.progress.connect(self.on_download_progress)
        self.download_thread.finished_ok.connect(self.on_download_done)
        self.download_thread.failed.connect(self.on_download_fail)
        self.download_thread.start()

    def on_download_progress(self, downloaded: int, total: int):
        if total > 0:
            pct = int(downloaded / total * 100)
            self.progress.setValue(pct)
            mb_d = downloaded / 1024 / 1024
            mb_t = total / 1024 / 1024
            self.status_label.setText(f"正在下载更新… {pct}% ({mb_d:.1f}/{mb_t:.1f} MB)")
        else:
            self.progress.setValue(0)
            self.status_label.setText("正在下载更新…")

    def on_download_done(self, file_path: str):
        self.progress.setValue(100)
        self.status_label.setText("下载完成")
        self._confirm_replace(file_path)

    def on_download_fail(self, msg: str):
        self.progress.setVisible(False)
        self.status_label.setText("")
        QMessageBox.warning(self, "更新失败", msg)

    def _confirm_replace(self, downloaded_exe: str):
        """确认替换当前 exe(替换后自动重启新版本)"""
        if not getattr(sys, "frozen", False):
            # 源码运行(非打包 exe):无法自动替换,提示手动
            QMessageBox.information(
                self, "更新",
                f"下载完成: {downloaded_exe}\n\n当前为源码运行,请手动替换 exe 文件。",
            )
            self.progress.setVisible(False)
            return
        cur_exe = os.path.abspath(sys.executable)
        box = QMessageBox(self)
        box.setWindowTitle("更新")
        box.setText(
            f"新版本已下载完成。\n\n将替换当前程序:\n{cur_exe}\n\n替换后会自动启动新版本,继续吗?"
        )
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.button(QMessageBox.Yes).setText("替换并重启")
        box.button(QMessageBox.No).setText("取消")
        if box.exec() == QMessageBox.Yes:
            self._do_replace(downloaded_exe, cur_exe)
        else:
            self.progress.setVisible(False)
            QMessageBox.information(
                self, "更新",
                f"新版本已保存到:\n{downloaded_exe}\n\n下次可直接用它替换旧版。",
            )

    def _do_replace(self, downloaded_exe: str, cur_exe: str):
        """执行替换:写批处理脚本等待进程退出后覆盖 exe 并重启"""
        if sys.platform != "win32":
            QMessageBox.information(
                self, "更新",
                f"新版本已保存到:\n{downloaded_exe}\n\n当前系统不支持自动替换,请手动替换。",
            )
            self.progress.setVisible(False)
            return
        try:
            exe_dir = os.path.dirname(cur_exe)
            new_name = os.path.basename(cur_exe)
            staged = os.path.join(exe_dir, f"{new_name}.new")
            # 把下载的 exe 复制到程序目录(避免临时目录权限/清理问题)
            import shutil
            shutil.copy2(downloaded_exe, staged)
            # 批处理:等待旧进程退出 → 覆盖 → 启动新版 → 清理
            bat_path = os.path.join(exe_dir, "_update_apply.bat")
            exe_name = os.path.basename(cur_exe)
            with open(bat_path, "w", encoding="gbk", errors="replace") as f:
                f.write('@echo off\r\n')
                f.write('chcp 65001 >nul\r\n')
                f.write(f'title Updating {APP_TITLE}...\r\n')
                f.write(f'set "OLD={cur_exe}"\r\n')
                f.write(f'set "NEW={staged}"\r\n')
                f.write(':wait\r\n')
                # 等主程序退出(找不到该进程名即已退出)
                f.write(f'tasklist /FI "IMAGENAME eq {exe_name}" 2>nul | find /I "{exe_name}" >nul 2>&1\r\n')
                f.write('if not errorlevel 1 (timeout /t 2 /nobreak >nul & goto wait)\r\n')
                f.write('copy /Y "%NEW%" "%OLD%" >nul\r\n')
                f.write('if errorlevel 1 (echo 替换失败 & pause & exit /b 1)\r\n')
                f.write('del /F /Q "%NEW%"\r\n')
                f.write('del /F /Q "%~f0"\r\n')
                f.write(f'start "" "{cur_exe}"\r\n')
            os.startfile(bat_path)  # noqa: S606
            # 提示后退出(批处理会完成替换并重启)
            QMessageBox.information(
                self, "更新", "正在替换程序,窗口将关闭。\n新版本会自动启动。"
            )
            QApplication.instance().quit()
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "更新失败", f"替换失败:\n{e}\n\n新版本已保存到:\n{downloaded_exe}")
            self.progress.setVisible(False)

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


def _single_instance_lock() -> QLockFile | None:
    """单实例锁:已有实例在运行时返回 None(第二实例直接退出)
    锁文件放临时目录,随系统清理。
    """
    lock_path = os.path.join(
        tempfile.gettempdir(), f"{APP_ORG}_{APP_SETTINGS_KEY}.lock"
    )
    lock = QLockFile(lock_path)
    lock.setStaleLockTime(0)  # 崩溃残留锁也视为过期(跨进程,0=不自动过期)
    if not lock.tryLock(50):
        return None
    return lock


def main():
    app = QApplication(sys.argv)
    # 单实例:已有窗口在运行则提示并退出(双击不会开第二个窗口)
    lock = _single_instance_lock()
    if lock is None:
        QMessageBox.information(
            None, "程序已在运行", "PDF 转 Word 已经在运行了,请在已有窗口中使用。"
        )
        return

    win = MainWindow()
    win.show()
    # 锁随 app 生命周期持有,退出时自动释放
    app.aboutToQuit.connect(lock.unlock)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
