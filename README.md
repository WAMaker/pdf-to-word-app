# PDF 转 Word(老人友好版)

一款极简的 PDF 转 Word 工具,专为老人设计:**三步完成转换,改字号不乱版**。

## 核心特性

- 🖱️ **三步操作**:选文件 → 选字号 → 点转换
- 🔠 **超大界面**:大按钮、大字,无复杂选项
- 📖 **字号可选**:小四(12)到二号(22),一键切换
- 👁️ **效果预览**:转换后自动预览,切换字号即时看到调整后的排版效果
- 🖼️ **图片保留**:PDF 中的图片自动提取并嵌入 Word,预览中也显示
- 📄 **段落重建**:PDF 转 Word 后是完整逻辑段落,不是一行一段
- ✍️ **改字号不乱版**:统一使用 Word 样式控制格式,调大字号时全文自动重排,不会异常换行/缩进错乱
- 📁 **自动整理**:转换结果保存在 PDF 同目录的「PDF转Word结果」文件夹

## 为什么不会乱版?

普通工具把 PDF 的"每一行"都变成 Word 里一个独立段落,改字号后 Word 重新排版就全乱了。

本工具在转换时做**段落重建**:
1. 用 PyMuPDF 提取每行文本的位置、字号、缩进信息
2. 根据行距、缩进、文本块归属,把同一逻辑段落的行**合并成一个真正的段落**
3. 删除硬换行和手动缩进,首行缩进改用 Word 段落属性
4. 全文统一应用"正文"样式,字号由样式控制

这样老人把字号从 16 调到 22,Word 只是整体重排,段落结构保持不变。

## 开发环境

```bash
# 安装依赖
pip install -r requirements.txt

# 运行
python run_app.py

# 运行测试(生成测试 PDF 并验证段落重建)
python tests/test_converter.py
```

## Windows 打包

项目内置 GitHub Actions 自动打包流程(`.github/workflows/build-windows.yml`):

1. 推送代码到 GitHub 仓库
2. 打 tag 触发打包:`git tag v1.0.0 && git push --tags`
3. Actions 自动在 Windows 环境打包成 `PDF转Word.exe`
4. 产物自动上传为 Release 附件

也可以手动在 Windows 上打包:

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --onefile --windowed --name "PDF转Word" run_app.py
# 产物在 dist/PDF转Word.exe
```

## 项目结构

```
pdf-to-word-app/
├── app/
│   ├── __init__.py
│   ├── main.py          # PySide6 极简界面
│   └── converter.py     # 核心转换:段落重建 + docx 生成
├── tests/
│   └── test_converter.py
├── .github/workflows/
│   └── build-windows.yml
├── requirements.txt
└── run_app.py           # 入口
```

## 已知限制

- 扫描版 PDF(纯图片、无文字层)无法转换文本,需先 OCR
- 复杂排版(多栏、表格、图文混排)会简化为顺序段落
- 加密 PDF 需要先解除密码
