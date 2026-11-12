# Markdown 转 Word 离线工具

一个面向 Windows 的本地 Markdown → Word 转换工具。通过简单页面选择文件、Word 模板和输出目录，无需账号或 API Key，转换时不上传文件。

## 功能

- 批量转换 `.md` / `.markdown`，支持移除已选文件。
- 提供多种中文 Word 模板，也可选择自定义 `.docx` 模板。
- 默认保存到当前用户的系统“下载”目录，也可自行选择并记住其他输出位置。
- 自动检测 Python 包和 Pandoc，缺少时一键安装；已准备好的依赖不会重复安装。
- 展示转换进度、结果和日志，并生成转换报告；同名文件自动添加序号，不覆盖源文件。
- 使用工具专属浏览器窗口，点击右上角 **×** 后同步结束后台服务及启动终端。

## 环境准备

1. 安装 **Python 3.10+**，需包含 `pip` 和 `tkinter`（Windows 官方安装版通常自带）。
2. 推荐使用 **Microsoft Edge 或 Google Chrome**，用于打开工具专属窗口。

### 两种获取方式

- **源码版**：下载或克隆仓库。首次点击“安装依赖”，会安装缺失的 Python 包，并从官方 GitHub 下载 Pandoc 3.6.4 Windows 64 位便携版到 `bin/`；下载包会进行 SHA-256 校验。
- **完整版**：在 GitHub Releases 下载 `Markdown-to-Word-Windows-full.zip` 并解压，已包含 Pandoc 和版权文件，首次只需安装缺失的 Python 包。

两种版本都需要先安装 Python。依赖准备需要联网，准备完成后的转换可离线运行。源码仓库不提交 Pandoc 可执行文件或发布 ZIP。

## 使用

1. 下载或克隆仓库，双击 **`启动页面.cmd`**。
2. 如出现“安装依赖”，点击安装并等待完成。
3. 在 **01** 选择一个或多个 Markdown 文件。
4. 在 **02** 确认或选择 Word 模板。
5. 在 **03** 确认或选择输出目录，点击 **开始转换**。
6. 完成后点击“打开输出目录”查看结果；关闭工具窗口即可退出。

“安装依赖”只补齐缺失项，全部就绪后安装区域自动隐藏。下载失败可查看日志并重试，或[手动安装官方 Pandoc](https://pandoc.org/installing.html)。没有 Edge/Chrome 时会回退到默认浏览器，此时不能保证关闭标签页会同步退出后台，可在启动终端按 `Ctrl+C` 结束。

### 手动安装依赖

在项目目录中打开 PowerShell：

```powershell
py -3 -m pip install -r requirements.txt
```

### 命令行转换（可选）

```powershell
py -3 md_to_word.py "D:\文档\示例.md" -d "D:\Word输出"
```

`config.json` 中 `output_dir` 留空表示使用系统下载目录；`template` 可设置默认 Word 模板。

## 测试

```powershell
py -3 -m unittest discover
```

## 发布完整版

在已准备好 `bin/pandoc.exe` 的项目目录运行：

```powershell
py -3 build_release.py
```

将 `dist/` 中生成的 ZIP 和 `.sha256` 文件上传到 GitHub Releases。打包脚本会保留运行文件、模板和 Pandoc 版权文件，排除个人路径记录、缓存、测试和 Git 文件，并重置发布包的默认配置。

Pandoc 的版权及许可说明保留在 `bin/COPYRIGHT.txt` 和 `bin/COPYING.rtf`。随便携版分发 Pandoc 时，应一并保留这些文件。
