@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 --version >nul 2>nul
if %errorlevel%==0 (
  py -3 "%~dp0web_ui.py" %*
) else (
  python "%~dp0web_ui.py" %*
)
if errorlevel 1 (
  echo 页面启动失败，请确认已安装 Python 3.10 或更新版本。
  pause
)
