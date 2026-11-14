@echo off
chcp 65001 >nul
cd /d %~dp0
REM 用带 tkinter 的 Python 启动器（py）打包，否则打包出的 exe 双击无反应
py -m PyInstaller --onefile --windowed --name 实用工具 --hidden-import docx --hidden-import PIL --hidden-import openpyxl 实用工具.py
echo.
echo 打包完成，产物在 dist/实用工具.exe
pause
