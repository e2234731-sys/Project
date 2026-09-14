@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "dist\校准报告智能归档系统\校准报告智能归档系统.exe" (
    start "" "dist\校准报告智能归档系统\校准报告智能归档系统.exe"
) else (
    echo 正在启动 Python GUI 客户端...
    python main_gui.py
)
