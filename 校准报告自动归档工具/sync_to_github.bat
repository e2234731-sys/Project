@echo off
chcp 65001 >nul
echo ========================================================
echo   一键同步当前项目至 GitHub 仓库 (e2234731-sys/Project)
echo ========================================================
python sync_project.py
pause
