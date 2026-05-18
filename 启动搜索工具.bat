@echo off
title Google 批量搜索工具
cd /d "%~dp0"
python searcher.py
if %errorlevel% neq 0 (
    echo.
    echo 启动失败！请确认已安装依赖：
    echo   pip install google openpyxl
    echo.
)
pause
