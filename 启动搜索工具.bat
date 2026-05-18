@echo off
chcp 65001 >nul 2>&1
title Google Batch Search
cd /d "%~dp0"
python searcher.py
if %errorlevel% neq 0 (
    echo.
    echo ========================================
    echo Start failed! Error code: %errorlevel%
    echo Please install dependencies:
    echo   pip install google openpyxl
    echo ========================================
    echo.
)
pause
