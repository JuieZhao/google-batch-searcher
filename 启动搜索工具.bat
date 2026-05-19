@echo off
title Google Batch Search
cd /d "%~dp0"

:: Use Anaconda Python (has selenium + undetected-chromedriver)
E:\anaconda3\python.exe searcher.py
set ERR=%errorlevel%
if %ERR% neq 0 (
    echo.
    echo ========================================
    echo Start failed! Error code: %ERR%
    echo Try: E:\anaconda3\python.exe -m pip install selenium undetected-chromedriver openpyxl
    echo ========================================
    echo.
)
pause
