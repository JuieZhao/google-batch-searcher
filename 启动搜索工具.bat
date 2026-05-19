@echo off
title Google Batch Search
cd /d "%~dp0"

:: Use Python 3.14 (has PySide6, selenium, undetected-chromedriver, openpyxl)
C:\Python314\python.exe searcher.py
set ERR=%errorlevel%
if %ERR% neq 0 (
    echo.
    echo ========================================
    echo Start failed! Error code: %ERR%
    echo Try: C:\Python314\python.exe -m pip install PySide6 selenium undetected-chromedriver openpyxl
    echo ========================================
    echo.
)
pause
