@echo off
title Google Batch Search
cd /d "%~dp0"
python searcher.py
set ERR=%errorlevel%
if %ERR% neq 0 (
    echo.
    echo ========================================
    echo Start failed! Error code: %ERR%
    echo Try: pip install requests beautifulsoup4 openpyxl
    echo ========================================
    echo.
)
pause
