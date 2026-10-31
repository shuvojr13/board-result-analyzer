@echo off
cd /d "%~dp0"
title Board Result Analyzer
echo ========================================================
echo  Starting Bangladesh Board Result Analyzer...
echo  Local URL: http://localhost:8000
echo ========================================================
echo.
.\venv\Scripts\python.exe app.py
pause
