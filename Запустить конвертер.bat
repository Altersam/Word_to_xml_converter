@echo off
cd /d "%~dp0"
py -3.12 "%~dp0converter_gui.py"
if %errorlevel% neq 0 (
    echo.
    echo [!] Error. Check Python and dependencies:
    echo     pip install PyQt5 lxml python-docx docxlatex
    pause
)
