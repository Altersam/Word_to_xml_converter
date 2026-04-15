@echo off
cd /d "%~dp0"

echo Checking dependencies...
python -c "import PyQt5" 2>nul
if %errorlevel% neq 0 (
    echo Installing dependencies...
    python -m pip install --user PyQt5 lxml python-docx docxlatex
)

python "%~dp0converter_gui.py"
if %errorlevel% neq 0 (
    echo.
    echo [!] Error. Check Python and dependencies:
    echo     pip install PyQt5 lxml python-docx docxlatex
    pause
)
