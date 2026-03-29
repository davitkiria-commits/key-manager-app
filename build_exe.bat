@echo off
setlocal

if not exist venv (
    echo [INFO] Virtual environment "venv" not found. Using current Python interpreter.
)

pyinstaller --noconfirm --onefile --windowed --name key-manager main.py

echo.
echo Build completed. EXE: dist\key-manager.exe
endlocal
