@echo off
setlocal

where py >nul 2>&1
if %errorlevel%==0 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set "PY_CMD=python"
    ) else (
        echo [ERROR] Python не найден.
        echo Установите Python 3.10+ и добавьте его в PATH.
        echo Затем установите зависимости: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

%PY_CMD% -c "import PySide6" >nul 2>&1
if not %errorlevel%==0 (
    echo [ERROR] Не найдены зависимости приложения.
    echo Выполните команду: pip install -r requirements.txt
    pause
    exit /b 1
)

%PY_CMD% main.py
set EXIT_CODE=%errorlevel%

if not %EXIT_CODE%==0 (
    echo.
    echo [ERROR] Приложение завершилось с кодом %EXIT_CODE%.
    pause
)

endlocal & exit /b %EXIT_CODE%
