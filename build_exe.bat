@echo off
setlocal

set "APP_NAME=key-manager"
set "VERSION_FILE=version.txt"

if not exist "%VERSION_FILE%" (
    echo [ERROR] Файл version.txt не найден.
    exit /b 1
)

set /p APP_VERSION=<"%VERSION_FILE%"
if "%APP_VERSION%"=="" (
    echo [ERROR] Не удалось прочитать версию из version.txt.
    exit /b 1
)

where py >nul 2>&1
if %errorlevel%==0 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set "PY_CMD=python"
    ) else (
        echo [ERROR] Python не найден.
        echo Установите Python 3.10+ и зависимости из requirements.txt.
        exit /b 1
    )
)

%PY_CMD% -c "import PyInstaller" >nul 2>&1
if not %errorlevel%==0 (
    echo [ERROR] PyInstaller не установлен.
    echo Установите зависимости: pip install -r requirements.txt
    exit /b 1
)

echo [INFO] Building %APP_NAME% version %APP_VERSION%

%PY_CMD% -m PyInstaller --noconfirm --windowed --name %APP_NAME% --distpath dist --workpath build --specpath build main.py

if not %errorlevel%==0 (
    echo [ERROR] Сборка завершилась с ошибкой.
    exit /b 1
)

echo.
echo Build completed.
echo EXE: dist\%APP_NAME%\%APP_NAME%.exe
echo Version: %APP_VERSION%
endlocal
