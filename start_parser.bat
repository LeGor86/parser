@echo off
chcp 65001 >nul

REM Автоматический запуск парсера

REM Установка зависимостей
echo Установка необходимых зависимостей...
pip install requests beautifulsoup4

REM Проверка успешности установки
if %errorlevel% neq 0 (
    echo Ошибка установки зависимостей
    pause
    exit /b %errorlevel%
)

echo.
echo Выберите режим запуска:
echo 1 - Веб-интерфейс (сервер)
echo 2 - Мониторинг одного URL в консоли
set /p MODE="Введите 1 или 2: "

if "%MODE%"=="2" (
    set /p URL="Введите URL профиля для мониторинга: "
    if defined URL (
        set /p INTERVAL="Интервал проверки в секундах (по умолчанию 300): "
        if not defined INTERVAL set INTERVAL=300
        echo Запуск мониторинга...
        python status_to_db.py "%URL%" --interval %INTERVAL%
    )
) else (
    echo Запуск веб-сервера...
    python site_parser.py --serve
)

pause