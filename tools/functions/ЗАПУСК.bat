@echo off
chcp 65001 >nul

rem Папка, где лежат GUI.py и остальные служебные файлы
set "SERVICE_DIR=%~dp0"

if not exist "%SERVICE_DIR%GUI.py" (
    echo Не найден файл GUI.py в "%SERVICE_DIR%".
    pause
    exit /b 1
)

rem Пытаемся запустить GUI без консольного окна
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" /d "%SERVICE_DIR%" pythonw "GUI.py"
) else (
    echo pythonw не найден, пробую обычный python...
    start "" /d "%SERVICE_DIR%" python "GUI.py"
    echo Если окно только моргнуло - значит python не в PATH.
    pause
)