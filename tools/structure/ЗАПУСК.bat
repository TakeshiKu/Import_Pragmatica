@echo off
chcp 65001 >nul

rem Пытаемся запустить GUI без консольного окна
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%~dp0GUI.py"
) else (
    echo pythonw не найден, пробую обычный python...
    start "" python "%~dp0GUI.py"
    echo Если окно только моргнуло - значит python не в PATH.
    pause
)