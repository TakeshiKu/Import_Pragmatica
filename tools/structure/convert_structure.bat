@echo off
chcp 65001 >nul

pushd "%~dp0"

python excel_to_xml_structure.py

echo.
pause
popd