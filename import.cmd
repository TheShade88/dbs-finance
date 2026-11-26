@echo off
rem Double-click after downloading a DBS export. Close the workbook in Excel first.
cd /d "%~dp0"
uv run finance run
echo.
pause
