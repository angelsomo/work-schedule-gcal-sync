@echo off
setlocal
cd /d "%~dp0"

REM Always use the project's venv python
".venv\Scripts\python.exe" "extract_to_gcal.py"

echo.
echo Finished. Press any key to close...
pause >nul
endlocal