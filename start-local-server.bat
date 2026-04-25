@echo off
setlocal

REM Always run from the project root (folder of this file)
cd /d "%~dp0"

REM Default port (can be changed by editing this value)
set PORT=8000

echo Starting local server in:
echo %cd%
echo.
echo Open in browser: http://localhost:%PORT%
echo Press Ctrl+C in this window to stop the server.
echo.
start "" "http://localhost:%PORT%/data/report/offload_report.html"

REM Prefer Python launcher on Windows
where py >nul 2>nul
if %errorlevel%==0 (
    py -m http.server %PORT%
    goto :end
)

REM Fallback to python command
where python >nul 2>nul
if %errorlevel%==0 (
    python -m http.server %PORT%
    goto :end
)

echo Python is not installed or not in PATH.
echo Please install Python, then run this file again.
pause

:end
endlocal
