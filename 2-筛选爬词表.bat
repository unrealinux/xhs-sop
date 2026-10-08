@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" (
  echo Usage: drag an xlsx/csv file onto this .bat, or type the path below.
  set /p F="File path: "
  py filter_scrape.py "%F%"
) else (
  py filter_scrape.py %*
)
echo.
pause
