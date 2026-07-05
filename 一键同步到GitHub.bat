@echo off
color 0A
echo =========================================
echo       One-Click GitHub Sync Tool
echo =========================================
echo.

cd /d "%~dp0"

echo [1/3] Scanning for changed files...
git add .

set /p msg="Enter commit message (Press Enter to use current time): "
if "%msg%"=="" (
    set msg=Auto update %date% %time:~0,8%
)

echo.
echo [2/3] Committing changes: [%msg%]
git commit -m "%msg%"

echo.
echo [3/3] Pushing to GitHub...
git push

echo.
if %errorlevel% equ 0 (
    echo =========================================
    echo          Sync Successful!
    echo =========================================
) else (
    color 0C
    echo =========================================
    echo     Sync Failed. Please check errors.
    echo =========================================
)

echo.
pause
