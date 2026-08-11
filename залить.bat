@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo.
echo Adding changes...
git add -A

echo.
echo Creating commit...

for /f "delims=" %%T in ('powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH-mm-ss'"') do set "COMMIT_TIME=%%T"

git diff --cached --quiet
if not errorlevel 1 (
echo No changes to commit.
goto PUSH
)

git commit -m "update %COMMIT_TIME%"

if errorlevel 1 (
echo.
echo Commit failed.
pause
exit /b 1
)

:PUSH

echo.
echo Pushing to GitHub...

git push

if errorlevel 1 (
echo.
echo ERROR: Git push failed.
echo.
echo Check your GitHub login, remote URL and internet connection.
pause
exit /b 1
)

echo.
echo ========================================
echo        Upload completed successfully
echo ========================================
echo.

pause
endlocal
