@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

set "REPO=analfox/ProtocolBot"
set "ZIP=%~dp0_update.zip"
set "TEMP=%~dp0_upd"

echo.
echo Downloading update...

powershell -NoProfile -Command "try { Invoke-WebRequest -Uri 'https://github.com/%REPO%/archive/refs/heads/main.zip' -OutFile '%ZIP%' -UseBasicParsing -ErrorAction Stop } catch { Write-Host $_.Exception.Message; exit 1 }"

if errorlevel 1 (
echo.
echo Download failed.
pause
exit /b 1
)

if not exist "%ZIP%" (
echo.
echo Update file was not created.
pause
exit /b 1
)

echo Extracting...

if exist "%TEMP%" rmdir /s /q "%TEMP%"

powershell -NoProfile -Command "Expand-Archive -Path '%ZIP%' -DestinationPath '%TEMP%' -Force"

if errorlevel 1 (
echo.
echo Extraction failed.
del /q "%ZIP%" 2>nul
pause
exit /b 1
)

echo Updating files...

for /d %%D in ("%TEMP%*") do (
xcopy "%%D*" "%~dp0" /E /I /Y /Q >nul
)

del /q "%ZIP%" 2>nul
rmdir /s /q "%TEMP%" 2>nul

echo.
echo Update completed successfully.
echo Your settings and protocols were not changed.
echo.
echo Close the program and start it again.

pause
endlocal
