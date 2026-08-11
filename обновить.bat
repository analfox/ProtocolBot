@echo off
chcp 65001 >nul
cd /d "%~dp0"
set REPO=analfox/ProtocolBot
echo Скачиваю последнюю версию...
powershell -NoProfile -Command "try { Invoke-WebRequest -Uri 'https://github.com/%REPO%/archive/refs/heads/main.zip' -OutFile '_update.zip' -UseBasicParsing } catch { Write-Host 'ОШИБКА скачивания:' $_.Exception.Message; exit 1 }"
if not exist "_update.zip" (echo Не удалось скачать. Проверь интернет. & pause & exit /b)
echo Распаковываю...
powershell -NoProfile -Command "if (Test-Path '_upd') { Remove-Item '_upd' -Recurse -Force }; Expand-Archive -Path '_update.zip' -DestinationPath '_upd' -Force"
for /d %%d in (_upd\*) do (
    echo Обновляю файлы...
    xcopy "%%d\*" "%~dp0" /Y /Q >nul
)
del /q _update.zip 2>nul
if exist "_upd" rmdir /s /q _upd
echo.
echo Готово! Файлы обновлены. Твои настройки и протоколы не тронуты.
echo Закрой программу и запусти заново.
pause