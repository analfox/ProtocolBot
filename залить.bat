@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Собираю изменения...
git add -A
git commit -m "update %date% %time%" >nul 2>&1
echo Заливаю на GitHub...
git push
if %errorlevel%==0 (echo. & echo Готово! Проект залит на GitHub.) else (echo. & echo Ошибка. Если просит вход - залогинься в GitHub.)
pause