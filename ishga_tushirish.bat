@echo off
chcp 65001 > nul
echo ========================================================
echo   Telegram Nomzodlar Botini ishga tushirish...
echo ========================================================
echo.

cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    echo Bot ishga tushirilmoqda...
    echo O'chirish uchun istalgan vaqtda Ctrl+C tugmalarini bosing.
    echo.
    ".venv\Scripts\python.exe" bot.py
) else (
    echo Xatolik: Virtual muhit topilmadi.
    pause
)

pause
