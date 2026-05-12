@echo off
cd /d C:\Users\Avihu\Documents\cellcom_reports
title סלקום טכנאים - Tunnel

echo.
echo  ====================================
echo      סלקום טכנאים - מפעיל...
echo  ====================================
echo.

:: עצור תהליכים ישנים
taskkill /F /IM python.exe /T >nul 2>&1
taskkill /F /IM cloudflared.exe /T >nul 2>&1
del tunnel.log >nul 2>&1
timeout /t 1 >nul

:: הפעל Flask
echo [1/2] מפעיל שרת Flask...
start "Flask" /MIN python chat_server.py
timeout /t 2 >nul

:: הפעל cloudflared ושמור output לקובץ
echo [2/2] מפעיל Cloudflare Tunnel...
start "Cloudflared" /MIN cmd /c "cloudflared.exe tunnel --url http://localhost:5000 2> tunnel.log"

:: המתן עד שה-URL מופיע בלוג
echo ממתין לכתובת...
:wait_loop
timeout /t 2 >nul
findstr "trycloudflare.com" tunnel.log >nul 2>&1
if errorlevel 1 goto wait_loop

echo.
echo  ====================================
echo   הכתובת שלך לטלפון:
echo.
for /f "tokens=*" %%a in ('findstr "trycloudflare.com" tunnel.log') do (
    echo   %%a
)
echo.
echo  ====================================
echo.
echo  שלח/י את הכתובת לטלפון ופתח אותה!
echo  (אל תסגור חלון זה - האפליקציה תפסיק לעבוד)
echo.
pause
