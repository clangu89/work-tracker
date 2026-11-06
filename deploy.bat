@echo off
chcp 65001 >nul
echo =========================================================================
echo   Work Tracker - 1-Click Production Deploy Script (Clasp + GitHub)
echo =========================================================================
echo.

cd /d "%~dp0"

echo [1/3] กำลังคอมไพล์ Frontend ด้วย build.js...
call node build.js
if %errorlevel% neq 0 (
    echo [ERROR] การคอมไพล์ล้มเหลว! ขอยกเลิกการ Deploy
    pause
    exit /b %errorlevel%
)
echo.

echo [2/3] กำลังซิงค์ขึ้น Google Apps Script (Clasp push & deploy)...
call clasp push -f
if %errorlevel% neq 0 (
    echo [ERROR] Clasp push ล้มเหลว!
    pause
    exit /b %errorlevel%
)

call clasp deploy -i AKfycbyox_uV2XJjgn96jK-XMKlr4vgucWG9FD1WqJR6Rhvu-L82El1nNhq7LyNvtwEBSjlrkg -d "Production deploy %date% %time%"
if %errorlevel% neq 0 (
    echo [WARN] Clasp redeploy ล้มเหลว แต่โค้ดล่าสุดถูก push ขึ้นเรียบร้อยแล้ว
)
echo.

echo [3/3] กำลังซิงค์ขึ้น GitHub Pages (Git commit & push)...
git add .
set commit_msg=Update Work Tracker production - %date% %time%
git commit -m "%commit_msg%"
git push origin main
if %errorlevel% neq 0 (
    echo [ERROR] Git push ล้มเหลว! กรุณาตรวจสอบอินเทอร์เน็ตหรือ Credentials
    pause
    exit /b %errorlevel%
)

echo.
echo =========================================================================
echo   [สำเร็จ] นำขึ้นระบบจริงเรียบร้อยแล้วทั้ง 2 ช่องทาง!
echo   - Web App URL: https://script.google.com/macros/s/AKfycbyox_uV2XJjgn96jK-XMKlr4vgucWG9FD1WqJR6Rhvu-L82El1nNhq7LyNvtwEBSjlrkg/exec
echo   - GitHub Pages: https://clangu89.github.io/work-tracker/
echo =========================================================================
echo.
pause
