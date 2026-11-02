@echo off
chcp 65001 >nul
echo ========================================================
echo   ถอนการติดตั้ง 1-Click Zero-Prompt RDP Protocol
echo ========================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$regPath = 'HKCU:\Software\Classes\rdpconnect'; if (Test-Path $regPath) { Remove-Item -Path $regPath -Recurse -Force | Out-Null }; $targetDir = Join-Path $env:LOCALAPPDATA 'RDPManager'; if (Test-Path $targetDir) { Remove-Item -Path $targetDir -Recurse -Force | Out-Null }; Write-Host '  [SUCCESS] ถอนการติดตั้งเรียบร้อยแล้ว' -ForegroundColor Green"

echo.
pause
