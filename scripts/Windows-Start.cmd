@echo off
setlocal
cd /d "%~dp0.."
if not exist node_modules (
  echo Run Windows-Setup.cmd first.
  pause
  exit /b 1
)
echo Keep this window open while using PowerPoint. Follow docs\WINDOWS.md.
call npm run dev
pause
