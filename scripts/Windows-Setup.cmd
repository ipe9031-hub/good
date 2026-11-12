@echo off
setlocal
cd /d "%~dp0.."
where node >nul 2>&1
if errorlevel 1 (
  echo Install Node.js 22 LTS from https://nodejs.org and run again.
  pause
  exit /b 1
)
call npm ci --no-audit --no-fund
if errorlevel 1 goto failed
call npm run certs
if errorlevel 1 goto failed
call npm run build
if errorlevel 1 goto failed
echo Setup complete. Run Windows-Start.cmd and follow docs\WINDOWS.md.
pause
exit /b 0
:failed
echo Setup failed. Read the error above. Do not bypass certificate verification.
pause
exit /b 1
