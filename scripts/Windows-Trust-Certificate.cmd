@echo off
setlocal
cd /d "%~dp0.."
set "PPT_CA_FILE=%USERPROFILE%\.office-addin-dev-certs\ca.crt"
if not exist "%PPT_CA_FILE%" (
  echo Certificate not found. Run Windows-Setup.cmd first.
  pause
  exit /b 1
)
echo This trusts the existing LOCAL Office development CA for the CURRENT USER.
echo Only continue on a development PC you control. Review the certificate below.
certutil -dump "%PPT_CA_FILE%"
if errorlevel 1 goto failed
choice /c YN /n /m "Trust this local development CA for your account? [Y/N]: "
if errorlevel 2 exit /b 1
certutil -user -addstore Root "%PPT_CA_FILE%"
if errorlevel 1 goto failed
echo Certificate imported. Run Windows-Start.cmd again.
echo If the browser still shows a certificate error, close and reopen PowerPoint and browser.
pause
exit /b 0
:failed
echo Certificate import failed. If access is denied, ask your Windows administrator.
echo Do not disable TLS verification or change organization security policies.
pause
exit /b 1
