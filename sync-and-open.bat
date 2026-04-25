@echo off
setlocal
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0sync-and-open.ps1"
if errorlevel 1 (
  echo.
  echo 同期に失敗しました。
  pause
)
