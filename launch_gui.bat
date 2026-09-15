@echo off
setlocal
cd /d "%~dp0"

set "ROOT_EXE=%~dp0PPTTemplateReplacer.exe"
set "DIST_EXE=%~dp0dist\PPTTemplateReplacer.exe"

if exist "%ROOT_EXE%" (
  start "" "%ROOT_EXE%"
  exit /b 0
)

if exist "%DIST_EXE%" (
  start "" "%DIST_EXE%"
  exit /b 0
)

set "PYTHONW="
for %%I in (pythonw.exe) do set "PYTHONW=%%~$PATH:I"
if defined PYTHONW (
  start "" "%PYTHONW%" "%~dp0run_gui.py"
  exit /b 0
)

python "%~dp0run_gui.py"
if errorlevel 1 (
  echo.
  echo GUI 启动失败，请确认 Python 环境可用。
  pause
)
