@echo off
setlocal
cd /d "%~dp0"

set "APP_EXE=PPTTemplateReplacer.exe"
set "PYI_ROOT=%~dp0build\pyinstaller"
set "SPEC_FILE=%~dp0PPTTemplateReplacer.spec"
if not exist "%PYI_ROOT%" mkdir "%PYI_ROOT%"

echo [1/3] Cleaning old exe...
if exist "%~dp0%APP_EXE%" del /f /q "%~dp0%APP_EXE%"

echo [2/3] Building onefile exe with PyInstaller spec...
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --distpath "%~dp0dist" ^
  --workpath "%PYI_ROOT%\work" ^
  "%SPEC_FILE%"

if errorlevel 1 (
  echo.
  echo Build failed. Please check the logs above.
  if not defined CODEX_NO_PAUSE pause
  exit /b 1
)

echo [3/3] Copying exe to project root...
copy /y "%~dp0dist\%APP_EXE%" "%~dp0%APP_EXE%" >nul

echo.
echo Build complete:
echo   root exe: %~dp0%APP_EXE%
echo   dist exe: %~dp0dist\%APP_EXE%
if not defined CODEX_NO_PAUSE pause
