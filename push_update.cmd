@echo off
setlocal EnableExtensions
chcp 65001 >nul

REM ===== SETTINGS =====
set "OWNER=leonidbkc-svg"
set "REPO=epidmonitor-updater"
set "TAG=vlatest"

set "SPEC=microbio_app.spec"
set "DIST_DIR=dist\microbio_app"
set "EXE_NAME=microbio_app.exe"

echo.
echo ==========================================
echo   EpidMonitor: build ^& upload (vlatest)
echo ==========================================
echo.

where gh >nul 2>&1
if errorlevel 1 (
  echo [ERROR] gh (GitHub CLI) not found. Install it and run: gh auth login
  pause
  exit /b 1
)

where py >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python (py) not found.
  pause
  exit /b 1
)

if not exist "%SPEC%" (
  echo [ERROR] Spec file not found: %SPEC%
  pause
  exit /b 1
)

set "VER="
set /p VER=Enter version (e.g. 1.0.3): 
if "%VER%"=="" (
  echo [ERROR] Version is empty.
  pause
  exit /b 1
)

set "ZIP=EpidMonitor-%VER%.zip"
set "MANIFEST=manifest.json"

echo.
echo [1/5] PyInstaller build...
py -m PyInstaller --noconfirm --clean "%SPEC%"
if errorlevel 1 (
  echo [ERROR] PyInstaller build failed.
  pause
  exit /b 1
)

if not exist "%DIST_DIR%\%EXE_NAME%" (
  echo [ERROR] %EXE_NAME% not found in %DIST_DIR%
  pause
  exit /b 1
)

echo.
echo [2/5] Create ZIP: %ZIP%
if exist "%ZIP%" del /f /q "%ZIP%" >nul 2>&1
tar -a -c -f "%ZIP%" -C "%DIST_DIR%" .
if errorlevel 1 (
  echo [ERROR] Failed to create ZIP.
  pause
  exit /b 1
)

echo.
echo [3/5] SHA256...
for /f %%H in ('powershell -NoProfile -Command "^(Get-FileHash -Algorithm SHA256 -Path \"%ZIP%\"^).Hash"') do set "HASH=%%H"
if "%HASH%"=="" (
  echo [ERROR] Failed to compute SHA256.
  pause
  exit /b 1
)
echo SHA256: %HASH%

echo.
echo [4/5] manifest.json...
> "%MANIFEST%" (
  echo {
  echo   "app": "EpidMonitor",
  echo   "version": "%VER%",
  echo   "zip": "%ZIP%",
  echo   "sha256": "%HASH%",
  echo   "exe_relpath": "%EXE_NAME%"
  echo }
)

echo.
echo [5/5] Upload to GitHub Release %TAG%...
gh release upload %TAG% "%ZIP%" "%MANIFEST%" --clobber --repo %OWNER%/%REPO%
if errorlevel 1 (
  echo [ERROR] Upload failed. Check: gh auth status
  pause
  exit /b 1
)

echo.
echo DONE.
echo manifest: https://github.com/%OWNER%/%REPO%/releases/download/%TAG%/manifest.json
echo zip:      https://github.com/%OWNER%/%REPO%/releases/download/%TAG%/%ZIP%
echo.
pause
endlocal
