@echo off
chcp 65001 > nul
title Subir cambios a GitHub

cd /d "%~dp0"

where git > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Git no encontrado en el PATH
    pause
    exit /b 1
)

for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set FECHA=%%a-%%b-%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set HORA=%%a%%b
set COMMIT_MSG=Actualizacion datos REM VIH %FECHA% %HORA%

echo.
echo Estado actual:
git status --short
echo.

echo Agregando archivos...
git add .
echo.

echo Commit: "%COMMIT_MSG%"
git commit -m "%COMMIT_MSG%"
if %errorlevel% neq 0 (
    echo [ERROR] No hay cambios para commit o fallo el commit
    pause
    exit /b 1
)

echo.
echo Subiendo a GitHub...
git push
if %errorlevel% neq 0 (
    echo [ERROR] Fallo el push a GitHub
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Cambios subidos exitosamente a GitHub!
echo ============================================
pause
