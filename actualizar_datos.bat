@echo off
chcp 65001 > nul
title Actualizar datos REM VIH 2025

echo ============================================
echo  Actualizando datos REM VIH 2025
echo  %date% - %time%
echo ============================================
echo.

cd /d "%~dp0"

where python > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no encontrado en el PATH
    pause
    exit /b 1
)

echo [1/2] Extrayendo datos REM (A05, A11, P1, P11)...
python "2025\extraer_rem_vih_2025.py"
if %errorlevel% neq 0 (
    echo [ERROR] Fallo la extraccion de datos REM
    pause
    exit /b 1
)
echo.

echo [2/2] Generando comparativos P11...
python "2025\comparar_rem_p11_cortes_2025.py"
if %errorlevel% neq 0 (
    echo [ERROR] Fallo la generacion de comparativos P11
    pause
    exit /b 1
)
echo.

echo ============================================
echo  Actualizacion completada exitosamente
echo ============================================
echo.
echo Archivos generados:
echo  - 2025\salida\A05_2025.xlsx
echo  - 2025\salida\A11_2025.xlsx
echo  - 2025\salida\P1_2025.xlsx
echo  - 2025\salida\P11_2025.xlsx
echo  - 2025\salida_comparativo_p11\P11_2025_comparativo_mes06_vs_mes12.xlsx
echo  - 2025\salida_comparativo_p11\P11_2025_informe_alertas_mes06_vs_mes12.xlsx
echo.

pause
