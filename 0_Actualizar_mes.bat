@echo off
rem ============================================================================
rem  Actualizacion mensual del Plan Estrategico 2026-2030
rem   1. 1_Extraer_indicadores.py  : fuentes de informacion -> Tablero_Indicadores_PE.xlsx
rem   2. digitacion manual           : IPC, fallecidos, encuestas, formacion, canales...
rem   3. 3_Generar_dashboard.py: Excel -> docs/index.html y publicacion en GitHub
rem ============================================================================
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo === 1/3  Extrayendo indicadores de las fuentes de informacion ===
python 1_Extraer_indicadores.py
if errorlevel 1 goto error

echo.
echo === 2/3  Datos manuales ===
echo Se abre el tablero: digite lo que no sale de las fuentes (IPC del mes y corrido,
echo fallecidos, encuestas, formacion, canales digitales...), GUARDE y CIERRE Excel.
echo Si quiere revisar diferencias, abra Revision\Auditoria_Extraccion_PE.xlsx (hoja Comparacion).
start "" "2_Tablero_Indicadores_PE.xlsx"
echo.
echo Cuando haya cerrado el Excel, presione una tecla para generar y publicar el dashboard.
pause >nul
if exist "~$2_Tablero_Indicadores_PE.xlsx" (
    echo El tablero sigue abierto: cierrelo y presione una tecla.
    pause >nul
)

echo.
echo === 3/3  Generando y publicando el dashboard (puede tardar mas de 10 minutos) ===
python 3_Generar_dashboard.py
if errorlevel 1 goto error

echo.
echo Listo: dashboard publicado.
pause
exit /b 0

:error
echo.
echo Hubo un error: revise el mensaje de arriba.
pause
exit /b 1
