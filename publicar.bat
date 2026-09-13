@echo off
REM ====================================================================
REM  Control de facturas - regenerar y publicar
REM
REM  Genera index.html + datos.js con los nombres reales de proveedor
REM  y los publica en GitHub Pages, que es lo que alcanza el servidor
REM  de las pantallas:
REM      https://martinezherrera.github.io/control-facturas/
REM
REM  Antes de generar hace fetch + merge, para que trabajar desde dos
REM  PC distintos no termine en un push rechazado.
REM
REM  Uso manual:      doble clic
REM  Uso programado:  publicar.bat /auto     (no espera tecla)
REM ====================================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"

set EXCEL=control de ingresos facturas.xlsx

set AUTO=0
if /i "%~1"=="/auto" set AUTO=1

echo.
echo ============================================
echo  CONTROL DE FACTURAS   %date% %time:~0,5%
echo ============================================

if not exist "%EXCEL%" (
  echo ERROR: no encuentro "%EXCEL%" en esta carpeta.
  if "%AUTO%"=="0" pause
  exit /b 1
)

echo.
echo [1/4] Trayendo cambios del otro PC...
git fetch origin >nul 2>&1
git merge origin/main --no-edit >nul 2>&1
if errorlevel 1 (
  REM index.html y datos.js son generados: el conflicto no importa,
  REM se sobrescriben completos en el paso siguiente.
  git checkout --ours index.html datos.js >nul 2>&1
  git add index.html datos.js >nul 2>&1
  git commit --no-edit >nul 2>&1
)
echo      OK.

echo.
echo [2/4] Generando dashboard...
python generar_dashboard_html.py "%EXCEL%" --web .
if errorlevel 1 (
  echo ERROR al generar. No se publico nada.
  if "%AUTO%"=="0" pause
  exit /b 1
)

echo.
echo [3/4] Revisando que no se cuele el Excel...
git status --porcelain | findstr /i /c:".xlsx" /c:".xlsm" >nul
if not errorlevel 1 (
  echo ABORTADO: el archivo fuente aparece sin ignorar.
  git status --short
  if "%AUTO%"=="0" pause
  exit /b 1
)
echo      OK.

echo.
echo [4/4] Publicando...
git add index.html datos.js README.md generar_dashboard_html.py publicar.bat .gitignore
git diff --cached --quiet
if not errorlevel 1 (
  echo      Sin cambios que publicar.
  if "%AUTO%"=="0" pause
  exit /b 0
)

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set HOY=%%a-%%b-%%c
git commit -m "Actualizacion de datos %HOY%"
git push
if errorlevel 1 (
  echo ERROR en el push. Revisa conexion o credenciales.
  if "%AUTO%"=="0" pause
  exit /b 1
)

echo.
echo  Publicado. La pantalla lo toma en pocos minutos:
echo  https://martinezherrera.github.io/control-facturas/
echo.
if "%AUTO%"=="0" pause
exit /b 0
