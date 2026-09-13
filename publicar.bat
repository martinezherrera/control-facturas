@echo off
REM ====================================================================
REM  Control de facturas - regenerar y publicar
REM
REM  Este .bat vive DENTRO del repositorio, junto al Excel.
REM  Misma estructura en el PC de la casa y en el del trabajo.
REM
REM  1. Regenera dashboard_facturas.html  (local, razones sociales reales)
REM  2. Regenera index.html + datos.js    (web, proveedores anonimizados)
REM  3. Commit y push -> GitHub Pages actualiza el enlace
REM ====================================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"

set EXCEL=control de ingresos facturas.xlsx

echo.
echo ============================================
echo  CONTROL DE FACTURAS - publicar
echo ============================================

REM --- comprobaciones previas ---
if not exist "%EXCEL%" (
  echo.
  echo ERROR: no encuentro "%EXCEL%" en esta carpeta.
  echo Copialo aqui antes de publicar.
  pause
  exit /b 1
)

if not exist "proveedores_map.csv" (
  echo.
  echo ATENCION: no esta proveedores_map.csv.
  echo Sin ese archivo los codigos PROV-nn se reasignan desde cero
  echo y dejan de coincidir con los ya publicados.
  echo.
  set /p SEGUIR="Escribe SI para continuar igual: "
  if /i not "!SEGUIR!"=="SI" exit /b 1
)

echo.
echo [1/3] Generando dashboards...
python generar_dashboard_html.py "%EXCEL%" "dashboard_facturas.html" --web .
if errorlevel 1 (
  echo.
  echo ERROR al generar. No se publico nada.
  pause
  exit /b 1
)

echo.
echo [2/3] Revisando que no se cuele nada sensible...
git status --porcelain | findstr /i /c:".xlsx" /c:"proveedores_map" /c:"dashboard_facturas.html" >nul
if not errorlevel 1 (
  echo.
  echo ABORTADO: hay archivos sensibles sin ignorar.
  git status --short
  echo.
  echo Revisa el .gitignore antes de continuar.
  pause
  exit /b 1
)
echo      OK, solo archivos publicables.

echo.
echo [3/3] Publicando...
git add index.html datos.js README.md generar_dashboard_html.py publicar.bat .gitignore
git diff --cached --quiet
if not errorlevel 1 (
  echo      No hay cambios que publicar.
  pause
  exit /b 0
)

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set HOY=%%a-%%b-%%c
git commit -m "Actualizacion de datos %HOY%"
git push
if errorlevel 1 (
  echo.
  echo ERROR en el push. Revisa conexion o credenciales de GitHub.
  pause
  exit /b 1
)

echo.
echo ============================================
echo  Listo. El sitio se actualiza en 30s a 2min:
echo  https://martinezherrera.github.io/control-facturas/
echo ============================================
pause
