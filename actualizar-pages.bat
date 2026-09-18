@echo off
setlocal

set "REPO_PATH=C:\Users\jose.valdez\Downloads\reporte_pao\semaforo_entidades-"
set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python313\python.exe"
set "BRANCH=main"

echo Actualizando reporte desde: %REPO_PATH%
cd /d "%REPO_PATH%"
if errorlevel 1 exit /b 1

for /d %%D in ("%REPO_PATH%\reporte-de-inventario-de-unidades-*") do set "APP_PATH=%%~fD"
if not defined APP_PATH (
  echo ERROR: no se encontro la carpeta del frontend.
  pause
  exit /b 1
)

if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

echo Ejecutando reporte_semaforo.py...
"%PYTHON_EXE%" reporte_semaforo.py
if errorlevel 1 (
  echo ERROR: no se pudo generar el reporte con Python.
  pause
  exit /b 1
)

echo Construyendo el frontend nuevo...
cd /d "%APP_PATH%"
if errorlevel 1 (
  echo ERROR: no se pudo entrar a la carpeta del frontend: %APP_PATH%
  pause
  exit /b 1
)
for /f "delims=" %%A in ('powershell -NoProfile -Command "(Get-Content '%APP_PATH%\public\reporte-inventario-data.json' -Raw | ConvertFrom-Json).lastUpdated"') do set "VITE_REPORT_LAST_UPDATED=%%A"
if not defined VITE_REPORT_LAST_UPDATED (
  echo ERROR: no se pudo obtener la fecha de actualizacion.
  pause
  exit /b 1
)
echo Fecha fija de la build: %VITE_REPORT_LAST_UPDATED%
call npm run build
if errorlevel 1 (
  echo ERROR: no se pudo construir el frontend.
  pause
  exit /b 1
)

echo Publicando la build nueva en la raiz del repositorio...
cd /d "%REPO_PATH%"
if exist "index.html" del /q "index.html"
xcopy /e /i /y "%APP_PATH%dist\*" "%REPO_PATH%" >nul
if errorlevel 1 (
  echo ERROR: no se pudo copiar la build a la raiz.
  pause
  exit /b 1
)

git checkout %BRANCH% 2>nul || git checkout -b %BRANCH%
git add -A
git commit -m "actualizacion automatica %DATE%"
git push -u origin %BRANCH%
if errorlevel 1 (
  echo ERROR: la actualizacion local fue creada, pero el push fallo.
  pause
  exit /b 1
)

echo Actualizacion completada. GitHub Pages puede tardar unos minutos.
pause