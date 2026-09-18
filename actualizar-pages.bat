@echo off
setlocal

set "REPO_PATH=%~dp0"
set "APP_PATH=%REPO_PATH%reporte-de-inventario-de-unidades-médicas"
set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python313\python.exe"
set "BRANCH=main"

echo Actualizando reporte desde: %REPO_PATH%
cd /d "%REPO_PATH%"
if errorlevel 1 exit /b 1

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