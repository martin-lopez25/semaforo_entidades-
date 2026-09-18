@echo off
setlocal EnableExtensions

set "REPO_PATH=C:\Users\jose.valdez\Downloads\reporte_pao\semaforo_entidades-"
set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python313\python.exe"
set "BRANCH=main"

echo Actualizando reporte desde: %REPO_PATH%
cd /d "%REPO_PATH%"
if errorlevel 1 (
  echo ERROR: no se encontro el repositorio.
  pause
  exit /b 1
)

for /d %%D in ("%REPO_PATH%\reporte-de-inventario-de-unidades-*") do set "APP_PATH=%%~fD"
if not defined APP_PATH (
  echo ERROR: no se encontro la carpeta del frontend.
  pause
  exit /b 1
)

if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"
where npm >nul 2>&1
if errorlevel 1 (
  echo ERROR: npm no esta disponible en PATH.
  pause
  exit /b 1
)

echo Ejecutando reporte_semaforo.py...
"%PYTHON_EXE%" reporte_semaforo.py
if errorlevel 1 (
  echo ERROR: fallo el script de Python.
  pause
  exit /b 1
)

echo Construyendo el nuevo frontend...
cd /d "%APP_PATH%"
if errorlevel 1 (
  echo ERROR: no se pudo entrar a la carpeta frontend.
  pause
  exit /b 1
)

for /f "delims=" %%A in ('powershell -NoProfile -Command "(Get-Content '%APP_PATH%\public\reporte-inventario-data.json' -Raw ^| ConvertFrom-Json).lastUpdated"') do set "VITE_REPORT_LAST_UPDATED=%%A"
if not defined VITE_REPORT_LAST_UPDATED (
  echo ERROR: no se pudo obtener la fecha de actualizacion.
  pause
  exit /b 1
)
echo Fecha fija de la build: %VITE_REPORT_LAST_UPDATED%

call npm run build
if errorlevel 1 (
  echo ERROR: fallo la construccion del frontend.
  pause
  exit /b 1
)

if not exist "%APP_PATH%\dist\index.html" (
  echo ERROR: no se genero dist\index.html.
  pause
  exit /b 1
)

echo Publicando la build nueva en la raiz...
cd /d "%REPO_PATH%"
if exist "index.html" del /q "index.html"
if exist "assets" rmdir /s /q "assets"
xcopy /e /i /y "%APP_PATH%\dist\*" "%REPO_PATH%\" >nul
if errorlevel 1 (
  echo ERROR: no se pudo copiar la build.
  pause
  exit /b 1
)

findstr /c:"./assets/" "index.html" >nul
if errorlevel 1 (
  echo ERROR: index.html no contiene las rutas esperadas de assets.
  pause
  exit /b 1
)

echo Agregando cambios a Git...
git checkout %BRANCH% 2>nul || git checkout -b %BRANCH%
git add -A

git diff --cached --quiet
if errorlevel 1 (
  echo Creando commit...
  git commit -m "actualizacion automatica %DATE%"
  if errorlevel 1 (
    echo ERROR: fallo el commit.
    pause
    exit /b 1
  )
) else (
  echo No hay cambios nuevos para commitear.
)

echo Subiendo cambios a GitHub...
git push -u origin %BRANCH%
if errorlevel 1 (
  echo ERROR: fallo el push a GitHub.
  pause
  exit /b 1
)

echo Actualizacion completada correctamente.
pause
