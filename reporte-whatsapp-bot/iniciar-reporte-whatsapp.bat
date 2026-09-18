@echo off
setlocal EnableExtensions

set "BOT_PATH=%~dp0"
cd /d "%BOT_PATH%"

where node >nul 2>&1
if errorlevel 1 (
  echo ERROR: Node.js no esta disponible en PATH.
  pause
  exit /b 1
)

if not exist ".env" (
  echo ERROR: no existe el archivo .env.
  echo Copia .env.example como .env y configura el grupo de WhatsApp.
  pause
  exit /b 1
)

if not exist "dist\index.js" (
  echo No existe la compilacion. Construyendo el bot...
  call npm run build
  if errorlevel 1 (
    echo ERROR: no se pudo construir el bot.
    pause
    exit /b 1
  )
)

echo Bot de WhatsApp iniciado.
echo Esta ventana debe permanecer abierta.
echo Presiona Ctrl+C para detenerlo.
echo.

:RESTART
npm start
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="0" (
  echo El bot se detuvo correctamente.
  pause
  exit /b 0
)

echo El bot se cerro con codigo %EXIT_CODE%.
echo Reiniciando en 10 segundos...
timeout /t 10 /nobreak >nul
goto RESTART
