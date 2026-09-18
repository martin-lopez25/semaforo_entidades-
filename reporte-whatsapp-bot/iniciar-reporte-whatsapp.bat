@echo off
setlocal EnableExtensions

set "BOT_PATH=%~dp0"
set "LOG_PATH=%BOT_PATH%logs\reporte-whatsapp.log"
mkdir "%BOT_PATH%logs" 2>nul
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

echo Iniciando el bot en segundo plano...
echo Log: %LOG_PATH%
echo.

powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "& { $ErrorActionPreference = 'Stop'; $log = '%LOG_PATH%'; $bot = '%BOT_PATH%'; while ($true) { try { $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Add-Content -Path $log -Value (\"[$stamp] Iniciando bot...\"); Set-Location $bot; npm start 2>&1 | Tee-Object -FilePath $log -Append; Add-Content -Path $log -Value ('[ ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + ' ] El bot se cerró. Reiniciando en 10 segundos...'); Start-Sleep -Seconds 10; } catch { Add-Content -Path $log -Value ('[ ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + ' ] Error fatal: ' + $_.Exception.Message); Start-Sleep -Seconds 10; } } }"

echo El bot fue lanzado en segundo plano.
echo Si quieres detenerlo, cierra el proceso "node.exe" o usa:
echo   taskkill /F /IM node.exe
exit /b 0
