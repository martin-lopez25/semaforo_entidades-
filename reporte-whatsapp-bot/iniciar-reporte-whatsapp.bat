@echo off
setlocal EnableExtensions

set "BOT_PATH=C:\Users\jose.valdez\Downloads\reporte_pao\semaforo_entidades-\reporte-whatsapp-bot"
set "ENV_PATH=%BOT_PATH%\.env"
set "LOG_PATH=%BOT_PATH%\logs\reporte-whatsapp.log"
mkdir "%BOT_PATH%\logs" 2>nul
cd /d "%BOT_PATH%"

where node >nul 2>&1
if errorlevel 1 (
  echo ERROR: Node.js no esta disponible en PATH.
  pause
  exit /b 1
)

if not exist "%ENV_PATH%" (
  echo ERROR: no existe el archivo .env.
  echo Ruta esperada: %ENV_PATH%
  pause
  exit /b 1
)

echo Construyendo el bot con el codigo actual...
call npm run build
if errorlevel 1 (
  echo ERROR: no se pudo construir el bot.
  pause
  exit /b 1
)

echo Iniciando el bot en segundo plano...
echo Log: %LOG_PATH%
echo.

start "ReporteWhatsAppBot" /MIN powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "& { $ErrorActionPreference = 'Continue'; $log = '%LOG_PATH%'; $bot = '%BOT_PATH%'; Set-Location $bot; while ($true) { $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Add-Content -Path $log -Value ('[' + $stamp + '] Iniciando bot...'); & node dist/index.js 2>&1 | ForEach-Object { $_.ToString() } | Add-Content -Path $log; $code = $LASTEXITCODE; $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); if ($code -eq 0) { Add-Content -Path $log -Value ('[' + $stamp + '] Apagado programado (codigo 0). No se reinicia.'); break }; Add-Content -Path $log -Value ('[' + $stamp + '] El bot se cerro con codigo ' + $code + '. Reiniciando en 10 segundos...'); Start-Sleep -Seconds 10 } }"

echo El bot fue lanzado en segundo plano.
echo Si quieres detenerlo, cierra el proceso "node.exe" o usa:
echo   taskkill /F /IM node.exe
exit /b 0
