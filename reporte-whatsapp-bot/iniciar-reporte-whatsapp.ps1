param([switch]$Restart)

$ErrorActionPreference = 'Continue'
$bot = $PSScriptRoot
$log = Join-Path $bot 'logs\reporte-whatsapp.log'

if ($Restart) {
    $scriptPath = [regex]::Escape($PSCommandPath)
    $existingSupervisors = Get-CimInstance Win32_Process -Filter "name = 'powershell.exe'" |
        Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -match $scriptPath }

    foreach ($process in $existingSupervisors) {
        & "$env:SystemRoot\System32\taskkill.exe" /PID $process.ProcessId /T /F | Out-Null
    }
}

$mutex = [Threading.Mutex]::new($false, 'Local\ReporteWhatsAppBot')
$hasMutex = $false

try {
    try {
        $hasMutex = $mutex.WaitOne(5000)
    } catch [Threading.AbandonedMutexException] {
        $hasMutex = $true
    }

    if (-not $hasMutex) {
        exit 0
    }

    while ($true) {
        try {
            $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Add-Content -Path $log -Value "[$stamp] Iniciando bot..."
            Set-Location -LiteralPath $bot
            & node .\dist\index.js 2>&1 | ForEach-Object { $_.ToString() } | Tee-Object -FilePath $log -Append
            $exitCode = $LASTEXITCODE

            if ($exitCode -eq 0) {
                $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Add-Content -Path $log -Value "[ $stamp ] El bot finalizo correctamente. No se reiniciara."
                break
            }

            $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Add-Content -Path $log -Value "[ $stamp ] El bot se cerro con error. Reiniciando en 10 segundos..."
            Start-Sleep -Seconds 10
        } catch {
            $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Add-Content -Path $log -Value "[ $stamp ] Error fatal: $($_.Exception.Message)"
            Start-Sleep -Seconds 10
        }
    }
} finally {
    if ($hasMutex) {
        $mutex.ReleaseMutex()
    }
    $mutex.Dispose()
}