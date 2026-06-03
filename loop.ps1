$LOG  = "C:\Users\MICRO1\Documents\icc-inadimplencia\execucao.log"
$PS   = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$SCPT = "C:\Users\MICRO1\Documents\icc-inadimplencia\atualizar.ps1"

while ($true) {
    $inicio = Get-Date -f "HH:mm:ss"
    "[${inicio}] Iniciando ciclo" | Add-Content $LOG
    
    $proc = Start-Process $PS -ArgumentList "-ExecutionPolicy Bypass -NonInteractive -File `"$SCPT`"" -PassThru -WindowStyle Hidden -RedirectStandardOutput "$LOG.tmp" 2>$null
    
    $concluiu = $proc.WaitForExit(180000)  # 3 minutos
    
    if (-not $concluiu) {
        $proc.Kill()
        "[$(Get-Date -f 'HH:mm:ss')] TIMEOUT - processo encerrado" | Add-Content $LOG
    } else {
        if (Test-Path "$LOG.tmp") { Get-Content "$LOG.tmp" | Add-Content $LOG; Remove-Item "$LOG.tmp" -Force }
        "[$(Get-Date -f 'HH:mm:ss')] Ciclo concluido" | Add-Content $LOG
    }
    
    Start-Sleep -Seconds 420
}