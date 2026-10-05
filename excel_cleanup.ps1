param([Parameter(Mandatory=$true)][string]$RecordPath)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $RecordPath)) { exit 0 }
$record = Get-Content -LiteralPath $RecordPath -Raw -Encoding UTF8 | ConvertFrom-Json
$ownedProcess = Get-Process -Id ([int]$record.id) -ErrorAction SilentlyContinue
if ($null -ne $ownedProcess -and $ownedProcess.ProcessName -eq 'EXCEL' -and
    $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString() -eq [string]$record.started) {
    Stop-Process -Id $ownedProcess.Id -Force
}
