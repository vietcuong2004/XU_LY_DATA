$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not $env:SHIPMENT_API_TOKEN -or $env:SHIPMENT_API_TOKEN.Length -lt 32) {
    throw 'Set SHIPMENT_API_TOKEN to a random secret of at least 32 characters. See WINDOWS_BACKEND_SETUP.md.'
}
if ($env:VERCEL -or $env:SHIPMENT_BACKEND_URL) {
    throw 'Do not set VERCEL or SHIPMENT_BACKEND_URL on the Windows backend.'
}
python ui_server.py --port 8765
