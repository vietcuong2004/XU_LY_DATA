param([Parameter(Mandatory=$true)][string]$RequestPath)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$requestData = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$excelApp = $null
$workbook = $null
$existingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ExcelWindowProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@
try {
    $excelApp = New-Object -ComObject Excel.Application
    [uint32]$ownedExcelId = 0
    [void][ExcelWindowProcess]::GetWindowThreadProcessId([IntPtr]$excelApp.Hwnd, [ref]$ownedExcelId)
    if ($ownedExcelId -and $existingExcel -notcontains $ownedExcelId) {
        $ownedProcess = Get-Process -Id $ownedExcelId
        @{ id = $ownedExcelId; started = $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString() } |
            ConvertTo-Json | Set-Content -LiteralPath ($RequestPath + '.process.json') -Encoding UTF8
    }
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.EnableEvents = $false
    $excelApp.AutomationSecurity = 3
    $excelApp.AskToUpdateLinks = $false
    $workbook = $excelApp.Workbooks.Open($requestData.path, 0, $false)
    if ($workbook.ReadOnly) { throw 'Workbook opened read-only.' }
    foreach ($edit in $requestData.edits) {
        $sheet = $workbook.Worksheets.Item([string]$edit.sheet)
        $cell = $sheet.Range([string]$edit.cell)
        try {
            if ($cell.HasArray) { throw 'Cannot edit part of an array formula.' }
            if ($edit.kind -eq 'formula') {
                $cell.Formula = [string]$edit.value
            } elseif ($edit.kind -eq 'blank') {
                $cell.ClearContents() | Out-Null
            } elseif ($edit.kind -eq 'number') {
                $cell.Value2 = [double]$edit.value
            } elseif ($edit.kind -eq 'boolean') {
                $cell.Value2 = [bool]$edit.value
            } elseif ($edit.kind -eq 'date') {
                $cell.Value2 = [datetime]::Parse([string]$edit.value, [cultureinfo]::InvariantCulture).ToOADate()
            } else {
                # Prefix with an apostrophe so text beginning '=' stays text.
                $cell.Value2 = "'" + [string]$edit.value
            }
        } finally {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
        }
    }
    $excelApp.CalculateFullRebuild()
    if ($excelApp.CalculationState -ne 0) { throw 'Workbook calculation did not finish.' }
    $workbook.Save()
    Write-Output 'CALCULATED'
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }
    if ($null -ne $excelApp) {
        $excelApp.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
