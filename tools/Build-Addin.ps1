<#
.SYNOPSIS
  Build ExcelLLMAddin.xlam from the text source modules (source of truth).

.DESCRIPTION
  Creates a fresh workbook, imports every VBA module from source, and saves it as
  a .xlam add-in. This makes the binary a build artifact rather than a
  hand-edited file that drifts from the .bas/.cls sources. Requires Desktop Excel
  with "Trust access to the VBA project object model" enabled (see Run-Tests.ps1).
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path "$PSScriptRoot\.."),
    [string]$Output   = "$PSScriptRoot\..\ExcelLLMAddin.xlam"
)

$ErrorActionPreference = "Stop"

# Import order: vendored parser + its Dictionary first, then first-party.
$modules = @(
    "vendor\Dictionary.cls",
    "vendor\JsonConverter.bas",
    "modText.bas",
    "IHttpClient.cls",
    "MockHttpClient.cls",
    "WinHttpClient.cls",
    "CurlClient.cls",
    "modHttp.bas",
    "modConfig.bas",
    "modLLMFunctions.bas",
    "modTasks.bas",
    "modMenu.bas",
    "modTests.bas"
)

$xlOpenXMLAddIn = 55

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Add()
    foreach ($m in $modules) {
        $path = Join-Path $RepoRoot $m
        if (-not (Test-Path $path)) { throw "Missing module: $path" }
        Write-Host "Importing $m"
        [void]$wb.VBProject.VBComponents.Import($path)
    }

    $outFull = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Output))
    if (Test-Path $outFull) { Remove-Item $outFull -Force }
    $wb.SaveAs($outFull, $xlOpenXMLAddIn)
    Write-Host "Built add-in -> $outFull"
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
