<#
.SYNOPSIS
  Headless test runner for the ExcelLLM add-in (Windows CI).

.DESCRIPTION
  Opens Excel via COM, imports every VBA source module (first-party + vendored),
  invokes the in-workbook RunAllTests harness with no UI, copies out the JUnit
  report, and exits non-zero if any test failed. Requires Desktop Excel.

  Because the harness injects MockHttpClient, no network or provider is needed.

.NOTES
  Enables "Trust access to the VBA project object model" for the current user,
  which is required to import modules programmatically.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path "$PSScriptRoot\.."),
    [string]$OutputXml = "$PSScriptRoot\..\test-results.xml"
)

$ErrorActionPreference = "Stop"

function Enable-VBOMAccess {
    # Find installed Office version keys and allow VBA project model access.
    Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        ForEach-Object {
            $sec = "HKCU:\Software\Microsoft\Office\$($_.PSChildName)\Excel\Security"
            New-Item -Path $sec -Force | Out-Null
            Set-ItemProperty -Path $sec -Name "AccessVBOM"   -Value 1 -Type DWord
            Set-ItemProperty -Path $sec -Name "VBAWarnings"  -Value 1 -Type DWord
        }
}

# Source modules to import, in dependency-friendly order. Vendored parser first.
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
    "modAgent.bas",
    "modMcp.bas",
    "modMenu.bas",
    "modTests.bas"
)

Enable-VBOMAccess

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable (no auto macros)

$failCount = 1
try {
    $wb = $excel.Workbooks.Add()

    foreach ($m in $modules) {
        $path = Join-Path $RepoRoot $m
        if (-not (Test-Path $path)) { throw "Missing module: $path" }
        Write-Host "Importing $m"
        [void]$wb.VBProject.VBComponents.Import($path)
    }

    Write-Host "Running RunAllTests..."
    # RunAllTests(showUI:=False) returns the failure count.
    $failCount = [int]$excel.Run("RunAllTests", $false)

    # Copy the JUnit report the harness wrote to the temp dir.
    $report = Join-Path $env:TEMP "excelllm_junit.xml"
    if (Test-Path $report) {
        Copy-Item $report $OutputXml -Force
        Write-Host "JUnit report -> $OutputXml"
    } else {
        Write-Warning "No JUnit report produced at $report"
    }

    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($failCount -ne 0) {
    Write-Error "$failCount test(s) failed."
    exit 1
}
Write-Host "All tests passed."
exit 0
