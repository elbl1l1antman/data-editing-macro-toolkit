param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$publishPath = Join-Path $ProjectRoot "scripts\publish_release.ps1"
$source = Get-Content -LiteralPath $modulePath -Raw -Encoding UTF8
$publishSource = Get-Content -LiteralPath $publishPath -Raw -Encoding UTF8

foreach ($token in @(
    "DataCleanup_ClearResidualBlankValues",
    "Ribbon_DataCleanupClearResidualBlanks",
    "9. 수식 빈값 정리",
    'label="수식 빈값 정리"'
)) {
    if (-not ($source.Contains($token) -or $publishSource.Contains($token))) {
        throw "Residual blank cleanup token missing: $token"
    }
}

function Get-VbaBlock {
    param([string]$SourceText, [string]$ProcedureName)
    $start = [regex]::Match($SourceText, "(?m)^(Public|Private)\s+(Function|Sub)\s+$([regex]::Escape($ProcedureName))\b.*$")
    if (-not $start.Success) { throw "Procedure not found: $ProcedureName" }
    $kind = $start.Groups[2].Value
    $endToken = if ($kind -eq "Function") { "End Function" } else { "End Sub" }
    $tail = $SourceText.Substring($start.Index)
    $end = [regex]::Match($tail, "(?m)^$endToken\s*$")
    if (-not $end.Success) { throw "Procedure end not found: $ProcedureName" }
    $tail.Substring(0, $end.Index + $end.Length)
}

$moduleCode = "Option Explicit`r`n" + @(
    Get-VbaBlock $source "DataCleanup_IsResidualBlankForTest"
    Get-VbaBlock $source "DataCleanup_IsResidualBlankValue"
) -join "`r`n`r`n"

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($moduleCode)

    if (-not [bool]$excel.Run("DataCleanup_IsResidualBlankForTest", "", $false)) { throw "zero-length value must be cleared" }
    if ([bool]$excel.Run("DataCleanup_IsResidualBlankForTest", "", $true)) { throw "formula result must be preserved" }
    if ([bool]$excel.Run("DataCleanup_IsResidualBlankForTest", " ", $false)) { throw "space-only value must be preserved" }
    if ([bool]$excel.Run("DataCleanup_IsResidualBlankForTest", 0, $false)) { throw "numeric zero must be preserved" }
    if ([bool]$excel.Run("DataCleanup_IsResidualBlankForTest", "text", $false)) { throw "ordinary text must be preserved" }

    Write-Host "Residual blank cleanup tests passed."
}
finally {
    if ($workbook -ne $null) { $workbook.Close($false) | Out-Null }
    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
