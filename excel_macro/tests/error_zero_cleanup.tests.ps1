param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"
$source = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot
$publish = Get-Content -LiteralPath (Join-Path $ProjectRoot "scripts\publish_release.ps1") -Raw -Encoding UTF8

foreach ($token in @(
    "DataCleanup_ClearErrorAndZeroValues",
    "Ribbon_DataCleanupClearErrorsAndZeros",
    "10. 오류값/0 지우기",
    'label="오류값/0 지우기"'
)) {
    if (-not ($source.Contains($token) -or $publish.Contains($token))) {
        throw "Error/zero cleanup token missing: $token"
    }
}

function Get-VbaBlock([string]$ProcedureName) {
    $start = [regex]::Match($source, "(?m)^(Public|Private)\s+(Function|Sub)\s+$([regex]::Escape($ProcedureName))\b.*$")
    if (-not $start.Success) { throw "Procedure not found: $ProcedureName" }
    $kind = $start.Groups[2].Value
    $tail = $source.Substring($start.Index)
    $finish = [regex]::Match($tail, "(?m)^End $kind\s*$")
    if (-not $finish.Success) { throw "Procedure end not found: $ProcedureName" }
    $tail.Substring(0, $finish.Index + $finish.Length)
}

$wrapper = @'
Public Function TestCleanupError(ByVal errorCode As Long) As Long
    TestCleanupError = DataCleanup_ClearValueKindForTest(CVErr(errorCode))
End Function
'@
$moduleCode = "Option Explicit`r`n" + @(
    Get-VbaBlock "DataCleanup_ClearValueKindForTest"
    Get-VbaBlock "DataCleanup_ClearValueKind"
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
    $module.CodeModule.AddFromString($moduleCode + "`r`n`r`n" + $wrapper)

    if ([int]$excel.Run("DataCleanup_ClearValueKindForTest", 0) -ne 2) { throw "numeric zero must be cleared" }
    if ([int]$excel.Run("DataCleanup_ClearValueKindForTest", "0") -ne 0) { throw "text zero must be preserved" }
    if ([int]$excel.Run("DataCleanup_ClearValueKindForTest", 1) -ne 0) { throw "non-zero number must be preserved" }
    if ([int]$excel.Run("DataCleanup_ClearValueKindForTest", "") -ne 0) { throw "blank string must be preserved" }
    foreach ($code in @(2007, 2015, 2023, 2042)) {
        if ([int]$excel.Run("TestCleanupError", $code) -ne 1) { throw "Excel error $code must be cleared" }
    }

    Write-Host "Error and zero cleanup tests passed."
}
finally {
    if ($workbook -ne $null) { $workbook.Close($false) | Out-Null }
    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
