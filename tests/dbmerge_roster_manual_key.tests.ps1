param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"

$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "MacroToolkit.bas not found: $modulePath"
}

$source = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot

function Get-VbaBlock {
    param(
        [string]$SourceText,
        [string]$ProcedureName
    )

    $startMatch = [regex]::Match(
        $SourceText,
        "(?m)^(Public|Private)\s+(Function|Sub)\s+$([regex]::Escape($ProcedureName))\b.*$"
    )
    if (-not $startMatch.Success) {
        throw "Procedure not found: $ProcedureName"
    }

    $procedureKind = $startMatch.Groups[2].Value
    $endToken = if ($procedureKind -eq "Function") { "End Function" } else { "End Sub" }
    $tail = $SourceText.Substring($startMatch.Index)
    $endMatch = [regex]::Match($tail, "(?m)^$endToken\s*$")
    if (-not $endMatch.Success) {
        throw "Procedure end not found: $ProcedureName"
    }

    return $tail.Substring(0, $endMatch.Index + $endMatch.Length)
}

$procedureNames = @(
    "DBMerge_ManualKeyColumnForTest",
    "DBMerge_ApplyManualKeyColumnToMap",
    "DBMerge_RequestManualKeyColumns",
    "DBMerge_NormalizeHeader",
    "Roster_ManualKeyRelativeColumnForTest",
    "Roster_RelativeColumnWithinRange",
    "Roster_ApplyManualIdColumn"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function ManualKeyRulesPassed() As String
    Dim resultText As String
    resultText = CheckLong(resultText, "DBMerge manual key overrides existing header", DBMerge_ManualKeyColumnForTest("ID", 2, 5), 5)
    resultText = CheckLong(resultText, "DBMerge manual key fills missing header", DBMerge_ManualKeyColumnForTest("ID", 0, 3), 3)
    resultText = CheckLong(resultText, "Roster selected key column becomes relative column", Roster_ManualKeyRelativeColumnForTest(2, 5, 4), 3)
    resultText = CheckLong(resultText, "Roster selected key column outside range is rejected", Roster_ManualKeyRelativeColumnForTest(2, 5, 8), 0)

    If Len(resultText) = 0 Then resultText = "PASS"
    ManualKeyRulesPassed = resultText
End Function

Private Function CheckLong(ByVal resultText As String, ByVal caseName As String, ByVal actual As Long, ByVal expected As Long) As String
    If actual <> expected Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected " & CStr(expected) & ", got " & CStr(actual)
    End If
    CheckLong = resultText
End Function
'@

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
    $wrapper = $workbook.VBProject.VBComponents.Add(1)
    $wrapper.CodeModule.AddFromString($wrapperCode)

    $result = $excel.Run("ManualKeyRulesPassed")
    if ($result -ne "PASS") {
        throw $result
    }

    Write-Host "Manual key selection rule tests passed."
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
