param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$publishScriptPath = Join-Path $ProjectRoot "scripts\publish_release.ps1"
$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
if (-not (Test-Path -LiteralPath $publishScriptPath)) { throw "publish_release.ps1 not found: $publishScriptPath" }
if (-not (Test-Path -LiteralPath $modulePath)) { throw "MacroToolkit.bas not found: $modulePath" }

$publishScript = Get-Content -LiteralPath $publishScriptPath -Raw -Encoding UTF8
$formStart = $publishScript.IndexOf("function Add-NaverLocalVerifyOptionsUserForm")
$formEnd = $publishScript.IndexOf("function Add-AdminDongMasterWorksheet", $formStart)
if ($formStart -lt 0 -or $formEnd -lt 0) {
    throw "Could not isolate Add-NaverLocalVerifyOptionsUserForm block."
}

$formBlock = $publishScript.Substring($formStart, $formEnd - $formStart)
if ($formBlock -match "Application\.InputBox") {
    throw "Naver verify UserForm must not call Application.InputBox directly; range picking must happen after the modal form closes."
}

$source = Get-Content -LiteralPath $modulePath -Raw -Encoding UTF8

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
    "NaverLocal_VerifyPromptListForTest",
    "NaverLocal_VerifyPromptList"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function NaverVerifyRangePickerRulesPassed() As String
    Dim resultText As String
    resultText = CheckCase(resultText, "both blank ranges should prompt both", NaverLocal_VerifyPromptListForTest(True, True, "", ""), "name,address")
    resultText = CheckCase(resultText, "typed name range should prompt address only", NaverLocal_VerifyPromptListForTest(True, True, "Sheet1!A:A", ""), "address")
    resultText = CheckCase(resultText, "address-only mode should prompt address", NaverLocal_VerifyPromptListForTest(False, True, "", ""), "address")
    resultText = CheckCase(resultText, "name-only mode with typed range should not prompt", NaverLocal_VerifyPromptListForTest(True, False, "Sheet1!B:B", ""), "")

    If Len(resultText) = 0 Then resultText = "PASS"
    NaverVerifyRangePickerRulesPassed = resultText
End Function

Private Function CheckCase(ByVal resultText As String, ByVal caseName As String, ByVal actual As String, ByVal expected As String) As String
    If actual <> expected Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected " & expected & ", got " & actual
    End If
    CheckCase = resultText
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

    $result = $excel.Run("NaverVerifyRangePickerRulesPassed")
    if ($result -ne "PASS") {
        throw $result
    }

    Write-Host "Naver verify range picker tests passed."
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
