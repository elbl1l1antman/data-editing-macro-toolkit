param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "MacroToolkit.bas not found: $modulePath"
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
    "Phone_NormalizeForTest",
    "Phone_ClassifyForTest",
    "NormalizeKoreanPhone",
    "Phone_ClassifyOne",
    "Phone_IsPlaceholderZeroPattern",
    "Phone_HasRepeatedDigitRun",
    "Phone_DigitsForKoreanRules",
    "Phone_IsSafeNumberDigits",
    "DigitsOnly"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function PhoneRulesPassed() As String
    Dim resultText As String
    resultText = CheckCase(resultText, "normalize 8-digit representative number", Phone_NormalizeForTest("15881222"), "1588-1222")
    resultText = CheckCase(resultText, "classify 000-0000 as error", Phone_ClassifyForTest("000-0000"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "classify 0000-0000 as error", Phone_ClassifyForTest("0000-0000"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "flag four repeated digits for review", Phone_ClassifyForTest("02-231-4444"), KoSeoulArea() & "|" & KoReview())
    resultText = CheckCase(resultText, "keep ordinary 8-digit representative number normal", Phone_ClassifyForTest("1588-1222"), KoRepresentative() & "|" & KoNormal())

    If Len(resultText) = 0 Then resultText = "PASS"
    PhoneRulesPassed = resultText
End Function

Private Function CheckCase(ByVal resultText As String, ByVal caseName As String, ByVal actual As String, ByVal expected As String) As String
    If actual <> expected Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected " & expected & ", got " & actual
    End If
    CheckCase = resultText
End Function

Private Function KoErrorPattern() As String
    KoErrorPattern = ChrW(&HC624) & ChrW(&HB958) & " " & ChrW(&HD328) & ChrW(&HD134)
End Function

Private Function KoError() As String
    KoError = ChrW(&HC624) & ChrW(&HB958)
End Function

Private Function KoSeoulArea() As String
    KoSeoulArea = ChrW(&HC11C) & ChrW(&HC6B8) & " " & ChrW(&HC9C0) & ChrW(&HC5ED) & ChrW(&HBC88) & ChrW(&HD638)
End Function

Private Function KoReview() As String
    KoReview = ChrW(&HD655) & ChrW(&HC778) & " " & ChrW(&HD544) & ChrW(&HC694)
End Function

Private Function KoRepresentative() As String
    KoRepresentative = ChrW(&HB300) & ChrW(&HD45C) & ChrW(&HBC88) & ChrW(&HD638)
End Function

Private Function KoNormal() As String
    KoNormal = ChrW(&HC815) & ChrW(&HC0C1)
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

    $result = $excel.Run("PhoneRulesPassed")
    if ($result -ne "PASS") {
        throw $result
    }

    Write-Host "Phone rule tests passed."
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
