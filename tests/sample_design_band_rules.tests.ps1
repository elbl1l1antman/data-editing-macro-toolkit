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
    "SampleDesign_ParseBandSpecs",
    "SampleDesign_BandTryNumber",
    "SampleDesign_StripBandNumberUnitSuffix",
    "SampleDesign_NormalizeBandStartOp",
    "SampleDesign_NormalizeBandEndOp",
    "SampleDesign_BandAutoLabel",
    "SampleDesign_BandFormulaForAddress",
    "SampleDesign_BandConditionFormula",
    "SampleDesign_BandValueFormula",
    "SampleDesign_FormulaNumberLiteral",
    "SampleDesign_FormulaStringLiteral"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function SampleDesignBandRulesPassed() As String
    Dim resultText As String
    Dim n As Double

    resultText = CheckBool(resultText, "plain numeric text parses", SampleDesign_BandTryNumber("1,000", n), True)
    resultText = CheckDouble(resultText, "plain numeric text value", n, 1000)

    resultText = CheckBool(resultText, "unit-suffixed Korean numeric text parses", SampleDesign_BandTryNumber("1,000" & KoPersonUnit(), n), True)
    resultText = CheckDouble(resultText, "unit-suffixed Korean numeric text value", n, 1000)

    resultText = CheckBool(resultText, "percent text parses", SampleDesign_BandTryNumber("75%", n), True)
    resultText = CheckDouble(resultText, "percent text value", n, 0.75)

    Dim bands As Collection
    Set bands = New Collection
    resultText = CheckBool(resultText, "valid band specs parse", SampleDesign_ParseBandSpecs("1" & vbTab & "" & vbTab & "이상" & vbTab & "5" & vbTab & "미만" & vbTab & "under5" & vbLf & "2" & vbTab & "5" & vbTab & "이상" & vbTab & "10" & vbTab & "미만" & vbTab & "5to9" & vbLf & "3" & vbTab & "10" & vbTab & "이상" & vbTab & "" & vbTab & "미만" & vbTab & "over10", bands), True)
    resultText = CheckContains(resultText, "formula uses IFS", SampleDesign_BandFormulaForAddress("Q4", bands), "IFS(")
    resultText = CheckContains(resultText, "formula references selected cell", SampleDesign_BandFormulaForAddress("Q4", bands), "Q4")
    resultText = CheckContains(resultText, "formula labels blank source", SampleDesign_BandFormulaForAddress("Q4", bands), ChrW(&HC790) & ChrW(&HB8CC) & " " & ChrW(&HC5C6) & ChrW(&HC74C))
    resultText = CheckContains(resultText, "formula keeps direct readable less-than condition", SampleDesign_BandFormulaForAddress("Q4", bands), "Q4<5")
    resultText = CheckContains(resultText, "formula preserves first label", SampleDesign_BandFormulaForAddress("Q4", bands), """under5""")
    resultText = CheckContains(resultText, "formula preserves open-ended last label", SampleDesign_BandFormulaForAddress("Q4", bands), """over10""")

    If Len(resultText) = 0 Then resultText = "PASS"
    SampleDesignBandRulesPassed = resultText
End Function

Private Function KoPersonUnit() As String
    KoPersonUnit = ChrW(&HBA85)
End Function

Private Function CheckBool(ByVal resultText As String, ByVal caseName As String, ByVal actual As Boolean, ByVal expected As Boolean) As String
    If actual <> expected Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected " & CStr(expected) & ", got " & CStr(actual)
    End If
    CheckBool = resultText
End Function

Private Function CheckDouble(ByVal resultText As String, ByVal caseName As String, ByVal actual As Double, ByVal expected As Double) As String
    If Abs(actual - expected) > 0.0000001 Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected " & CStr(expected) & ", got " & CStr(actual)
    End If
    CheckDouble = resultText
End Function

Private Function CheckContains(ByVal resultText As String, ByVal caseName As String, ByVal actual As String, ByVal expectedPart As String) As String
    If InStr(1, actual, expectedPart, vbBinaryCompare) = 0 Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected formula to contain " & expectedPart & ", got " & actual
    End If
    CheckContains = resultText
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
    $module.CodeModule.AddFromString($moduleCode + "`r`n`r`n" + $wrapperCode)

    $result = $excel.Run("SampleDesignBandRulesPassed")
    if ($result -ne "PASS") {
        throw $result
    }

    Write-Host "Sample design band rule tests passed."
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
