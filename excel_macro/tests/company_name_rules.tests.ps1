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
    "CompanyName_DefaultOptions",
    "CompanyName_CleanOne",
    "CompanyName_RemoveSpaces",
    "CompanyName_HasCorpMarker",
    "CompanyName_RemoveCorpMarkers",
    "CompanyName_RemoveBranchMarkers",
    "CompanyName_RemoveLegalEntityMarkers",
    "CompanyName_RemoveKoreanLegalEntityMarkers",
    "CompanyName_RemoveLegalMarkerAtEdges",
    "CompanyName_RemoveForeignLegalEntityMarkers",
    "CompanyName_CleanLeadingSeparators",
    "CompanyName_CleanTrailingSeparators",
    "CompanyName_IsLegalSeparator",
    "CompanyName_RemoveLeadingBranchMarker",
    "CompanyName_RemoveTrailingBranchMarker",
    "CompanyName_RemoveKnownTrailingBranch",
    "CompanyName_IsBranchSuffix",
    "CompanyName_VisibleLength",
    "CompanyName_NormalizeCorpSymbol",
    "CompanyName_RemoveSpecialChars",
    "CompanyName_IsNameCharacter",
    "CompanyName_MoveCorpMarker",
    "CompanyName_PronounceEnglishWords",
    "CompanyName_IsAsciiLetter",
    "CompanyName_EnglishWordKorean",
    "CompanyName_SpellAlphabetKorean",
    "CompanyName_AlphabetKorean"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function CompanyNameRulesPassed() As String
    Dim resultText As String
    Dim options As Object
    Set options = CompanyName_DefaultOptions()

    resultText = CheckCase(resultText, "number inside Hangul company name keeps Hangul", CompanyName_CleanOne("A4" & KoSystem(), options), KoA() & "4" & KoSystem())
    resultText = CheckCase(resultText, "number-prefixed English company name keeps word", CompanyName_CleanOne("3S" & KoTech(), options), "3" & KoS() & KoTech())
    resultText = CheckCase(resultText, "mixed Hangul and acronym keeps Hangul", CompanyName_CleanOne(KoJangHan() & "MRS", options), KoJangHan() & KoM() & KoR() & KoS())
    resultText = CheckCase(resultText, "numeric branch marker does not strip core company name", CompanyName_CleanOne(KoSamsung() & "4" & KoBranchPoint(), options), KoSamsung() & "4" & KoBranchPoint())

    options("removeSpecialChars") = True
    resultText = CheckCase(resultText, "special char removal keeps Hangul around number", CompanyName_CleanOne("A4-" & KoSystem(), options), KoA() & "4" & KoSystem())

    If Len(resultText) = 0 Then resultText = "PASS"
    CompanyNameRulesPassed = resultText
End Function

Private Function CheckCase(ByVal resultText As String, ByVal caseName As String, ByVal actual As String, ByVal expected As String) As String
    If actual <> expected Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected " & expected & ", got " & actual
    End If
    CheckCase = resultText
End Function

Private Function KoA() As String
    KoA = ChrW(&HC5D0) & ChrW(&HC774)
End Function

Private Function KoS() As String
    KoS = ChrW(&HC5D0) & ChrW(&HC2A4)
End Function

Private Function KoM() As String
    KoM = ChrW(&HC5E0)
End Function

Private Function KoR() As String
    KoR = ChrW(&HC54C)
End Function

Private Function KoSystem() As String
    KoSystem = ChrW(&HC2DC) & ChrW(&HC2A4) & ChrW(&HD15C)
End Function

Private Function KoTech() As String
    KoTech = ChrW(&HD14C) & ChrW(&HD06C)
End Function

Private Function KoJangHan() As String
    KoJangHan = ChrW(&HC7A5) & ChrW(&HD55C)
End Function

Private Function KoSamsung() As String
    KoSamsung = ChrW(&HC0BC) & ChrW(&HC131)
End Function

Private Function KoBranchPoint() As String
    KoBranchPoint = ChrW(&HC810)
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

    $result = $excel.Run("CompanyNameRulesPassed")
    if ($result -ne "PASS") {
        throw $result
    }

    Write-Host "Company name rule tests passed."
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
