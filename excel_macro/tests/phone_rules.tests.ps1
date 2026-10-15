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
    "Phone_NormalizeForTest",
    "Phone_NormalizeForTestWithOptions",
    "Phone_NormalizeResultForTest",
    "Phone_ClassifyForTest",
    "Phone_SplitRepresentativeForTest",
    "NormalizeKoreanPhone",
    "NormalizeKoreanPhoneWithOptions",
    "Phone_DefaultNormalizeOptions",
    "Phone_NormalizeOptionEnabled",
    "Phone_ApplyLeadingZeroOption",
    "Phone_ShouldWriteReviewInstead",
    "Phone_IsKnownRepresentativePrefix",
    "Phone_TrySplitRepresentative",
    "Phone_IsSequentialPlaceholderDigits",
    "Phone_SubscriberDigitsForSequenceCheck",
    "Phone_IsSequentialRun",
    "Phone_BlockHasDistinctDigits",
    "Phone_HasRepeatedBlockPattern",
    "Phone_ClassifyOne",
    "Phone_IsPlaceholderZeroPattern",
    "Phone_HasRepeatedDigitRun",
    "Phone_DigitsForKoreanRules",
    "Phone_IsSafeNumberDigits",
    "Phone_IsSafeNumberText",
    "DigitsOnly"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function PhoneRulesPassed() As String
    Dim resultText As String
    Dim representativePrefixes As Variant
    Dim representativePrefix As Variant

    representativePrefixes = Array("1588", "1577", "1899", "1544", "1644", "1661", "1566", "1600", "1670", "1688", "1666", "1599", "1877", "1855", "1800")
    For Each representativePrefix In representativePrefixes
        resultText = CheckCase(resultText, "prioritize representative prefix " & CStr(representativePrefix), _
            Phone_NormalizeForTestWithOptions(CStr(representativePrefix) & "1234", False, False, True), _
            CStr(representativePrefix) & "-1234")
    Next representativePrefix

    resultText = CheckCase(resultText, "normalize 8-digit representative number", Phone_NormalizeForTest("15881222"), "1588-1222")
    resultText = CheckCase(resultText, "keep valid safe number classification", Phone_ClassifyForTest("0507-2468-1357"), KoSafeNumber() & "|" & KoNormal())
    resultText = CheckCase(resultText, "keep unformatted safe number classification", Phone_ClassifyForTest("050724681357"), KoSafeNumber() & "|" & KoNormal())
    resultText = CheckCase(resultText, "restore missing zero safe number classification", Phone_ClassifyForTest("5072468135"), KoSafeNumber() & "|" & KoNormal())
    resultText = CheckCase(resultText, "keep international safe number classification", Phone_ClassifyForTest("+82-507-246-8135"), KoSafeNumber() & "|" & KoNormal())
    resultText = CheckCase(resultText, "reject three digit 050 group as safe number", Phone_ClassifyForTest("050-7246-8135"), KoUnknown() & "|" & KoReview())
    resultText = CheckCase(resultText, "reject trailing 0507 digits as safe number", Phone_ClassifyForTest("050-2468-0507"), KoUnknown() & "|" & KoReview())
    resultText = CheckCase(resultText, "do not normalize malformed 050 group as safe number", Phone_NormalizeResultForTest("050-7246-8135"), "FALSE|")
    resultText = CheckCase(resultText, "do not normalize trailing 0507 digits as safe number", Phone_NormalizeResultForTest("050-2468-0507"), "FALSE|")
    resultText = CheckCase(resultText, "keep mobile ending in 0507 as mobile", Phone_ClassifyForTest("010-1234-0507"), KoMobile() & "|" & KoNormal())
    resultText = CheckCase(resultText, "split compact multiple representative number", Phone_SplitRepresentativeForTest("1588-12345"), "TRUE|1588-1234|1588-1235")
    resultText = CheckCase(resultText, "split tilde multiple representative number", Phone_SplitRepresentativeForTest("1588-1234~5"), "TRUE|1588-1234|1588-1235")
    resultText = CheckCase(resultText, "classify multiple representative number for review", Phone_ClassifyForTest("1588-12345"), KoRepresentative() & "|" & KoRepresentativeMultipleReview())
    resultText = CheckCase(resultText, "classify 000-0000 as error", Phone_ClassifyForTest("000-0000"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "classify 0000-0000 as error", Phone_ClassifyForTest("0000-0000"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "keep four repeated digits normal", Phone_ClassifyForTest("02-231-4444"), KoSeoulArea() & "|" & KoNormal())
    resultText = CheckCase(resultText, "flag five repeated digits for review", Phone_ClassifyForTest("02-1111-1222"), KoSeoulArea() & "|" & KoReview())
    resultText = CheckCase(resultText, "classify six repeated digits as error", Phone_ClassifyForTest("02-1111-1122"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "classify repeated two digit block as error", Phone_ClassifyForTest("02-1212-3456"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "classify repeated three digit block as error", Phone_ClassifyForTest("031-456-4567"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "classify repeated four digit block as error", Phone_ClassifyForTest("010-1234-1234"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "keep ordinary 8-digit representative number normal", Phone_ClassifyForTest("1588-1222"), KoRepresentative() & "|" & KoNormal())
    resultText = CheckCase(resultText, "prepend zero option repairs missing Seoul prefix", Phone_NormalizeForTestWithOptions("221345678", True, False, False), "02-2134-5678")
    resultText = CheckCase(resultText, "repeated run option writes review text instead of number", Phone_NormalizeForTestWithOptions("02-231-4444", False, True, False), KoReview())
    resultText = CheckCase(resultText, "mobile option treats 8 digits as 010 personal phone", Phone_NormalizeForTestWithOptions("24681357", False, False, True), "010-2468-1357")
    resultText = CheckCase(resultText, "default keeps 8 digits as representative number", Phone_NormalizeForTestWithOptions("24681357", False, False, False), "2468-1357")
    resultText = CheckCase(resultText, "delete sequential local phone during normalization", Phone_NormalizeResultForTest("02-123-4567"), "TRUE|")
    resultText = CheckCase(resultText, "delete sequential 8 digit placeholder during normalization", Phone_NormalizeResultForTest("12345678"), "TRUE|")
    resultText = CheckCase(resultText, "classify sequential local phone as error", Phone_ClassifyForTest("02-123-4567"), KoErrorPattern() & "|" & KoError())
    resultText = CheckCase(resultText, "classify sequential mobile subscriber digits as error", Phone_ClassifyForTest("010-1234-5678"), KoErrorPattern() & "|" & KoError())

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

Private Function KoRepresentativeMultipleReview() As String
    KoRepresentativeMultipleReview = KoRepresentative() & " " & ChrW(&HBCF5) & ChrW(&HC218) & " " & KoReview()
End Function

Private Function KoNormal() As String
    KoNormal = ChrW(&HC815) & ChrW(&HC0C1)
End Function

Private Function KoSafeNumber() As String
    KoSafeNumber = ChrW(&HC548) & ChrW(&HC2EC) & ChrW(&HBC88) & ChrW(&HD638)
End Function

Private Function KoUnknown() As String
    KoUnknown = ChrW(&HC54C) & " " & ChrW(&HC218) & " " & ChrW(&HC5C6) & ChrW(&HC74C)
End Function

Private Function KoMobile() As String
    KoMobile = ChrW(&HD734) & ChrW(&HB300) & ChrW(&HD3F0)
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
