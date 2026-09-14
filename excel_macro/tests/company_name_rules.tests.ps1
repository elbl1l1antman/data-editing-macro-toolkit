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
    "CompanyName_DefaultOperationOrder",
    "CompanyName_UserRequestedOperationOrder",
    "CompanyName_PreserveParentheticalFirstOrder",
    "CompanyName_CleanOne",
    "CompanyName_ApplyOperation",
    "CompanyName_RemoveSpaces",
    "CompanyName_HasCorpMarker",
    "CompanyName_RemoveCorpMarkers",
    "CompanyName_RemoveParentheticalText",
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
    "CompanyName_ReplaceBusinessSymbols",
    "CompanyName_RemoveSpecialChars",
    "CompanyName_IsNameCharacter",
    "CompanyName_MoveCorpMarker",
    "CompanyName_RemoveBilingualDuplicates",
    "CompanyName_BilingualBlocksMatch",
    "CompanyName_KoreanComparable",
    "CompanyName_EnglishBlockKoreanComparable",
    "CompanyName_ReadKoreanBlockEnd",
    "CompanyName_ReadEnglishBlockEnd",
    "CompanyName_SkipBilingualSeparators",
    "CompanyName_NextNonBilingualSeparatorPos",
    "CompanyName_ReadTrailingBilingualSeparators",
    "CompanyName_IsBilingualSeparator",
    "CompanyName_IsHangulLetter",
    "CompanyName_PronounceEnglishWords",
    "CompanyName_PronounceEnglishToken",
    "CompanyName_LongestEnglishWordKorean",
    "CompanyName_TryEnglishWordKorean",
    "CompanyName_IsAsciiLetter",
    "CompanyName_IsAsciiDigit",
    "CompanyName_IsAsciiLetterOrDigit",
    "CompanyName_DigitKorean",
    "CompanyName_EnglishWordKorean",
    "CompanyName_KRDictLoanwordKorean",
    "CompanyName_KRDictLoanwordKorean_A",
    "CompanyName_KRDictLoanwordKorean_B",
    "CompanyName_KRDictLoanwordKorean_C",
    "CompanyName_KRDictLoanwordKorean_D",
    "CompanyName_KRDictLoanwordKorean_E",
    "CompanyName_KRDictLoanwordKorean_F",
    "CompanyName_KRDictLoanwordKorean_G",
    "CompanyName_KRDictLoanwordKorean_H",
    "CompanyName_KRDictLoanwordKorean_I",
    "CompanyName_KRDictLoanwordKorean_J",
    "CompanyName_KRDictLoanwordKorean_K",
    "CompanyName_KRDictLoanwordKorean_L",
    "CompanyName_KRDictLoanwordKorean_M",
    "CompanyName_KRDictLoanwordKorean_N",
    "CompanyName_KRDictLoanwordKorean_O",
    "CompanyName_KRDictLoanwordKorean_P",
    "CompanyName_KRDictLoanwordKorean_Q",
    "CompanyName_KRDictLoanwordKorean_R",
    "CompanyName_KRDictLoanwordKorean_S",
    "CompanyName_KRDictLoanwordKorean_T",
    "CompanyName_KRDictLoanwordKorean_U",
    "CompanyName_KRDictLoanwordKorean_V",
    "CompanyName_KRDictLoanwordKorean_W",
    "CompanyName_KRDictLoanwordKorean_X",
    "CompanyName_KRDictLoanwordKorean_Y",
    "CompanyName_KRDictLoanwordKorean_Z",
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
    resultText = CheckCase(resultText, "business ampersand is pronounced instead of removed", CompanyName_CleanOne("A&B", options), KoA() & KoAnd() & KoB())
    resultText = CheckCase(resultText, "business co dot is pronounced instead of removed", CompanyName_CleanOne("ABC Co.", options), KoA() & KoB() & KoC() & KoCompanyRequested())

    Set options = CompanyName_DefaultOptions()
    options("removeLegalEntityMarkers") = True
    options("removeSpecialChars") = True
    resultText = CheckCase(resultText, "foreign legal suffix removal takes precedence over co dot pronunciation", CompanyName_CleanOne("ABC Co., Ltd.", options), KoA() & KoB() & KoC())

    Set options = CompanyName_DefaultOptions()
    options("removeParentheticalText") = True
    resultText = CheckCase(resultText, "remove halfwidth parenthetical publisher prefix", CompanyName_CleanOne("(" & KoBookPublisher() & ")" & KoYongYongDragon(), options), KoYongYongDragon())
    resultText = CheckCase(resultText, "remove fullwidth parenthetical publisher prefix", CompanyName_CleanOne(FullLeftParen() & KoBookPublisher() & FullRightParen() & KoLifeTree(), options), KoLifeTree())
    resultText = CheckCase(resultText, "remove fullwidth short publisher prefix", CompanyName_CleanOne(FullLeftParen() & KoDoShort() & FullRightParen() & KoSpring(), options), KoSpring())

    Set options = CompanyName_DefaultOptions()
    options("pronounceEnglishWords") = True
    options("spellAlphabetKorean") = True
    resultText = CheckCase(resultText, "known English suffix is transliterated before alphabet spelling", CompanyName_CleanOne("AIsystem", options), KoAI() & KoSystem())
    resultText = CheckCase(resultText, "acronym prefix and known English suffix both survive combined conversion", CompanyName_CleanOne("KTmobile", options), KoKT() & KoMobile())
    resultText = CheckCase(resultText, "soft is transliterated as soft not software", CompanyName_CleanOne("SOFT", options), KoSoft())
    resultText = CheckCase(resultText, "software is transliterated as software", CompanyName_CleanOne("SOFTWARE", options), KoSoftware())
    resultText = CheckCase(resultText, "medical sector word is transliterated inside combined token", CompanyName_CleanOne("Smartclinic", options), KoSmart() & KoClinic())
    resultText = CheckCase(resultText, "agriculture sector word is transliterated inside combined token", CompanyName_CleanOne("Greenfarm", options), KoGreen() & KoFarm())
    resultText = CheckCase(resultText, "real estate sector word is transliterated inside combined token", CompanyName_CleanOne("Cityrealty", options), KoCity() & KoRealty())
    resultText = CheckCase(resultText, "security sector word is transliterated inside combined token", CompanyName_CleanOne("Websecurity", options), KoWeb() & KoSecurity())
    resultText = CheckCase(resultText, "food service sector words are transliterated inside combined token", CompanyName_CleanOne("Coffeehouse", options), KoCoffee() & KoHouse())
    resultText = CheckCase(resultText, "KRDict English loanword maps computer inside combined token", CompanyName_CleanOne("Computeracademy", options), KoComputer() & KoAcademy())
    resultText = CheckCase(resultText, "KRDict English loanword maps laptop inside combined token", CompanyName_CleanOne("Laptopservice", options), KoLaptop() & KoService())
    resultText = CheckCase(resultText, "KRDict English loanword maps keyboard inside combined token", CompanyName_CleanOne("Keyboardcenter", options), KoKeyboard() & KoCenter())
    resultText = CheckCase(resultText, "space removal can run before English word conversion by selected priority", CompanyName_CleanOne("KT mobile", options), KoKT() & KoMobile())

    Set options = CompanyName_DefaultOptions()
    options("removeSpaces") = False
    options("spellAlphabetKorean") = False
    options("removeSpecialChars") = True
    options("removeParentheticalText") = True
    options("operationOrder") = CompanyName_UserRequestedOperationOrder()
    resultText = CheckCase(resultText, "selected priority can remove special chars before parenthetical text", CompanyName_CleanOne("(" & KoBookPublisher() & ")" & KoYongYongDragon(), options), KoBookPublisher() & KoYongYongDragon())

    Set options = CompanyName_DefaultOptions()
    options("removeSpaces") = False
    options("spellAlphabetKorean") = False
    options("removeSpecialChars") = True
    options("removeParentheticalText") = True
    options("operationOrder") = CompanyName_PreserveParentheticalFirstOrder()
    resultText = CheckCase(resultText, "selected priority can remove parenthetical text before special chars", CompanyName_CleanOne("(" & KoBookPublisher() & ")" & KoYongYongDragon(), options), KoYongYongDragon())

    Set options = CompanyName_DefaultOptions()
    options("pronounceEnglishWords") = False
    options("spellAlphabetKorean") = False
    options("removeBilingualDuplicates") = True
    resultText = CheckCase(resultText, "remove Korean then English duplicate company name", CompanyName_CleanOne(KoK() & KoIndustry() & "Kindustry", options), KoK() & KoIndustry())
    resultText = CheckCase(resultText, "remove English then Korean duplicate company name", CompanyName_CleanOne("Kindustry" & KoK() & KoIndustry(), options), KoK() & KoIndustry())
    resultText = CheckCase(resultText, "remove Korean and digit-bearing English duplicate company name", CompanyName_CleanOne(KoForward() & KoThreeD() & KoKorea() & "FORWARD3DKOREA", options), KoForward() & KoThreeD() & KoKorea())
    resultText = CheckCase(resultText, "remove Korean and compact brand English duplicate company name", CompanyName_CleanOne(KoTvaSoft() & "TVASOFT", options), KoTvaSoft())
    resultText = CheckCase(resultText, "keep non-equivalent Korean and English mixed company name", CompanyName_CleanOne(KoK() & KoIndustry() & "Korea", options), KoK() & KoIndustry() & "Korea")

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

Private Function KoB() As String
    KoB = ChrW(&HBE44)
End Function

Private Function KoC() As String
    KoC = ChrW(&HC528)
End Function

Private Function KoAI() As String
    KoAI = KoA() & ChrW(&HC544) & ChrW(&HC774)
End Function

Private Function KoAnd() As String
    KoAnd = ChrW(&HC564)
End Function

Private Function KoCompanyRequested() As String
    KoCompanyRequested = ChrW(&HCEF4) & ChrW(&HD328) & ChrW(&HB2C8)
End Function

Private Function KoKT() As String
    KoKT = ChrW(&HCF00) & ChrW(&HC774) & ChrW(&HD2F0)
End Function

Private Function KoAcademy() As String
    KoAcademy = ChrW(&HC544) & ChrW(&HCE74) & ChrW(&HB370) & ChrW(&HBBF8)
End Function

Private Function KoComputer() As String
    KoComputer = ChrW(&HCEF4) & ChrW(&HD4E8) & ChrW(&HD130)
End Function

Private Function KoLaptop() As String
    KoLaptop = ChrW(&HB7A9) & ChrW(&HD1B1)
End Function

Private Function KoKeyboard() As String
    KoKeyboard = ChrW(&HD0A4) & ChrW(&HBCF4) & ChrW(&HB4DC)
End Function

Private Function KoK() As String
    KoK = ChrW(&HCF00) & ChrW(&HC774)
End Function

Private Function KoIndustry() As String
    KoIndustry = ChrW(&HC778) & ChrW(&HB354) & ChrW(&HC2A4) & ChrW(&HD2B8) & ChrW(&HB9AC)
End Function

Private Function KoForward() As String
    KoForward = ChrW(&HD3EC) & ChrW(&HC6CC) & ChrW(&HB4DC)
End Function

Private Function KoThreeD() As String
    KoThreeD = ChrW(&HC4F0) & ChrW(&HB9AC) & ChrW(&HB514)
End Function

Private Function KoKorea() As String
    KoKorea = ChrW(&HCF54) & ChrW(&HB9AC) & ChrW(&HC544)
End Function

Private Function KoTvaSoft() As String
    KoTvaSoft = ChrW(&HD2F0) & ChrW(&HBC14) & ChrW(&HC18C) & ChrW(&HD504) & ChrW(&HD2B8)
End Function

Private Function KoMobile() As String
    KoMobile = ChrW(&HBAA8) & ChrW(&HBC14) & ChrW(&HC77C)
End Function

Private Function KoSmart() As String
    KoSmart = ChrW(&HC2A4) & ChrW(&HB9C8) & ChrW(&HD2B8)
End Function

Private Function KoClinic() As String
    KoClinic = ChrW(&HD074) & ChrW(&HB9AC) & ChrW(&HB2C9)
End Function

Private Function KoGreen() As String
    KoGreen = ChrW(&HADF8) & ChrW(&HB9B0)
End Function

Private Function KoFarm() As String
    KoFarm = ChrW(&HD31C)
End Function

Private Function KoCity() As String
    KoCity = ChrW(&HC2DC) & ChrW(&HD2F0)
End Function

Private Function KoRealty() As String
    KoRealty = ChrW(&HB9AC) & ChrW(&HC5BC) & ChrW(&HD2F0)
End Function

Private Function KoWeb() As String
    KoWeb = ChrW(&HC6F9)
End Function

Private Function KoSecurity() As String
    KoSecurity = ChrW(&HC2DC) & ChrW(&HD050) & ChrW(&HB9AC) & ChrW(&HD2F0)
End Function

Private Function KoService() As String
    KoService = ChrW(&HC11C) & ChrW(&HBE44) & ChrW(&HC2A4)
End Function

Private Function KoSoft() As String
    KoSoft = ChrW(&HC18C) & ChrW(&HD504) & ChrW(&HD2B8)
End Function

Private Function KoSoftware() As String
    KoSoftware = KoSoft() & ChrW(&HC6E8) & ChrW(&HC5B4)
End Function

Private Function KoCenter() As String
    KoCenter = ChrW(&HC13C) & ChrW(&HD130)
End Function

Private Function KoCoffee() As String
    KoCoffee = ChrW(&HCEE4) & ChrW(&HD53C)
End Function

Private Function KoHouse() As String
    KoHouse = ChrW(&HD558) & ChrW(&HC6B0) & ChrW(&HC2A4)
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

Private Function FullLeftParen() As String
    FullLeftParen = ChrW(&HFF08)
End Function

Private Function FullRightParen() As String
    FullRightParen = ChrW(&HFF09)
End Function

Private Function KoBookPublisher() As String
    KoBookPublisher = ChrW(&HB3C4) & ChrW(&HC11C) & ChrW(&HCD9C) & ChrW(&HD310)
End Function

Private Function KoDoShort() As String
    KoDoShort = ChrW(&HB3C4)
End Function

Private Function KoYongYongDragon() As String
    KoYongYongDragon = ChrW(&HC6A9) & ChrW(&HC6A9) & ChrW(&HB4DC) & ChrW(&HB798) & ChrW(&HACE4)
End Function

Private Function KoLifeTree() As String
    KoLifeTree = ChrW(&HC0DD) & ChrW(&HBA85) & ChrW(&HB098) & ChrW(&HBB34)
End Function

Private Function KoSpring() As String
    KoSpring = ChrW(&HBD04)
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
