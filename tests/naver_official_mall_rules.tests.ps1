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
    "NaverBusinessSearch_BuildQueryForTest",
    "NaverBusinessSearch_ResultHeaderTextForTest",
    "NaverBusinessSearch_BuildQuery",
    "NaverBusinessSearch_ResultHeaders",
    "NaverOfficialMall_ClassifyUrlForTest",
    "NaverOfficialMall_ClassifyUrl",
    "NaverOfficialMall_Domain",
    "NaverOfficialMall_DomainMatches",
    "NaverOfficialMall_IsOpenMarketDomain",
    "NaverOfficialMall_IsSocialOrBlogDomain",
    "NaverOfficialMall_HasMallEvidence"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function OfficialMallRulesPassed() As String
    Dim resultText As String
    resultText = CheckText(resultText, "business search query joins company and user keyword", NaverBusinessSearch_BuildQueryForTest("  ACME  ", BusinessSearchOfficialSiteKeyword()), "ACME " & BusinessSearchOfficialSiteKeyword())
    resultText = CheckText(resultText, "business search query falls back to company when keyword is blank", NaverBusinessSearch_BuildQueryForTest("ACME", " "), "ACME")
    resultText = CheckText(resultText, "business search result headers include local address and official site columns", NaverBusinessSearch_ResultHeaderTextForTest(), BusinessSearchExpectedHeaderText())
    resultText = CheckCase(resultText, "https://brand.naver.com/acme", "ACME official brand store", "", "ACME", "exclude:naver_brand_store")
    resultText = CheckCase(resultText, "https://smartstore.naver.com/acme", "ACME official smart store", "", "ACME", "exclude:naver_smartstore")
    resultText = CheckCase(resultText, "https://www.coupang.com/vp/products/123", "ACME official", "", "ACME", "exclude:open_market")
    resultText = CheckCase(resultText, "https://blog.naver.com/acme", "ACME blog", "", "ACME", "exclude:social_or_blog")
    resultText = CheckCase(resultText, "https://shop.acme.co.kr/goods/1", "ACME official mall", "cart payment", "ACME", "candidate:self_mall")
    If Len(resultText) = 0 Then resultText = "PASS"
    OfficialMallRulesPassed = resultText
End Function

Private Function BusinessSearchOfficialSiteKeyword() As String
    BusinessSearchOfficialSiteKeyword = ChrW(&HACF5) & ChrW(&HC2DD) & " " & ChrW(&HC0AC) & ChrW(&HC774) & ChrW(&HD2B8)
End Function

Private Function BusinessSearchExpectedHeaderText() As String
    BusinessSearchExpectedHeaderText = _
        ChrW(&HAC80) & ChrW(&HC0C9) & ChrW(&HC5B4) & "|" & _
        ChrW(&HC9C0) & ChrW(&HC5ED) & ChrW(&HAC80) & ChrW(&HC0C9) & " " & ChrW(&HC5C5) & ChrW(&HCCB4) & ChrW(&HBA85) & "|" & _
        ChrW(&HC9C0) & ChrW(&HC5ED) & ChrW(&HAC80) & ChrW(&HC0C9) & " " & ChrW(&HB9C1) & ChrW(&HD06C) & "|" & _
        ChrW(&HC804) & ChrW(&HD654) & ChrW(&HBC88) & ChrW(&HD638) & "|" & _
        ChrW(&HC9C0) & ChrW(&HBC88) & ChrW(&HC8FC) & ChrW(&HC18C) & "|" & _
        ChrW(&HB3C4) & ChrW(&HB85C) & ChrW(&HBA85) & ChrW(&HC8FC) & ChrW(&HC18C) & "|" & _
        ChrW(&HC6F9) & ChrW(&HBB38) & ChrW(&HC11C) & " " & ChrW(&HD6C4) & ChrW(&HBCF4) & " URL|" & _
        ChrW(&HC6F9) & ChrW(&HBB38) & ChrW(&HC11C) & " " & ChrW(&HD6C4) & ChrW(&HBCF4) & " " & ChrW(&HC81C) & ChrW(&HBAA9) & "|" & _
        ChrW(&HC6F9) & ChrW(&HBB38) & ChrW(&HC11C) & " " & ChrW(&HD6C4) & ChrW(&HBCF4) & " " & ChrW(&HC694) & ChrW(&HC57D) & "|" & _
        ChrW(&HC6F9) & ChrW(&HBB38) & ChrW(&HC11C) & " " & ChrW(&HD6C4) & ChrW(&HBCF4) & " " & ChrW(&HB3C4) & ChrW(&HBA54) & ChrW(&HC778) & "|" & _
        ChrW(&HAC80) & ChrW(&HC0C9) & ChrW(&HC77C) & ChrW(&HC2DC)
End Function

Private Function CheckText(ByVal resultText As String, ByVal caseName As String, ByVal actual As String, ByVal expected As String) As String
    If actual <> expected Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & caseName & ": expected " & expected & ", got " & actual
    End If
    CheckText = resultText
End Function

Private Function CheckCase(ByVal resultText As String, ByVal urlValue As String, ByVal titleText As String, ByVal descriptionText As String, ByVal companyName As String, ByVal expected As String) As String
    Dim actual As String
    actual = NaverOfficialMall_ClassifyUrlForTest(urlValue, titleText, descriptionText, companyName)
    If actual <> expected Then
        If Len(resultText) > 0 Then resultText = resultText & vbLf
        resultText = resultText & "Expected " & expected & ", got " & actual & " for " & urlValue
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

    $result = $excel.Run("OfficialMallRulesPassed")
    if ($result -ne "PASS") {
        throw $result
    }

    Write-Host "Naver official mall rule tests passed."
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
