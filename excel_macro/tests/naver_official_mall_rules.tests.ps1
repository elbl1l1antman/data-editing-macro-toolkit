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
    resultText = CheckCase(resultText, "https://brand.naver.com/acme", "ACME official brand store", "", "ACME", "exclude:naver_brand_store")
    resultText = CheckCase(resultText, "https://smartstore.naver.com/acme", "ACME official smart store", "", "ACME", "exclude:naver_smartstore")
    resultText = CheckCase(resultText, "https://www.coupang.com/vp/products/123", "ACME official", "", "ACME", "exclude:open_market")
    resultText = CheckCase(resultText, "https://blog.naver.com/acme", "ACME blog", "", "ACME", "exclude:social_or_blog")
    resultText = CheckCase(resultText, "https://shop.acme.co.kr/goods/1", "ACME official mall", "cart payment", "ACME", "candidate:self_mall")
    If Len(resultText) = 0 Then resultText = "PASS"
    OfficialMallRulesPassed = resultText
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
