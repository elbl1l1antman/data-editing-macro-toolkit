param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot "src\MacroToolkit.bas") -Raw -Encoding UTF8
$publishSource = Get-Content -LiteralPath (Join-Path $ProjectRoot "scripts\publish_release.ps1") -Raw -Encoding UTF8
foreach ($token in @("Ribbon_AddressConvertProvinceNotation", 'label="광역시도 표기"', "12. 광역시도 표기 변환")) {
    if (-not ($source.Contains($token) -or $publishSource.Contains($token))) {
        throw "Address province UI token missing: $token"
    }
}

function Get-VbaBlock {
    param([string]$SourceText, [string]$ProcedureName)
    $start = [regex]::Match($SourceText, "(?m)^(Public|Private)\s+(Function|Sub)\s+$([regex]::Escape($ProcedureName))\b.*$")
    if (-not $start.Success) { throw "Procedure not found: $ProcedureName" }
    $kind = $start.Groups[2].Value
    $endToken = if ($kind -eq "Function") { "End Function" } else { "End Sub" }
    $tail = $SourceText.Substring($start.Index)
    $end = [regex]::Match($tail, "(?m)^$endToken\s*$")
    if (-not $end.Success) { throw "Procedure end not found: $ProcedureName" }
    $tail.Substring(0, $end.Index + $end.Length)
}

$procedures = @(
    "Address_ProvinceNotationForTest",
    "Address_ConvertProvinceFirstToken",
    "Address_FirstWhitespacePosition",
    "Address_ProvinceShortName",
    "Address_IsProvinceFormalToken",
    "Address_ProvinceFormalFromShort"
)
$moduleCode = "Option Explicit`r`n" + (($procedures | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$cases = @(
    @("expand first token only", "경기 고양시 세종빌딩", "formal", "latest", "경기도 고양시 세종빌딩"),
    @("latest Gangwon", "강원 춘천시", "formal", "latest", "강원특별자치도 춘천시"),
    @("legacy Gangwon", "강원 춘천시", "formal", "legacy", "강원도 춘천시"),
    @("latest Jeonbuk", "전북 전주시", "formal", "latest", "전북특별자치도 전주시"),
    @("legacy Jeonbuk", "전북 전주시", "formal", "legacy", "전라북도 전주시"),
    @("latest Jeju", "제주 제주시", "formal", "latest", "제주특별자치도 제주시"),
    @("legacy Jeju", "제주 제주시", "formal", "legacy", "제주도 제주시"),
    @("recognize legacy formal name", "전라북도 전주시", "short", "latest", "전북 전주시"),
    @("recognize latest formal name", "전북특별자치도 전주시", "short", "latest", "전북 전주시"),
    @("leave unmatched first token", "세종빌딩 경기", "formal", "latest", "세종빌딩 경기"),
    @("normalize legacy formal name", "강원도 춘천시", "normalize", "latest", "강원특별자치도 춘천시")
)
$latestMappings = @(
    @("서울", "서울특별시"), @("부산", "부산광역시"), @("대구", "대구광역시"),
    @("인천", "인천광역시"), @("광주", "광주광역시"), @("대전", "대전광역시"),
    @("울산", "울산광역시"), @("세종", "세종특별자치시"), @("경기", "경기도"),
    @("강원", "강원특별자치도"), @("충북", "충청북도"), @("충남", "충청남도"),
    @("전북", "전북특별자치도"), @("전남", "전라남도"), @("경북", "경상북도"),
    @("경남", "경상남도"), @("제주", "제주특별자치도")
)

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
    foreach ($procedure in $procedures) {
        $null = $module.CodeModule.ProcStartLine($procedure, 0)
    }
    foreach ($case in $cases) {
        $actual = $excel.Run("Address_ProvinceNotationForTest", $case[1], $case[2], $case[3])
        if ($actual -ne $case[4]) {
            throw "$($case[0]): expected $($case[4]), got $actual"
        }
    }
    foreach ($mapping in $latestMappings) {
        $actualFormal = $excel.Run("Address_ProvinceNotationForTest", "$($mapping[0]) 테스트", "formal", "latest")
        if ($actualFormal -ne "$($mapping[1]) 테스트") { throw "latest mapping failed: $($mapping[0])" }
        $actualShort = $excel.Run("Address_ProvinceNotationForTest", "$($mapping[1]) 테스트", "short", "latest")
        if ($actualShort -ne "$($mapping[0]) 테스트") { throw "short mapping failed: $($mapping[1])" }
    }
    Write-Host "Address province notation tests passed."
}
finally {
    if ($workbook -ne $null) { $workbook.Close($false) | Out-Null }
    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
