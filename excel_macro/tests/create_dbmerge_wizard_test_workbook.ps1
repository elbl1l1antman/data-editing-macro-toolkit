param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$outputPath = Join-Path $ProjectRoot "dev\dbmerge_wizard_test.xlsx"
$excel = $null
$workbook = $null

function Set-SheetData {
    param($Sheet, [object[]]$Headers, [object[]]$Rows, [int]$HeaderRow = 1)

    for ($c = 0; $c -lt $Headers.Count; $c++) {
        $Sheet.Cells.Item($HeaderRow, $c + 1).Value2 = $Headers[$c]
    }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $Rows[$r].Values.Count; $c++) {
            $Sheet.Cells.Item($HeaderRow + $r + 1, $c + 1).Value2 = [string]$Rows[$r].Values[$c]
        }
    }
    $Sheet.Rows.Item($HeaderRow).Font.Bold = $true
    $Sheet.UsedRange.Columns.AutoFit() | Out-Null
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -lt 5) {
        $workbook.Worksheets.Add() | Out-Null
    }

    $seoul = $workbook.Worksheets.Item(1)
    $seoul.Name = "서울"
    Set-SheetData $seoul @("ID", "사업체명", "주소", "매출액") @(
        @{ Values = @( "S1", "서울상사", "서울 중구", 100 ) },
        @{ Values = @( "S2", "한빛", "서울 강남구", 200 ) }
    )

    $busan = $workbook.Worksheets.Item(2)
    $busan.Name = "부산"
    Set-SheetData $busan @("관리번호", "사업체명", "주소", "종사자수") @(
        @{ Values = @( "B1", "부산기업", "부산 해운대구", 5 ) }
    )

    $noId = $workbook.Worksheets.Item(3)
    $noId.Name = "ID없음"
    Set-SheetData $noId @("사업체명", "주소", "매출액") @(
        @{ Values = @( "무지개", "대전 서구", 300 ) }
    )

    $twoRows = $workbook.Worksheets.Item(4)
    $twoRows.Name = "2줄헤더"
    Set-SheetData $twoRows @("관리 정보", "업체 정보", "소재지", "매출 정보") @() 1
    Set-SheetData $twoRows @("case_id", "사업체명", "주소", "매출액") @(
        @{ Values = @( "T1", "두줄기업", "인천 남동구", 400 ) }
    ) 2

    $internal = $workbook.Worksheets.Item(5)
    $internal.Name = "_내부자료"
    $internal.Cells.Item(1, 1).Value2 = "내부"
    $internal.Visible = 0

    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }
    $workbook.SaveAs($outputPath, 51)
    Write-Host "Created DBMerge wizard test workbook: $outputPath"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
