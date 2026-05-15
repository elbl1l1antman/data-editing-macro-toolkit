param(
    [string]$WorkbookPath = "C:\Users\COM\OneDrive\Program\excel_macro\dev\DataEditingMacroToolkit_dev.xlsm",
    [string]$ModulePath = "C:\Users\COM\OneDrive\Program\excel_macro\src\MacroToolkit.bas"
)

$ErrorActionPreference = "Stop"

$moduleText = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
$toolkitVersion = "unknown"
if ($moduleText -match 'TOOLKIT_VERSION\s+As\s+String\s*=\s*"([^"]+)"') {
    $toolkitVersion = $Matches[1]
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "MacroToolkit"
    $sheet.Cells.Clear()
    $sheet.Range("A1").Value = "데이터 편집 매크로 툴킷"
    $sheet.Range("A2").Value = "버전"
    $sheet.Range("B2").Value = $toolkitVersion
    $sheet.Range("A4").Value = "설치 상태"
    $sheet.Range("B4").Value = "VBA 모듈 자동 삽입 여부는 Excel 보안 설정에 따라 달라집니다."
    $sheet.Range("A6").Value = "구현 기준 파일"
    $sheet.Range("B6").Value = "src\MacroToolkit.bas"
    $sheet.Range("A8").Value = "확인할 시트"
    $sheet.Range("B8").Value = "README, 매크로 목록, VBA 소스 미리보기"
    $sheet.Range("A10").Value = "포함 매크로"
    $macros = @(
        "ConvertTextNumbersInSelection",
        "ConvertTextNumbersInUsedRange",
        "ConvertTextNumbersInColumns",
        "NormalizePhonesInSelection",
        "NormalizePhonesInColumns",
        "OutlierDetect_Setup",
        "OutlierDetect_Run"
    )
    for ($i = 0; $i -lt $macros.Count; $i++) {
        $sheet.Cells.Item(11 + $i, 1).Value2 = [string]$macros[$i]
    }
    $sheet.Range("A1:B1").Font.Bold = $true
    $sheet.Range("A2:A10").Font.Bold = $true
    $sheet.Columns("A:B").AutoFit()
    $workbook.Save()
    Write-Output "Updated shell sheet: $WorkbookPath"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($true) | Out-Null } catch {} }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
