param(
    [string]$WorkbookPath = "C:\Users\COM\OneDrive\Program\excel_macro\dev\DataEditingMacroToolkit_dev.xlsm",
    [string]$ModulePath = "C:\Users\COM\OneDrive\Program\excel_macro\src\MacroToolkit.bas"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
if (-not (Test-Path -LiteralPath $ModulePath)) { throw "Module not found: $ModulePath" }

$moduleText = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
$toolkitVersion = "unknown"
if ($moduleText -match 'TOOLKIT_VERSION\s+As\s+String\s*=\s*"([^"]+)"') {
    $toolkitVersion = $Matches[1]
}

function Remove-SheetIfExists($workbook, [string]$name) {
    foreach ($sheet in @($workbook.Worksheets)) {
        if ($sheet.Name -eq $name) {
            $sheet.Delete()
            return
        }
    }
}

function Add-SheetAfter($workbook, [string]$name) {
    $sheet = $workbook.Worksheets.Add([System.Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
    $sheet.Name = $name
    return $sheet
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Open($WorkbookPath)

    Remove-SheetIfExists $workbook "README"
    Remove-SheetIfExists $workbook "매크로 목록"
    Remove-SheetIfExists $workbook "VBA 소스 미리보기"
    Remove-SheetIfExists $workbook "Macro List"
    Remove-SheetIfExists $workbook "VBA Source Preview"

    $readme = Add-SheetAfter $workbook "README"
    $readme.Range("A1").Value = "데이터 편집 매크로 툴킷"
    $readme.Range("A2").Value = "개발용 통합문서"
    $readme.Range("A4").Value = "현재 버전"
    $readme.Range("B4").Value = $toolkitVersion
    $readme.Range("A5").Value = "구현 위치"
    $readme.Range("B5").Value = "src\MacroToolkit.bas"
    $readme.Range("A6").Value = "이 통합문서에서 확인할 곳"
    $readme.Range("B6").Value = "VBA 소스 미리보기 시트에서 src\MacroToolkit.bas 내용을 그대로 확인할 수 있습니다."
    $readme.Range("A7").Value = "매크로 실행 방법"
    $readme.Range("B7").Value = "VBA 편집기에서 src\MacroToolkit.bas를 가져오거나, Excel의 'VBA 프로젝트 개체 모델 접근 신뢰' 옵션을 켠 뒤 다시 빌드하세요."
    $readme.Range("A9").Value = "주의"
    $readme.Range("B9").Value = "Excel 보안 설정이 자동 삽입을 막으면 이 파일에는 실행 가능한 VBA 모듈이 없을 수 있습니다. 이 경우 소스 미리보기와 .bas 파일이 기준입니다."
    $readme.Range("A1:B1").Font.Bold = $true
    $readme.Range("A4:A9").Font.Bold = $true
    $readme.Columns("A").ColumnWidth = 22
    $readme.Columns("B").ColumnWidth = 110
    $readme.Columns("B").WrapText = $true

    $list = Add-SheetAfter $workbook "매크로 목록"
    $headers = @("매크로", "기능", "입력")
    for ($i = 0; $i -lt $headers.Count; $i++) { $list.Cells.Item(1, $i + 1).Value2 = [string]$headers[$i] }
    $rows = @(
        @("Toolkit_About", "툴킷 버전과 주요 기능을 보여줍니다.", "없음"),
        @("Toolkit_SetupWorkbook", "숨김 버전 정보 시트를 생성/갱신합니다.", "없음"),
        @("ConvertTextNumbersInSelection", "선택한 셀의 텍스트 숫자를 실제 숫자로 변환합니다.", "현재 선택 영역"),
        @("ConvertTextNumbersInUsedRange", "활성 시트의 사용 범위 전체에서 텍스트 숫자를 변환합니다.", "활성 시트"),
        @("ConvertTextNumbersInColumns", "마우스로 지정한 범위의 텍스트 숫자를 변환합니다.", "범위 선택, 예: A:A / B1:B100"),
        @("NormalizePhonesInSelection", "선택한 셀의 한국 전화번호 형식을 정규화합니다.", "현재 선택 영역"),
        @("NormalizePhonesInColumns", "마우스로 지정한 범위의 한국 전화번호 형식을 정규화합니다.", "범위 선택, 예: B:B / C1:C100"),
        @("OutlierDetect_Setup", "이상치 탐지 설정 시트를 생성하고 드롭다운을 준비합니다.", "현재 통합문서의 데이터 시트"),
        @("OutlierDetect_Run", "설정값에 따라 이상치를 탐지하고 대체 결과 시트를 생성합니다.", "_이상치탐지_설정 시트"),
        @("SPSS_DeleteSelectedBanner", "선택한 SPSS 가로/세로 배너 블록을 삭제하고 남은 열/행을 당겨 붙입니다.", "선택한 ◐ 또는 ▣ 배너 셀")
    )
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[$r].Count; $c++) {
            $list.Cells.Item($r + 2, $c + 1).Value2 = [string]$rows[$r][$c]
        }
    }
    $list.Range("A1:C1").Font.Bold = $true
    $list.Columns("A:C").AutoFit()

    $source = Add-SheetAfter $workbook "VBA 소스 미리보기"
    $source.Range("A1").Value = "줄"
    $source.Range("B1").Value = "src\MacroToolkit.bas 코드"
    $source.Range("A1:B1").Font.Bold = $true
    $lines = Get-Content -Encoding UTF8 -Path $ModulePath
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $source.Cells.Item($i + 2, 1).Value2 = [string]($i + 1)
        $source.Cells.Item($i + 2, 2).Value2 = [string]$lines[$i]
    }
    $source.Columns("A").ColumnWidth = 8
    $source.Columns("B").ColumnWidth = 140
    $source.Columns("B").Font.Name = "Consolas"
    $source.Columns("B").Font.Size = 9
    $source.Range("A2").Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    $readme.Activate() | Out-Null
    $workbook.Save()
    Write-Output "Updated workbook documentation sheets: $WorkbookPath"
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($true) | Out-Null } catch {}
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}


