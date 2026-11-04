param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)
. (Join-Path $PSScriptRoot "test_helpers.ps1")

$ErrorActionPreference = "Stop"
$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$fixturePath = Join-Path $ProjectRoot "dev\dbmerge_wizard_test.xlsx"
$source = Get-ToolkitVbaSource -ProjectRoot $ProjectRoot

function Get-VbaBlock {
    param([string]$SourceText, [string]$ProcedureName)

    $start = [regex]::Match($SourceText, "(?m)^(Public|Private)\s+(Function|Sub)\s+$([regex]::Escape($ProcedureName))\b.*$")
    if (-not $start.Success) { throw "Procedure not found: $ProcedureName" }
    $endToken = if ($start.Groups[2].Value -eq "Function") { "End Function" } else { "End Sub" }
    $tail = $SourceText.Substring($start.Index)
    $end = [regex]::Match($tail, "(?m)^$endToken\s*$")
    if (-not $end.Success) { throw "Procedure end not found: $ProcedureName" }
    $tail.Substring(0, $end.Index + $end.Length)
}

function Assert-Contains([string]$Actual, [string]$Expected) {
    if (-not $Actual.Contains($Expected)) { throw "Expected text not found: $Expected`nActual:`n$Actual" }
}

function Assert-NotContains([string]$Actual, [string]$Unexpected) {
    if ($Actual.Contains($Unexpected)) { throw "Unexpected text found: $Unexpected`nActual:`n$Actual" }
}

if (-not (Test-Path -LiteralPath $fixturePath)) {
    throw "Fixture not found: $fixturePath"
}

$procedureNames = @(
    "DBMerge_WizardSheetInfoText",
    "DBMerge_WizardDefaultSheetNames",
    "DBMerge_WizardHeaderInfoText",
    "DBMerge_WizardDuplicateHeaderCount",
    "DBMerge_WizardIsDefaultSheet",
    "DBMerge_WizardSelectedSheetCount",
    "DBMerge_WizardManualKeyColumns",
    "DBMerge_WizardWarningSummary",
    "DBMerge_NameIsSelected",
    "DBMerge_HeaderRowForSheet",
    "DBMerge_RowLooksLikeMarker",
    "DBMerge_DetectHeaderRow",
    "DBMerge_HeaderModeIsUsable",
    "DBMerge_DataStartRow",
    "DBMerge_HeaderMap",
    "DBMerge_AddHeaderMapRow",
    "DBMerge_HeaderTextForColumn",
    "DBMerge_RowTextCount",
    "DBMerge_LastUsedRow",
    "DBMerge_LastUsedCol",
    "DBMerge_NormalizeHeader"
)

$moduleCode = "Option Explicit`r`n" + (($procedureNames | ForEach-Object {
    Get-VbaBlock -SourceText $source -ProcedureName $_
}) -join "`r`n`r`n")

$wrapperCode = @'
Public Function DBMergeWizardControllerRulesPassed(ByVal wb As Workbook) As String
    Dim columns As Object
    Set columns = DBMerge_WizardManualKeyColumns(wb, _
        "서울" & vbTab & "ID" & vbLf & "부산" & vbTab & "관리번호", "자동", "single")
    If Not columns.Exists("서울") Or CLng(columns("서울")) <> 1 Then
        DBMergeWizardControllerRulesPassed = "서울 ID 열 매칭 실패"
        Exit Function
    End If
    If Not columns.Exists("부산") Or CLng(columns("부산")) <> 1 Then
        DBMergeWizardControllerRulesPassed = "부산 관리번호 열 매칭 실패"
        Exit Function
    End If

    Dim plan As Object
    Set plan = CreateObject("Scripting.Dictionary")
    plan.Add "idMode", "source_auto"
    Dim headers As Collection
    Set headers = New Collection
    headers.Add "ID"
    headers.Add "사업체명"
    headers.Add "주소"
    plan.Add "headers", headers

    Dim sheetPlans As Collection
    Set sheetPlans = New Collection
    Dim sheetPlan As Object
    Set sheetPlan = CreateObject("Scripting.Dictionary")
    sheetPlan.Add "sheetName", "서울"
    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.Add "id", 1
    headerMap.Add "사업체명", 2
    sheetPlan.Add "headerMap", headerMap
    sheetPlans.Add sheetPlan
    plan.Add "sheets", sheetPlans

    Dim warnings As String
    warnings = DBMerge_WizardWarningSummary(plan)
    If InStr(1, warnings, "일부 누락", vbTextCompare) = 0 Then
        DBMergeWizardControllerRulesPassed = "누락 헤더 경고 없음"
        Exit Function
    End If
    If InStr(1, warnings, "원본 변경", vbTextCompare) = 0 Then
        DBMergeWizardControllerRulesPassed = "원본 변경 경고 없음"
        Exit Function
    End If
    DBMergeWizardControllerRulesPassed = "PASS"
End Function
'@

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($fixturePath)
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = "DBMergeWizardTests"
    $module.CodeModule.AddFromString($moduleCode)
    $wrapper = $workbook.VBProject.VBComponents.Add(1)
    $wrapper.CodeModule.AddFromString($wrapperCode)

    $sheetInfo = [string]$excel.Run("DBMerge_WizardSheetInfoText", $workbook)
    $defaultSheets = [string]$excel.Run("DBMerge_WizardDefaultSheetNames", $workbook)
    $selected = "서울`n부산`nID없음"
    $headerInfo = [string]$excel.Run("DBMerge_WizardHeaderInfoText", $workbook, $selected, "자동", "single")

    Assert-Contains $sheetInfo "서울`t1`t2`t선택"
    Assert-Contains $sheetInfo "_내부자료"
    Assert-Contains $defaultSheets "서울"
    Assert-NotContains $defaultSheets "_내부자료"
    Assert-Contains $headerInfo "사업체명`t3`t전체 시트"
    Assert-Contains $headerInfo "매출액`t2`t일부 시트만 있음"

    $bad = $workbook.Worksheets.Add()
    $bad.Name = "오류헤더"
    $bad.Cells.Item(1, 1).Value2 = "한개"
    $duplicate = $workbook.Worksheets.Add()
    $duplicate.Name = "중복헤더"
    $duplicate.Cells.Item(1, 1).Value2 = "사업체명"
    $duplicate.Cells.Item(1, 2).Value2 = "사업 체명"
    $duplicate.Cells.Item(1, 3).Value2 = "주소"
    $duplicate.Cells.Item(2, 1).Value2 = "가나다"

    $warningInfo = [string]$excel.Run("DBMerge_WizardSheetInfoText", $workbook)
    Assert-Contains $warningInfo "오류헤더`t0`t0`t확인 필요"
    Assert-Contains $warningInfo "중복헤더`t1`t1`t선택`t중복 헤더"

    $controllerResult = [string]$excel.Run("DBMergeWizardControllerRulesPassed", $workbook)
    if ($controllerResult -ne "PASS") { throw $controllerResult }

    Write-Host "DBMerge wizard logic tests passed."
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
