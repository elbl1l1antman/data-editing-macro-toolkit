param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$publishPath = Join-Path $ProjectRoot "scripts\publish_release.ps1"
$modulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$publishSource = Get-Content -LiteralPath $publishPath -Raw -Encoding UTF8
$moduleSource = Get-Content -LiteralPath $modulePath -Raw -Encoding UTF8

$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($publishPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$fn = $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Add-NaverLocalVerifyOptionsUserForm" }, $true) | Select-Object -First 1
if ($null -eq $fn) { throw "Add-NaverLocalVerifyOptionsUserForm not found" }
$formSource = $fn.Extent.Text

foreach ($removed in @("txtNameRange", "txtAddressRange", "cmdPickName", "cmdPickAddress", "cmbMode")) {
    if ($formSource.Contains($removed)) { throw "Obsolete Naver verify control remains: $removed" }
}
foreach ($required in @(
    'chkUseName.Caption = "업체명 범위 사용"',
    'chkUseAddress.Caption = "주소 범위 사용"',
    '실행을 누르면 업체명, 주소 순서로 Excel 범위 선택창이 열립니다.',
    '실행 후 업체명 범위 선택',
    '실행 후 주소 범위 선택'
)) {
    if (-not $formSource.Contains($required)) { throw "Naver verify guidance missing: $required" }
}
if ($formSource -match "Application\.InputBox") {
    throw "Naver verify UserForm must close before Excel range selection starts"
}
foreach ($required in @(
    'Toolkit_GetRangeFromUserForm("네이버 주소 검증 - 업체명 범위"',
    'Toolkit_GetRangeFromUserForm("네이버 주소 검증 - 주소 범위"'
)) {
    if (-not $moduleSource.Contains($required)) { throw "Post-form range picker missing: $required" }
}

Invoke-Expression $formSource
$excel = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()
    Add-NaverLocalVerifyOptionsUserForm -Workbook $workbook
    $component = $workbook.VBProject.VBComponents.Item("frmNaverLocalVerifyOptions")
    $designer = $component.Designer

    $names = @($designer.Controls | ForEach-Object { [string]$_.Name })
    foreach ($required in @("chkUseName", "lblNameRangeStep", "chkUseAddress", "lblAddressRangeStep", "cmdOK", "cmdCancel")) {
        if ($names -notcontains $required) { throw "Required Naver verify control missing: $required" }
    }
    foreach ($removed in @("txtNameRange", "txtAddressRange", "cmdPickName", "cmdPickAddress", "cmbMode")) {
        if ($names -contains $removed) { throw "Obsolete Naver verify control created: $removed" }
    }

    $formWidth = [double]$component.Properties.Item("Width").Value
    $formHeight = [double]$component.Properties.Item("Height").Value
    foreach ($control in $designer.Controls) {
        if (($control.Left + $control.Width) -gt ($formWidth - 8)) { throw "Control clipped horizontally: $($control.Name)" }
        if (($control.Top + $control.Height) -gt ($formHeight - 18)) { throw "Control clipped vertically: $($control.Name)" }
    }

    $wrapper = $workbook.VBProject.VBComponents.Add(1)
    $wrapper.CodeModule.AddFromString(@'
Public Function NaverVerifyFormPassed() As String
    Dim form As Object
    Set form = VBA.UserForms.Add("frmNaverLocalVerifyOptions")
    If form.QueryMode <> "name_address" Then
        NaverVerifyFormPassed = "default query mode mismatch"
    ElseIf Not form.UseName Or Not form.UseAddress Then
        NaverVerifyFormPassed = "default range choices mismatch"
    Else
        NaverVerifyFormPassed = "PASS"
    End If
    Unload form
End Function
'@)
    $result = [string]$excel.Run("NaverVerifyFormPassed")
    if ($result -ne "PASS") { throw $result }

    Write-Host "Naver verify range picker tests passed."
}
finally {
    if ($workbook -ne $null) { $workbook.Close($false) | Out-Null }
    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
