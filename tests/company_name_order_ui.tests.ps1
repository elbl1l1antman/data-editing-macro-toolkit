param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot "scripts\publish_release.ps1") -Raw -Encoding UTF8
foreach ($token in @("RefreshOrderNumbers", "ClearOrderNumbers", 'ctl.Caption = CStr(orderNo) & ". " & ctl.Caption')) {
    if (-not $source.Contains($token)) { throw "Company name order UI token missing: $token" }
}

Write-Host "Company name order UI tests passed."

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw $parseErrors[0] }
$builder = $ast.Find({ param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Add-CompanyNameOptionsUserForm'
}, $true)
if (-not $builder) { throw 'Company name form builder missing.' }
Invoke-Expression $builder.Extent.Text

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Add()
    Add-CompanyNameOptionsUserForm -Workbook $workbook
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function TestCompanyOptionsRuntime() As String
    On Error GoTo Failed
    Dim form As Object
    Set form = VBA.UserForms.Add("frmCompanyNameOptions")
    If Not form.Cancelled Then Err.Raise 513, , "Default must be cancelled"
    If form.RemoveLocationUnits Then Err.Raise 513, , "Location cleanup must be opt-in"
    If form.PreserveEstablishment Then Err.Raise 513, , "Conservative profile must be opt-in"
    form.Controls("chkPreserveEstablishment").Value = True
    If Not form.PreserveEstablishment Then Err.Raise 513, , "Conservative profile not returned"
    Dim tip As Object
    Set tip = form.Controls("lblTip")
    If InStr(tip.Caption, "추천 순서: (주) 삭제 → 지명까지 함께 제거 → 사업장 표기 제거") = 0 Then Err.Raise 513, , "Recommended cleanup order missing"
    If InStr(tip.Caption, "엘지이노텍(주)구미1공장") = 0 Then Err.Raise 513, , "User example missing"
    If InStr(tip.Caption, "숫자가 붙은 지명은 결과 확인이 필요") = 0 Then Err.Raise 513, , "Numbered location limitation missing"
    If Not tip.WordWrap Or tip.Height < 90 Then Err.Raise 513, , "Recommendation text area too small"
    If tip.Top + tip.Height > form.Controls("cmdOK").Top Then Err.Raise 513, , "Recommendation overlaps execute button"
    If form.Controls("chkRemoveSpaces").Caption <> "모든 공백 제거" Then Err.Raise 513, , "Plain caption changed"
    form.Controls("chkTrimOuter").Value = True
    If Left$(form.Controls("chkTrimOuter").Caption, 3) <> "5. " Then Err.Raise 513, , "Click order missing"
    form.Controls("chkTrimOuter").Value = False
    If form.Controls("chkTrimOuter").Caption <> "앞뒤 공백 제거" Then Err.Raise 513, , "Caption not restored"
    form.Controls("cboOrderPreset").ListIndex = 0
    If form.Controls("chkRemoveSpaces").Caption <> "모든 공백 제거" Then Err.Raise 513, , "Preset caption not restored"
    form.Controls("chkRemoveLocationUnits").Value = True
    If Not form.RemoveLocationUnits Then Err.Raise 513, , "Location option not returned"
    If InStr(form.OperationOrder, "removeLocationUnits") = 0 Then Err.Raise 513, , "Location step missing"
    form.Controls("chkRemoveBranchMarkers").Value = True
    If Not form.RemoveBranchMarkers Then Err.Raise 513, , "Business-unit option not returned"
    If InStr(form.OperationOrder, "removeBranchMarkers") = 0 Then Err.Raise 513, , "Business-unit step missing"
    Dim ctl As Object
    For Each ctl In form.Controls
        If ctl.Left + ctl.Width > form.InsideWidth Then Err.Raise 513, , "Clipped horizontally: " & ctl.Name
        If ctl.Top + ctl.Height > form.InsideHeight Then Err.Raise 513, , "Clipped vertically: " & ctl.Name
    Next ctl
    Unload form
    TestCompanyOptionsRuntime = "PASS"
    Exit Function
Failed:
    TestCompanyOptionsRuntime = CStr(Err.Number) & ": " & Err.Description
End Function
'@)
    $result = $excel.Run("'$($workbook.Name)'!TestCompanyOptionsRuntime")
    if ($result -ne 'PASS') { throw "Company options runtime failed: $result" }
    Write-Host 'Company name options initialize and click-order runtime tests passed.'
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
