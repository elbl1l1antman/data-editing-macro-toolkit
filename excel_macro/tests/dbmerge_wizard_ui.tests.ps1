param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$publishPath = Join-Path $ProjectRoot "scripts\publish_release.ps1"
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($publishPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }

$functionAst = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq "Add-DatabaseMergeSetupUserForm"
}, $true) | Select-Object -First 1
if ($null -eq $functionAst) { throw "Add-DatabaseMergeSetupUserForm not found" }
Invoke-Expression $functionAst.Extent.Text

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Add()

    Add-DatabaseMergeSetupUserForm -Workbook $workbook
    $component = $workbook.VBProject.VBComponents.Item("frmDatabaseMergeSetup")
    $designer = $component.Designer
    try {
        $multiPage = $designer.Controls.Item("mpSteps")
    }
    catch {
        throw "Wizard MultiPage control was not generated"
    }

    if ($multiPage.Pages.Count -ne 3) { throw "Expected 3 wizard pages, got $($multiPage.Pages.Count)" }
    $expectedCaptions = @("1. 통합할 시트 선택", "2. 가져올 열과 ID 설정", "3. 통합 결과 확인")
    for ($i = 0; $i -lt 3; $i++) {
        if ([string]$multiPage.Pages.Item($i).Caption -ne $expectedCaptions[$i]) {
            throw "Unexpected page $i caption: $($multiPage.Pages.Item($i).Caption)"
        }
    }

    foreach ($controlName in @("lstSheets", "lstHeaders", "optAutoId", "optExistingId", "optManualId", "fraAdvanced", "txtReviewSummary", "txtReviewMapping", "cmdBack", "cmdNext", "cmdRun", "cmdCancel")) {
        $found = $false
        try {
            $null = $designer.Controls.Item($controlName)
            $found = $true
        }
        catch {}
        for ($pageIndex = 0; $pageIndex -lt $multiPage.Pages.Count; $pageIndex++) {
            if ($found) { break }
            try {
                $null = $multiPage.Pages.Item($pageIndex).Controls.Item($controlName)
                $found = $true
                break
            }
            catch {}
        }
        if (-not $found) { throw "Wizard control missing: $controlName" }
    }

    for ($pageIndex = 0; $pageIndex -lt $multiPage.Pages.Count; $pageIndex++) {
        $page = $multiPage.Pages.Item($pageIndex)
        foreach ($control in $page.Controls) {
            if (($control.Left + $control.Width) -gt ($multiPage.Width - 12)) {
                throw "Control exceeds wizard width: $($control.Name)"
            }
            if (($control.Top + $control.Height) -gt ($multiPage.Height - 22)) {
                throw "Control exceeds wizard height: $($control.Name)"
            }
        }
    }

    $wrapper = $workbook.VBProject.VBComponents.Add(1)
    $wrapper.CodeModule.AddFromString(@'
Public Function DBMergeWizardFormPassed() As String
    Dim form As Object
    Set form = VBA.UserForms.Add("frmDatabaseMergeSetup")
    form.Configure "서울" & vbTab & "1" & vbTab & "2" & vbTab & "선택" & vbLf & _
                   "_내부자료" & vbTab & "0" & vbTab & "0" & vbTab & "확인 필요", _
                   "서울", "ID", "통합데이터", "맨 뒤", "자동", "single", "output_auto", "header", False
    If form.SelectedSheets <> "서울" Then
        DBMergeWizardFormPassed = "selected sheets not preserved"
        Exit Function
    End If
    form.ConfigureHeaders "ID" & vbTab & "1" & vbTab & "전체 시트" & vbLf & _
                          "사업체명" & vbTab & "1" & vbTab & "전체 시트" & vbLf & _
                          "매출액" & vbTab & "0" & vbTab & "일부 시트만 있음"
    If InStr(1, form.SelectedHeaders, "사업체명", vbTextCompare) = 0 Then
        DBMergeWizardFormPassed = "common header not selected"
        Exit Function
    End If
    form.ConfigureReview "요약", "매칭", "경고 없음"
    If form.Action <> "review" Then
        DBMergeWizardFormPassed = "review state not applied"
        Exit Function
    End If
    Unload form
    DBMergeWizardFormPassed = "PASS"
End Function
'@)
    $result = [string]$excel.Run("DBMergeWizardFormPassed")
    if ($result -ne "PASS") { throw $result }

    Write-Host "DBMerge wizard UI tests passed."
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
