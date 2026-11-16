param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$tokens = $null; $errors = $null
$path = Join-Path $ProjectRoot 'scripts\publish_release.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Publisher parse failed' }
$fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Add-SpssRebuildUserForm' }, $true)
if (-not $fn) { throw 'SPSS form builder missing' }
Invoke-Expression $fn.Extent.Text
$excel = New-Object -ComObject Excel.Application
$workbook = $null
try {
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    Add-SpssRebuildUserForm -Workbook $workbook
    $form = $workbook.VBProject.VBComponents.Item('frmSpssRebuildOptions')
    $form.CodeModule.AddFromString(@'
Public Function VerifyBannerActions() As String
    On Error GoTo Failed
    Dim tables As New Collection, banners As New Collection
    tables.Add Array("Table", 1, 2, "Sheet")
    banners.Add "Alpha": banners.Add "Beta": banners.Add "Gamma"
    InitializeOptions tables, banners
    lstBanners.Selected(0) = True
    cmdBannerDown_Click
    cmdOK_Click
    If Len(DeleteBannersText) <> 0 Then Err.Raise 513, , "Moving selected banner marked it for deletion"
    If BannerOrderText <> "2,1,3" Then Err.Raise 513, , "Banner move order incorrect"
    cmdToggleBannerDelete_Click
    cmdOK_Click
    If DeleteBannersText <> "1" Then Err.Raise 513, , "Explicit deletion not recorded"
    cmdToggleBannerDelete_Click
    cmdOK_Click
    If Len(DeleteBannersText) <> 0 Then Err.Raise 513, , "Restored banner still deleted"
    InitializeOptions tables, banners
    lstBanners.Selected(0) = True: lstBanners.Selected(1) = True
    cmdBannerDown_Click
    cmdOK_Click
    If BannerOrderText <> "3,1,2" Or Len(DeleteBannersText) <> 0 Then Err.Raise 513, , "Multi-select banner move incorrect"
    cmdToggleBannerDelete_Click
    cmdOK_Click
    If DeleteBannersText <> "1,2" Then Err.Raise 513, , "Multi-select banner deletion incorrect"
    VerifyBannerActions = "PASS"
    Exit Function
Failed:
    VerifyBannerActions = CStr(Err.Number) & ": " & Err.Description
End Function
'@)
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function RunBannerActions() As String
    Dim f As Object
    Set f = VBA.UserForms.Add("frmSpssRebuildOptions")
    RunBannerActions = f.VerifyBannerActions()
    Unload f
End Function
'@)
    $result = $excel.Run("'$($workbook.Name)'!RunBannerActions")
    if ($result -ne 'PASS') { throw "SPSS form test failed: $result" }
    Write-Host 'SPSS banner move, explicit delete and restore passed.'
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
