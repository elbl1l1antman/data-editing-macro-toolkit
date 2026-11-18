param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$wizard = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\KosisWizard.bas') -Raw -Encoding UTF8
if ($wizard -match '(?m)^\s*Toolkit_ProgressClose\s*:') { throw 'Progress close is parsed as a VBA label.' }
$publish = Get-Content -LiteralPath (Join-Path $ProjectRoot 'scripts\publish_release.ps1') -Raw -Encoding UTF8
$match = [regex]::Match($publish, '(?ms)^function Add-KosisPickerUserForm \{.*?^\}\s*(?=function )')
if (-not $match.Success) { throw 'Picker builder missing.' }
Invoke-Expression $match.Value
$test = @'
Option Explicit
Public Function TestPicker() As String
    On Error GoTo Failed
    Dim items As New Collection, row As Object, form As Object
    Set row = CreateObject("Scripting.Dictionary")
    row.Add "axis", "B": row.Add "code", "01": row.Add "name", "Agriculture": row.Add "parent", ""
    items.Add row
    Set row = CreateObject("Scripting.Dictionary")
    row.Add "axis", "B": row.Add "code", "011": row.Add "name", "Crops": row.Add "parent", "01"
    items.Add row
    Set form = VBA.UserForms.Add("frmKosisPicker")
    form.Configure "Test", "Select", items, "B", "01"
    If form.Controls("lstItems").ListCount <> 1 Then Err.Raise 5, , "wrong initial count"
    If CStr(form.Controls("lstItems").List(0, 0)) <> "01" Then Err.Raise 5, , "code lost"
    If Not form.Controls("lstItems").Selected(0) Then Err.Raise 5, , "default lost"
    form.Controls("lstItems").Selected(0) = False
    If InStr(CStr(form.Controls("lblCount").Caption), "0") = 0 Then Err.Raise 5, , "selection change not reflected"
    form.Controls("lstItems").Selected(0) = True
    If InStr(CStr(form.Controls("lblCount").Caption), "1") = 0 Then Err.Raise 5, , "selection change not reflected"
    form.Controls("cmdExpand").Value = True
    If form.Controls("lstItems").ListCount <> 2 Then Err.Raise 5, , "expand failed"
    form.Controls("cmdChildren").Value = True
    If Not form.Controls("lstItems").Selected(1) Then Err.Raise 5, , "child selection failed"
    form.Controls("txtSearch").Text = "Crops"
    form.Controls("cmdSearch").Value = True
    If form.Controls("lstItems").ListCount <> 1 Or Not form.Controls("lstItems").Selected(0) Then Err.Raise 5, , "search selection failed"
    form.Controls("cmdOK").Value = True
    If Not form.Accepted Or form.SelectedKeys <> "01" & vbLf & "011" Then Err.Raise 5, , "hidden selection lost"
    Unload form
    Set form = VBA.UserForms.Add("frmKosisPicker")
    form.Configure "Test", "Select", items, "B", "01"
    form.Controls("cmdBack").Value = True
    If Not form.BackRequested Or form.Accepted Then Err.Raise 5, , "back was not distinguished from acceptance"
    Unload form
    TestPicker = "PASS"
    Exit Function
Failed:
    TestPicker = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    Add-KosisPickerUserForm -Workbook $book
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($test)
    $result = [string]$excel.Run("'$($book.Name)'!TestPicker")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'KOSIS picker form configuration passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
