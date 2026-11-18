param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$agePath = Join-Path $ProjectRoot 'src\KosisAge.bas'
if (-not (Test-Path -LiteralPath $agePath)) { throw 'KOSIS age aggregation module missing.' }
$age = (Get-Content -LiteralPath $agePath -Raw -Encoding UTF8) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
if (-not $age.Contains('Public Function Kosis_AggregateAges(')) { throw 'KOSIS age aggregation missing.' }
$age = $age -replace 'Option Explicit', "Option Explicit`r`nPublic Const BIZ_APP_NAME As String = `"TestOnly`"`r`nPublic Const TOOLKIT_CANCEL_ERROR_NUMBER As Long = 12345"
$core = (Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\KosisTools.bas') -Raw -Encoding UTF8) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$core = $core -replace '(?m)^Option Explicit\s*', ''
$test = @'
Public Sub Toolkit_RaiseIfCancelled()
End Sub
Public Sub Toolkit_ProgressUpdate(ByVal current As Long, ByVal total As Long, ByVal message As String)
End Sub
Public Function TestAgeBands() As String
    On Error GoTo Failed
    Dim spec As String, bands As Collection, uncovered As Long, row As Object, n As Long
    Dim metadata As New Collection, records As New Collection, regions As New Collection, review As Collection, result As Collection, codes As Collection
    spec = "1" & vbTab & "0" & vbTab & ChrW(&HC774) & ChrW(&HC0C1) & vbTab & "20" & vbTab & ChrW(&HBBF8) & ChrW(&HB9CC) & vbTab & "young" & vbLf & _
           "2" & vbTab & "20" & vbTab & ChrW(&HC774) & ChrW(&HC0C1) & vbTab & "" & vbTab & ChrW(&HBBF8) & ChrW(&HB9CC) & vbTab & "adult"
    Set bands = Kosis_ParseAgeBands(spec, 100, uncovered)
    Set row = bands(1): If row("from") <> 0 Or row("to") <> 19 Then Err.Raise 5, , "first bounds wrong"
    Set row = bands(2): If row("from") <> 20 Or row("to") <> 100 Or uncovered <> 0 Then Err.Raise 5, , "tail wrong"
    spec = "1" & vbTab & "0" & vbTab & ChrW(&HC774) & ChrW(&HC0C1) & vbTab & "100" & vbTab & ChrW(&HC774) & ChrW(&HD558) & vbTab & "split"
    On Error Resume Next
    Set bands = Kosis_ParseAgeBands(spec, 100, uncovered): n = Err.Number: Err.Clear
    On Error GoTo Failed
    If n = 0 Then Err.Raise 5, , "open ended category was split"
    Set bands = New Collection
    Set row = CreateObject("Scripting.Dictionary"): row.Add "name", "early": row.Add "from", 0: row.Add "to", 1: bands.Add row
    Set row = CreateObject("Scripting.Dictionary"): row.Add "name", "2" & ChrW(&HC138) & " " & ChrW(&HC774) & ChrW(&HC0C1): row.Add "from", 2: row.Add "to", 2: bands.Add row
    Set row = CreateObject("Scripting.Dictionary"): row.Add "axis", "YRE": row.Add "code", "01": row.Add "name", "0" & ChrW(&HC138): metadata.Add row
    Set row = CreateObject("Scripting.Dictionary"): row.Add "axis", "YRE": row.Add "code", "02": row.Add "name", "1" & ChrW(&HC138): metadata.Add row
    Set row = CreateObject("Scripting.Dictionary"): row.Add "axis", "YRE": row.Add "code", "03": row.Add "name", "2" & ChrW(&HC138) & " " & ChrW(&HC774) & ChrW(&HC0C1): metadata.Add row
    If Kosis_FindTopAge(metadata) <> 2 Then Err.Raise 5, , "top age not detected"
    Set codes = Kosis_AgeSourceCodes(metadata, bands, 2)
    If codes.Count <> 3 Or codes(1) <> "01" Or codes(3) <> "03" Then Err.Raise 5, , "age codes wrong"
    Set row = CreateObject("Scripting.Dictionary"): row.Add "C_A", "00": row.Add "C_YRE", "01": row.Add "ITEM", "T2": row.Add "period", "2024": row.Add "value", "1": records.Add row
    Set row = CreateObject("Scripting.Dictionary"): row.Add "C_A", "00": row.Add "C_YRE", "02": row.Add "ITEM", "T2": row.Add "period", "2024": row.Add "value", "2": records.Add row
    Set row = CreateObject("Scripting.Dictionary"): row.Add "C_A", "00": row.Add "C_YRE", "03": row.Add "ITEM", "T2": row.Add "period", "2024": row.Add "value", "X": records.Add row
    regions.Add "00"
    Set result = Kosis_AggregateAges(records, metadata, bands, regions, "2024", "T2", 2, review)
    Set row = result(1): If row("value") <> "3" Then Err.Raise 5, , "age sum mismatch"
    Set row = result(2): If IsNumeric(row("value")) Or review.Count <> 1 Then Err.Raise 5, , "suppressed age sum treated as numeric"
    records.Remove 1
    Set result = Kosis_AggregateAges(records, metadata, bands, regions, "2024", "T2", 2, review)
    Set row = result(1): If IsNumeric(row("value")) Or review.Count <> 2 Then Err.Raise 5, , "missing age sum treated as complete"
    TestAgeBands = "PASS": Exit Function
Failed:
    TestAgeBands = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$core = $core -replace '(?m)^Public Function Kosis_ValidateAgeBands', 'Public Function Kosis_ValidateAgeBands'
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($age + "`r`n" + $test)
    $helper = $book.VBProject.VBComponents.Add(1)
    $helper.CodeModule.AddFromString("Option Explicit`r`n" + $core)
    $result = [string]$excel.Run("'$($book.Name)'!TestAgeBands")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'KOSIS age band tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
