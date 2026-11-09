param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path)
$ErrorActionPreference = 'Stop'
$key = [Environment]::GetEnvironmentVariable('OPENDART_TEST_KEY', 'Process')
if (-not $key) { Write-Host 'DART live test skipped: OPENDART_TEST_KEY not set.'; exit 0 }
$excel = New-Object -ComObject Excel.Application
$workbook = $null; $tempPath = $null
try {
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Const BIZ_APP_NAME As String = "DataEditingMacroToolkit"
Public Const TOOLKIT_CANCEL_ERROR_NUMBER As Long = vbObjectError + 9130
Public Function Toolkit_GetRangeFromUserForm(a As String, b As String, ByRef c As Range) As Boolean
    Set c = ActiveSheet.Range("A1:A2")
    Toolkit_GetRangeFromUserForm = True
End Function
Public Sub Dart_TestMessage(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant)
End Sub
Public Sub Toolkit_ProgressShow(a As String, b As String, c As Long)
End Sub
Public Sub Toolkit_ProgressUpdate(a As Long, b As Long, c As String)
End Sub
Public Sub Toolkit_ProgressClose()
End Sub
Public Sub Toolkit_RaiseIfCancelled()
End Sub
Public Function CheckDartLive(ByVal key As String) As String
    On Error GoTo Failed
    Dim names As Object, codes As Object, stocks As Object, fields As Variant
    Set names = CreateObject("Scripting.Dictionary")
    Set codes = CreateObject("Scripting.Dictionary")
    Set stocks = CreateObject("Scripting.Dictionary")
    Dart_LoadCompanyMaster key, names, codes, stocks
    If Not codes.Exists("00126380") Then Err.Raise 513, , "Known company not found in master"
    fields = Dart_QueryCompany(key, "00126380", 2023)
    If Len(CStr(fields(1))) = 0 Then Err.Raise 513, , "Company information missing"
    If Len(CStr(fields(13))) = 0 Or Len(CStr(fields(17))) = 0 Then Err.Raise 513, , "Revenue or employee count missing"
    CheckDartLive = "PASS"
    Exit Function
Failed:
    CheckDartLive = CStr(Err.Number) & ": " & Replace(Err.Description, key, "[redacted]")
End Function
Public Function CheckDartEndToEnd() As String
    On Error GoTo Failed
    Dim ws As Worksheet, result As Worksheet, oldCount As Long
    Set ws = ActiveSheet
    ws.Range("A1:A2").NumberFormat = "@"
    ws.Range("A1").Value2 = "00126380"
    ws.Range("A2").Value2 = "NONEXISTENT-TEST"
    oldCount = ActiveWorkbook.Worksheets.Count
    Dart_SearchCompanies
    If ActiveWorkbook.Worksheets.Count <> oldCount + 1 Then Err.Raise 513, , "Result sheet not created"
    Set result = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    If result.Range("C2").Value2 <> "00126380" Then Err.Raise 513, , "Result code mismatch"
    If Len(CStr(result.Range("D2").Value2)) = 0 Or Len(CStr(result.Range("P2").Value2)) = 0 Or Len(CStr(result.Range("T2").Value2)) = 0 Then Err.Raise 513, , "Result columns missing"
    If Len(CStr(result.Range("U3").Value2)) = 0 Then Err.Raise 513, , "Unmatched row lacks warning"
    If ws.Range("A1").Value2 <> "00126380" Then Err.Raise 513, , "Original data changed"
    CheckDartEndToEnd = "PASS"
    Exit Function
Failed:
    CheckDartEndToEnd = CStr(Err.Number) & ": " & Err.Description
End Function
'@)
    $tempPath = Join-Path $env:TEMP ('DartTools-' + [guid]::NewGuid().ToString('N') + '.bas')
    $source = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot 'src\DartTools.bas')
    $source = [regex]::Replace($source, '(?m)^    key = GetSetting\([^\r\n]+', '    key = Environ$("OPENDART_TEST_KEY")')
    $source = [regex]::Replace($source, '(?m)^    headerChoice = MsgBox\([^\r\n]+', '    headerChoice = vbNo')
    $source = [regex]::Replace($source, '(?m)^    yearValue = Application\.InputBox\([^\r\n]+', '    yearValue = 2023')
    $source = [regex]::Replace($source, '(?m)^(\s*)MsgBox ', '$1Dart_TestMessage ')
    $encoding = [Text.Encoding]::GetEncoding(949, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
    [IO.File]::WriteAllText($tempPath, $source, $encoding)
    $workbook.VBProject.VBComponents.Import($tempPath) | Out-Null
    $result = $excel.Run("'$($workbook.Name)'!CheckDartLive", $key)
    if ($result -ne 'PASS') { throw "DART live test failed: $result" }
    $result = $excel.Run("'$($workbook.Name)'!CheckDartEndToEnd")
    if ($result -ne 'PASS') { throw "DART end-to-end test failed: $result" }
    Write-Host 'DART master, company, revenue, employee and result-sheet tests passed.'
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    if ($tempPath -and (Test-Path -LiteralPath $tempPath)) { Remove-Item -LiteralPath $tempPath -Force }
}
