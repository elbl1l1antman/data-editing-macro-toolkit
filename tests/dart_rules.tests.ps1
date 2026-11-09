param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$workbook = $null
try {
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Const BIZ_APP_NAME As String = "DataEditingMacroToolkit"
Public Const TOOLKIT_CANCEL_ERROR_NUMBER As Long = vbObjectError + 9130
Public Function Toolkit_GetRangeFromUserForm(a As String, b As String, ByRef c As Range) As Boolean
End Function
Public Sub Toolkit_ProgressShow(a As String, b As String, c As Long)
End Sub
Public Sub Toolkit_ProgressUpdate(a As Long, b As Long, c As String)
End Sub
Public Sub Toolkit_ProgressClose()
End Sub
Public Sub Toolkit_RaiseIfCancelled()
End Sub
Public Function CheckDartRules() As String
    On Error GoTo Failed
    Dim doc As Object, value As Variant, names As Object, codes As Object, stocks As Object, issue As String
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.LoadXML "<result><status>000</status><list><sj_div>IS</sj_div><account_id>ifrs-full_Revenue</account_id><account_nm>Revenue</account_nm><thstrm_amount>12,345</thstrm_amount><currency>KRW</currency></list></result>"
    value = Dart_ExtractRevenue(doc)
    If value(0) <> "12345" Or value(1) <> "KRW" Then Err.Raise 513, , "Revenue parse"
    doc.LoadXML "<result><status>000</status><list><fo_bbm>" & ChrW(&HD569) & ChrW(&HACC4) & "</fo_bbm><sexdstn>" & ChrW(&HD569) & ChrW(&HACC4) & "</sexdstn><sm>100</sm></list><list><fo_bbm>A</fo_bbm><sexdstn>M</sexdstn><sm>60</sm></list><list><fo_bbm>A</fo_bbm><sexdstn>F</sexdstn><sm>40</sm></list></result>"
    value = Dart_ExtractEmployees(doc)
    If value(0) <> "100" Then Err.Raise 513, , "Employee total double counted"
    Set names = CreateObject("Scripting.Dictionary")
    Set codes = CreateObject("Scripting.Dictionary")
    Set stocks = CreateObject("Scripting.Dictionary")
    names.Add "ACME", "00123456"
    names.Add "DUPLICATE", ""
    codes.Add "00123456", True
    stocks.Add "005930", "00123456"
    If Dart_FindCompany("acme", names, codes, stocks, issue) <> "00123456" Then Err.Raise 513, , "Exact name"
    If Dart_FindCompany("005930", names, codes, stocks, issue) <> "00123456" Then Err.Raise 513, , "Stock code"
    issue = ""
    If Dart_FindCompany("duplicate", names, codes, stocks, issue) <> "" Or Len(issue) = 0 Then Err.Raise 513, , "Ambiguous name"
    If Not Dart_ValidKey(String$(40, "A")) Or Dart_ValidKey("wrong") Then Err.Raise 513, , "Key validation"
    CheckDartRules = "PASS"
    Exit Function
Failed:
    CheckDartRules = CStr(Err.Number) & ": " & Err.Description
End Function
'@)
    $tempPath = Join-Path $env:TEMP ('DartTools-' + [guid]::NewGuid().ToString('N') + '.bas')
    $source = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot 'src\DartTools.bas')
    $encoding = [Text.Encoding]::GetEncoding(949, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
    [IO.File]::WriteAllText($tempPath, $source, $encoding)
    $workbook.VBProject.VBComponents.Import($tempPath) | Out-Null
    $result = $excel.Run("'$($workbook.Name)'!CheckDartRules")
    if ($result -ne 'PASS') { throw "DART test failed: $result" }
    Write-Host 'DART matching, revenue and employee rules passed.'
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    if ($tempPath -and (Test-Path -LiteralPath $tempPath)) { Remove-Item -LiteralPath $tempPath -Force }
}
