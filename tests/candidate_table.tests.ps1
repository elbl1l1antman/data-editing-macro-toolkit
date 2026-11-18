param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$test = @'
Public Function TestCandidate() As String
    On Error GoTo Failed
    Dim ws As Worksheet, result As Collection, row As Variant
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("B1:C1").Merge
    ws.Range("B1").Value = "Male"
    ws.Range("D1:E1").Merge
    ws.Range("D1").Value = "Female"
    ws.Range("B2").Value = "20s": ws.Range("C2").Value = "30s"
    ws.Range("D2").Value = "20s": ws.Range("E2").Value = "30s"
    ws.Range("A3").Value = "Zone A": ws.Range("A4").Value = "Zone B"
    ws.Range("B3").Value = 10: ws.Range("C3").Value = 20
    ws.Range("D3").Value = 30: ws.Range("E3").Value = 40
    ws.Range("B4").Value = 0: ws.Range("C4").Value = 5
    ws.Range("D4").Value = 6: ws.Range("E4").Value = 7
    Set result = SampleCandidate_ReadTable(ws.Range("A3:A4"), ws.Range("B1:E2"), ws.Range("B3:E4"))
    If result.Count <> 8 Then Err.Raise 5, , "Incorrect record count"
    row = result(3)
    If row(0) <> "Zone A" Or row(1) <> "Female" Or row(2) <> "20s" Or row(3) <> 30 Then Err.Raise 5, , "Merged heading was lost"
    row = result(5)
    If row(3) <> 0 Then Err.Raise 5, , "Zero population was lost"
    ws.Range("F1:F2").Merge: ws.Range("F1").Value = "Total"
    ws.Range("F3").Value = 100: ws.Range("F4").Value = 18
    ws.Range("A5").Value = "Total": ws.Range("B5").Value = 10
    ws.Range("C5").Value = 25: ws.Range("D5").Value = 36: ws.Range("E5").Value = 47
    Set result = SampleCandidate_ReadTable(ws.Range("A3:A5"), ws.Range("B1:F2"), ws.Range("B3:F5"))
    If result.Count <> 8 Then Err.Raise 5, , "Totals were included as population"
    ws.Range("B3").Value = -1
    On Error Resume Next
    Set result = SampleCandidate_ReadTable(ws.Range("A3:A4"), ws.Range("B1:E2"), ws.Range("B3:E4"))
    Dim errorNumber As Long: errorNumber = Err.Number
    Err.Clear
    On Error GoTo Failed
    If errorNumber = 0 Then Err.Raise 5, , "Negative population accepted"
    ws.Range("B3").Value = 10
    On Error Resume Next
    Set result = SampleCandidate_ReadTable(ws.Range("A3:A20000"), ws.Range("B1:B2"), ws.Range("B3:B20000"))
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    If errorNumber = 0 Then Err.Raise 5, , "Oversized range accepted"
    TestCandidate = "PASS"
    Exit Function
Failed:
    TestCandidate = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestCandidate")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'Candidate multi-level population table tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
