param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
if (-not $source.Contains('Public Function SampleCandidate_DetectTable(')) { throw 'Table detection is missing.' }
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$test = @'
Public Function TestDetection() As String
    On Error GoTo Failed
    Dim ws As Worksheet, axes As Range, headings As Range, values As Range, records As Collection
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value = "district"
    ws.Range("B1:C1").Merge: ws.Range("B1").Value = "male"
    ws.Range("D1:E1").Merge: ws.Range("D1").Value = "female"
    ws.Range("F1:F2").Merge: ws.Range("F1").Value = "total"
    ws.Range("B2").Value = "20s": ws.Range("C2").Value = "30s"
    ws.Range("D2").Value = "20s": ws.Range("E2").Value = "30s"
    ws.Range("A3").Value = "area A": ws.Range("A4").Value = "area B": ws.Range("A5").Value = "total"
    ws.Range("B3:F3").Value = Array(10, 20, 30, 40, 100)
    ws.Range("B4:F4").Value = Array(2, 5, 6, 7, 20)
    ws.Range("B5:F5").Value = Array(12, 25, 36, 47, 120)
    If Not SampleCandidate_DetectTable(ws.Range("A1:F5"), axes, headings, values) Then Err.Raise 5, , "table not detected"
    If axes.Address <> "$A$3:$A$5" Then Err.Raise 5, , "row axes wrong: " & axes.Address
    If headings.Address <> "$B$1:$F$2" Then Err.Raise 5, , "multi-level header wrong: " & headings.Address
    If values.Address <> "$B$3:$F$5" Then Err.Raise 5, , "value area wrong: " & values.Address
    Set records = SampleCandidate_ReadTable(axes, headings, values)
    If records.Count <> 8 Then Err.Raise 5, , "subtotal exclusion wrong"
    ws.Cells.Clear
    ws.Range("A1").Value = "code": ws.Range("B1").Value = "industry"
    ws.Range("C1:E1").Merge: ws.Range("C1").Value = "size"
    ws.Range("C2").Value = "small": ws.Range("D2").Value = "mid": ws.Range("E2").Value = "large"
    ws.Range("A3").Value = 26: ws.Range("B3").Value = "manufacturing"
    ws.Range("C3:E3").Value = Array(100, 50, 10)
    If Not SampleCandidate_DetectTable(ws.Range("A1:E3"), axes, headings, values) Then Err.Raise 5, , "industry table not detected"
    If axes.Address <> "$A$3:$B$3" Or values.Address <> "$C$3:$E$3" Then Err.Raise 5, , "numeric identifier mistaken for data"
    Set records = SampleCandidate_ReadTable(axes, headings, values)
    If records.Count <> 3 Then Err.Raise 5, , "industry cells wrong"
    ws.Cells.Clear
    ws.Range("A1").Value = "region": ws.Range("B1").Value = "one": ws.Range("C1").Value = "two"
    ws.Range("D1").Value = "three": ws.Range("E1").Value = "four"
    ws.Range("A2").Value = "area": ws.Range("B2").Value = 10
    ws.Range("D2").Value = 30: ws.Range("E2").Value = 40
    If SampleCandidate_DetectTable(ws.Range("A1:E2"), axes, headings, values) Then Err.Raise 5, , "ambiguous sparse data was accepted"
    ws.Range("C1").Value = "target": ws.Range("C2").Value = 20
    If SampleCandidate_DetectTable(ws.Range("A1:E2"), axes, headings, values) Then Err.Raise 5, , "target column was mistaken for population"
    ws.Cells.Clear
    ws.Range("A1").Value = "region"
    ws.Range("B1:C1").Merge: ws.Range("B1").Value = "male"
    ws.Range("D1:E1").Merge: ws.Range("D1").Value = "female"
    ws.Range("B2").Value = "one": ws.Range("C2").Value = "two"
    ws.Range("D2").Value = "three": ws.Range("E2").Value = "target"
    ws.Range("A3").Value = "area": ws.Range("B3:E3").Value = Array(10, 20, 30, 40)
    If SampleCandidate_DetectTable(ws.Range("A1:E3"), axes, headings, values) Then Err.Raise 5, , "second header target column was accepted"
    TestDetection = "PASS"
    Exit Function
Failed:
    TestDetection = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestDetection")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'Candidate automatic population table detection tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
