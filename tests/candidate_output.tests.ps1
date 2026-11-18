param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
if (-not $source.Contains('Public Sub SampleCandidate_WriteOutput(')) { throw 'Candidate output is missing.' }
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$test = @'
Public Function TestCandidateOutput() As String
    On Error GoTo Failed
    Dim inputWs As Worksheet, outputWs As Worksheet, records As Collection, grouped As Collection
    Dim allocated As Collection, targets As Object, firstRecord As Variant, lastRecord As Variant
    Set inputWs = ThisWorkbook.Worksheets(1)
    inputWs.Range("B1:C1").Merge: inputWs.Range("B1").Value = "Male"
    inputWs.Range("D1:E1").Merge: inputWs.Range("D1").Value = "Female"
    inputWs.Range("B2").Value = "20s": inputWs.Range("C2").Value = "30s"
    inputWs.Range("D2").Value = "20s": inputWs.Range("E2").Value = "30s"
    inputWs.Range("A3").Value = "Zone A": inputWs.Range("A4").Value = "Zone B"
    inputWs.Range("B3:E3").Value = Array(10, 20, 30, 40)
    inputWs.Range("B4:E4").Value = Array(0, 5, 6, 7)
    Set records = SampleCandidate_ReadTable(inputWs.Range("A3:A4"), inputWs.Range("B1:E2"), inputWs.Range("B3:E4"))
    Set grouped = SampleCandidate_GroupRecords(records, Array(1))
    Set targets = CreateObject("Scripting.Dictionary")
    firstRecord = records(1): lastRecord = records(4)
    targets.Add SampleCandidate_GroupKey(firstRecord, Array(1)), 10&
    targets.Add SampleCandidate_GroupKey(lastRecord, Array(1)), 10&
    Set allocated = SampleCandidate_Allocate(grouped, targets, 0, "proportional", 0.5, 0, 0)
    Set outputWs = ThisWorkbook.Worksheets.Add
    SampleCandidate_WriteOutput outputWs, inputWs.Range("A3:A4"), inputWs.Range("B1:E2"), inputWs.Range("B3:E4"), allocated, "proportional", 0.5
    Application.CalculateFull
    If outputWs.Range("A6").Value2 <> "Zone A" Then Err.Raise 5, , "row label"
    If outputWs.Range("D4").Value2 <> "Female" Then Err.Raise 5, , "merged column label"
    If outputWs.Range("F8").Value2 <> 20 Then Err.Raise 5, , "table total"
    If outputWs.Range("I6").Value2 <> 10 Then Err.Raise 5, , "population table missing"
    If outputWs.Range("G13").Value2 <> 10 Then Err.Raise 5, , "weight formula"
    If outputWs.Range("H13").Value2 <> 35 Then Err.Raise 5, , "group weight formula"
    If Abs(outputWs.Range("I13").Value2 - (100# / 35#)) > 0.000001 Then Err.Raise 5, , "theoretical allocation formula"
    outputWs.Range("J13").Value2 = outputWs.Range("J13").Value2 + 1
    Application.CalculateFull
    If outputWs.Range("F8").Value2 <> 21 Then Err.Raise 5, , "table is not linked to reviewed allocation"
    If outputWs.Range("T13").Value2 <> 1 Then Err.Raise 5, , "manual change is not flagged against group target"
    If inputWs.Range("B3").Value2 <> 10 Then Err.Raise 5, , "input modified"
    If Not outputWs.Range("A8").HasFormula And outputWs.Range("A8").Value2 = "" Then Err.Raise 5, , "total label missing"
    TestCandidateOutput = "PASS"
    Exit Function
Failed:
    TestCandidateOutput = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestCandidateOutput")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'Candidate crosstab output tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
