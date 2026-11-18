param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
if (-not $source.Contains('Public Function SampleCandidate_AllocateStages(')) { throw 'Staged allocation is missing.' }
if (-not $source.Contains('Public Sub SampleCandidate_WriteStageSummary(')) { throw 'Stage summary is missing.' }
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$test = @'
Public Function TestStagedAllocation() As String
    On Error GoTo Failed
    Dim records As New Collection, targets As Object, result As Collection, review As Collection, row As Object
    Set targets = CreateObject("Scripting.Dictionary")
    records.Add Array("zone", "X", "M", 80#, 1, 1)
    records.Add Array("zone", "X", "F", 20#, 1, 2)
    records.Add Array("zone", "Y", "M", 30#, 2, 1)
    records.Add Array("zone", "Y", "F", 20#, 2, 2)
    targets.Add SampleCandidate_GroupKey(records(1), Array(0)), 12&
    Set result = SampleCandidate_AllocateStages(records, targets, Array(0), 1, "equal", 0.5, "proportional", 0.5, review)
    If review.Count <> 2 Or result.Count <> 4 Then Err.Raise 5, , "missing staged results"
    Set row = result(1): If row("final") <> 5 Or row("target") <> 6 Then Err.Raise 5, , "stage X male"
    Set row = result(2): If row("final") <> 1 Or row("target") <> 6 Then Err.Raise 5, , "stage X female"
    Set row = result(3): If row("final") <> 4 Or row("target") <> 6 Then Err.Raise 5, , "stage Y male"
    Set row = result(4): If row("final") <> 2 Or row("target") <> 6 Then Err.Raise 5, , "stage Y female"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value = "result"
    SampleCandidate_WriteStageSummary ws, review, "equal", 0.5
    If ws.Range("A5").Value2 <> "zone" Then Err.Raise 5, , "parent summary missing"
    If ws.Range("B5").Value2 <> "zone / X" Then Err.Raise 5, , "child summary missing"
    If ws.Range("C5").Value2 <> 100 Or ws.Range("F5").Value2 <> 6 Then Err.Raise 5, , "stage summary values wrong"
    records.Add Array("zone B", "X", "M", 1#, 3, 1)
    records.Add Array("zone B", "X", "F", 0#, 3, 2)
    records.Add Array("zone B", "Y", "M", 10#, 4, 1)
    records.Add Array("zone B", "Y", "F", 10#, 4, 2)
    targets.Add SampleCandidate_GroupKey(records(5), Array(0)), 4&
    Set result = SampleCandidate_AllocateStages(records, targets, Array(0), 1, "equal", 0.5, "proportional", 0.5, review)
    If review.Count <> 4 Or result.Count <> 8 Then Err.Raise 5, , "multiple groups missing"
    Dim i As Long, totalA As Long, totalB As Long
    For i = 1 To result.Count
        Set row = result(i)
        If i <= 4 Then totalA = totalA + row("final") Else totalB = totalB + row("final")
    Next i
    If totalA <> 12 Or totalB <> 4 Then Err.Raise 5, , "group targets not preserved"
    Set row = result(5): If row("final") <> 1 Then Err.Raise 5, , "population cap not enforced"
    Dim errorNumber As Long
    On Error Resume Next
    Set result = SampleCandidate_AllocateStages(records, targets, Array(0), 0, "equal", 0.5, "proportional", 0.5, review)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    If errorNumber = 0 Then Err.Raise 5, , "repeated grouping axis accepted"
    TestStagedAllocation = "PASS"
    Exit Function
Failed:
    TestStagedAllocation = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestStagedAllocation")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'Candidate staged allocation tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
