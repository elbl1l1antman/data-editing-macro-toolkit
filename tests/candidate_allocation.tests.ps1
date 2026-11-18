param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
if (-not $source.Contains('Public Function SampleCandidate_Allocate(')) { throw 'Candidate allocator is missing.' }
if (-not $source.Contains('Public Function SampleCandidate_GroupRecords(')) { throw 'Composite grouping is missing.' }
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$test = @'
Public Function TestCandidateAllocation() As String
    On Error GoTo Failed
    Dim rows As New Collection, targets As Object, result As Collection, row As Object
    Set targets = CreateObject("Scripting.Dictionary")
    rows.Add Array("A", 100#, 1, 1)
    rows.Add Array("A", 50#, 2, 1)
    rows.Add Array("A", 30#, 3, 1)
    rows.Add Array("B", 0#, 4, 1)
    rows.Add Array("B", 20#, 5, 1)
    targets.Add "A", 180&: targets.Add "B", 10&
    Set result = SampleCandidate_Allocate(rows, targets, 0, "proportional", 0.5, 0, 0)
    Set row = result(1): If row("final") <> 100 Then Err.Raise 5, , "baseline A1"
    Set row = result(2): If row("final") <> 50 Then Err.Raise 5, , "baseline A2"
    Set row = result(3): If row("final") <> 30 Then Err.Raise 5, , "baseline A3"
    targets("A") = 90&
    Set result = SampleCandidate_Allocate(rows, targets, 0, "proportional", 0.5, 0, 0)
    If result.Count <> 5 Then Err.Raise 5, , "record count"
    Set row = result(1): If row("final") <> 50 Then Err.Raise 5, , "proportional A1"
    Set row = result(2): If row("final") <> 25 Then Err.Raise 5, , "proportional A2"
    Set row = result(3): If row("final") <> 15 Then Err.Raise 5, , "proportional A3"
    Set row = result(4): If row("final") <> 0 Then Err.Raise 5, , "zero population"
    Set row = result(5): If row("final") <> 10 Then Err.Raise 5, , "group B"
    Set result = SampleCandidate_Allocate(rows, targets, 0, "sqrt", 0.5, 0, 0)
    Set row = result(1): If Abs(row("weight") - 10#) > 0.000001 Then Err.Raise 5, , "sqrt weight"
    Set row = result(1): Dim total As Long: total = row("final")
    Set row = result(2): total = total + row("final")
    Set row = result(3): total = total + row("final")
    If total <> 90 Then Err.Raise 5, , "sqrt target"
    Set result = SampleCandidate_Allocate(rows, targets, 0, "equal", 0.5, 0, 35)
    Set row = result(1): If row("final") > 35 Then Err.Raise 5, , "maximum ignored"
    Set row = result(3): If row("final") > 30 Then Err.Raise 5, , "population exceeded"
    Set result = SampleCandidate_Allocate(rows, targets, 0, "power", 1, 0, 0)
    Set row = result(1): If Abs(row("weight") - 100#) > 0.000001 Then Err.Raise 5, , "power weight"
    Dim constrained As New Collection, constrainedTargets As Object
    Set constrainedTargets = CreateObject("Scripting.Dictionary")
    constrained.Add Array("C", 20#, 1, 1)
    constrained.Add Array("C", 1000#, 2, 1)
    constrained.Add Array("C", 20#, 3, 1)
    constrainedTargets.Add "C", 70&
    Set result = SampleCandidate_Allocate(constrained, constrainedTargets, 0, "proportional", 0.5, 10, 30)
    Set row = result(1): If row("final") <> 20 Then Err.Raise 5, , "lower bound was locked too early"
    Set row = result(2): If row("final") <> 30 Then Err.Raise 5, , "upper bound was not applied"
    Set row = result(3): If row("final") <> 20 Then Err.Raise 5, , "remainder was not redistributed"
    Dim composite As New Collection, compositeRows As Collection, compositeTargets As Object, groupKey As String
    Set compositeTargets = CreateObject("Scripting.Dictionary")
    composite.Add Array("SW", "North", 100#, 1, 1)
    composite.Add Array("SW", "South", 50#, 2, 1)
    composite.Add Array("Non-SW", "North", 40#, 3, 1)
    Set compositeRows = SampleCandidate_GroupRecords(composite, Array(0, 1))
    groupKey = SampleCandidate_GroupKey(composite(1), Array(0, 1)): compositeTargets.Add groupKey, 30&
    groupKey = SampleCandidate_GroupKey(composite(2), Array(0, 1)): compositeTargets.Add groupKey, 20&
    groupKey = SampleCandidate_GroupKey(composite(3), Array(0, 1)): compositeTargets.Add groupKey, 10&
    Set result = SampleCandidate_Allocate(compositeRows, compositeTargets, 0, "proportional", 0.5, 0, 0)
    Set row = result(1): If row("final") <> 30 Then Err.Raise 5, , "first composite group"
    Set row = result(2): If row("final") <> 20 Then Err.Raise 5, , "second composite group"
    Set row = result(3): If row("final") <> 10 Then Err.Raise 5, , "third composite group"
    targets("A") = 0: targets("B") = 0
    Set result = SampleCandidate_Allocate(rows, targets, 0, "proportional", 0.5, 0, 0)
    Set row = result(1): If row("final") <> 0 Then Err.Raise 5, , "zero target"
    targets("A") = 181
    On Error Resume Next
    Set result = SampleCandidate_Allocate(rows, targets, 0, "proportional", 0.5, 0, 0)
    Dim errorNumber As Long: errorNumber = Err.Number
    On Error GoTo Failed
    If errorNumber = 0 Then Err.Raise 5, , "infeasible target accepted"
    TestCandidateAllocation = "PASS"
    Exit Function
Failed:
    TestCandidateAllocation = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestCandidateAllocation")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'Candidate grouped allocation tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
