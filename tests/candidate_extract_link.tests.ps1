param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$candidate = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
$candidate = $candidate -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$extract = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleTools.bas') -Raw -Encoding UTF8
if (-not $extract.Contains('Public Function SampleExtract_LoadCandidateTargets(')) { throw 'Candidate extraction link is missing.' }
if (-not $extract.Contains('Private Function SampleExtract_SelectCandidateColumns(')) { throw 'Candidate mapping selection is missing.' }
if (-not $extract.Contains('Private Function SampleExtract_SelectDesignSheet(')) { throw 'Candidate design selection is missing.' }
if (-not $extract.Contains('ByRef candidateValid As Boolean')) { throw 'Candidate review gate is missing.' }
if (-not $extract.Contains('Private Function SampleExtract_ValidateCandidateCodes(')) { throw 'Candidate code comparison is missing.' }
$blocks = [ordered]@{}
function Add-Procedure([string]$Name) {
    if ($blocks.Contains($Name)) { return }
    $start = [regex]::Match($extract, "(?m)^(Public|Private) (Sub|Function) $([regex]::Escape($Name))\b.*$")
    if (-not $start.Success) { throw "Procedure missing: $Name" }
    $tail = $extract.Substring($start.Index)
    $end = [regex]::Match($tail, "(?m)^End $($start.Groups[2].Value)\s*$")
    if (-not $end.Success) { throw "Procedure end missing: $Name" }
    $body = $tail.Substring(0, $end.Index + $end.Length)
    $blocks[$Name] = $body
    $withoutStrings = [regex]::Replace($body, '"(?:[^"\r\n]|"")*"', '""')
    foreach ($match in [regex]::Matches($withoutStrings, '\bSampleExtract_[A-Za-z0-9_]+\b')) { Add-Procedure $match.Value }
}
Add-Procedure 'SampleExtract_LoadCandidateTargets'
Add-Procedure 'SampleExtract_SelectCandidateColumns'
Add-Procedure 'SampleExtract_SelectDesignSheet'
Add-Procedure 'SampleExtract_PrepareReview'
Add-Procedure 'SampleExtract_ValidateCandidateCodes'
$testClass = @'
Option Explicit
Public HasHeader As Boolean
Public StratumColumnIndexes As String
Public TargetMode As String
Public ReplacementMultiplier As Long
Public ReviewResult As String
Public TargetResult As String
Public WarningResult As String
Public Sub ConfigureReview(ByVal reviewText As String, ByVal targetText As String, ByVal warningText As String)
    ReviewResult = reviewText: TargetResult = targetText: WarningResult = warningText
End Sub
'@
$test = @'
Public Function Toolkit_ButtonMenu(ByVal title As String, ByVal guide As String, ByVal options As String, ByVal defaultChoice As String) As String
    Dim lines As Variant, item As Variant
    If Len(defaultChoice) > 0 Then Toolkit_ButtonMenu = defaultChoice: Exit Function
    lines = Split(options, vbLf): item = Split(CStr(lines(0)), vbTab)
    Toolkit_ButtonMenu = CStr(item(0))
End Function
Public Function TestCandidateExtractLink() As String
    On Error GoTo Failed
    Dim src As Worksheet, dest As Worksheet, records As Collection, grouped As Collection, allocated As Collection
    Dim targets As Object, loaded As Object, record As Variant, warning As String, k As String
    Set src = ThisWorkbook.Worksheets(1)
    src.Range("B1:C1").Merge: src.Range("B1").Value = "Male"
    src.Range("D1:E1").Merge: src.Range("D1").Value = "Female"
    src.Range("B2").Value = "20s": src.Range("C2").Value = "30s"
    src.Range("D2").Value = "20s": src.Range("E2").Value = "30s"
    src.Range("A3").Value = "Zone A": src.Range("A4").Value = "Zone B"
    src.Range("B3:E3").Value = Array(10, 20, 30, 40)
    src.Range("B4:E4").Value = Array(0, 5, 6, 7)
    Set records = SampleCandidate_ReadTable(src.Range("A3:A4"), src.Range("B1:E2"), src.Range("B3:E4"))
    Set grouped = SampleCandidate_GroupRecords(records, Array())
    Set targets = CreateObject("Scripting.Dictionary"): targets.Add "*", 20&
    Set allocated = SampleCandidate_Allocate(grouped, targets, 0, "proportional", 0.5, 0, 0)
    Set dest = ThisWorkbook.Worksheets.Add
    SampleCandidate_WriteOutput dest, src.Range("A3:A4"), src.Range("B1:E2"), src.Range("B3:E4"), allocated, "proportional", 0.5
    Application.CalculateFull
    Dim chosen As Worksheet, cancelled As Boolean
    Set chosen = SampleExtract_SelectDesignSheet(ThisWorkbook, dest, cancelled)
    If chosen Is Nothing Or cancelled Then Err.Raise 5, , "design sheet not selected"
    If Not chosen Is dest Then Err.Raise 5, , "wrong design sheet selected"
    Dim mapping As Variant
    mapping = SampleExtract_SelectCandidateColumns(dest, Array("zone", "gender", "age"), Array(1, 2, 3), warning)
    If Not IsArray(mapping) Then Err.Raise 5, , "mapping choice missing"
    If mapping(0) <> 1 Or mapping(1) <> 2 Or mapping(2) <> 3 Then Err.Raise 5, , "mapping choices incorrect"
    Set loaded = SampleExtract_LoadCandidateTargets(dest, Array(1, 2, 3), warning)
    k = "6:Zone A|4:Male|3:20s|"
    If Len(warning) <> 0 Or Not loaded.Exists(k) Then Err.Raise 5, , "candidate key missing"
    If CLng(loaded(k)) <> 2 Or loaded.Count <> 8 Then Err.Raise 5, , "candidate targets wrong"
    Dim roster As Worksheet, wizard As TestWizard, i As Long, n As Long, z As Long, g As Long, a As Long, valid As Boolean
    Set roster = ThisWorkbook.Worksheets.Add
    roster.Range("A1:C1").Value = Array("zone", "gender", "age")
    n = 2
    For z = 1 To 2
        For g = 1 To 2
            For a = 1 To 2
                For i = 1 To 10
                    roster.Cells(n, 1).Value = IIf(z = 1, "Zone A", "Zone B")
                    roster.Cells(n, 2).Value = IIf(g = 1, "Male", "Female")
                    roster.Cells(n, 3).Value = IIf(a = 1, "20s", "30s")
                    n = n + 1
                Next i
            Next a
        Next g
    Next z
    Set wizard = New TestWizard
    wizard.HasHeader = True: wizard.StratumColumnIndexes = "1,2,3"
    wizard.TargetMode = "auto": wizard.ReplacementMultiplier = 3
    SampleExtract_PrepareReview wizard, roster.Range("A1:C81"), dest, valid
    If Not valid Or Len(wizard.WarningResult) <> 0 Or InStr(wizard.TargetResult, k & vbTab & "2") = 0 Then Err.Raise 5, , "review did not connect candidate targets"
    dest.Range("A13").Value2 = "001": roster.Range("A2").Value2 = "1"
    warning = ""
    If SampleExtract_ValidateCandidateCodes(dest, Array(1, 2, 3), roster.Range("A1:C81").Value2, Array(1, 2, 3), 2, warning) Then Err.Raise 5, , "cross-source numeric label mismatch accepted"
    If Len(warning) = 0 Then Err.Raise 5, , "numeric label mismatch was not explained"
    dest.Range("A13").Value2 = "Zone A": roster.Range("A2").Value2 = "Zone A"
    roster.Range("A2:C10").Delete xlShiftUp
    SampleExtract_PrepareReview wizard, roster.Range("A1:C72"), dest, valid
    If valid Or Len(wizard.WarningResult) = 0 Then Err.Raise 5, , "insufficient population was ignored"
    roster.Range("A2:C2").Delete xlShiftUp
    SampleExtract_PrepareReview wizard, roster.Range("A1:C71"), dest, valid
    If valid Or Len(wizard.WarningResult) = 0 Then Err.Raise 5, , "missing population stratum was ignored"
    warning = ""
    Set loaded = SampleExtract_LoadCandidateTargets(dest, Array(1, 1), warning)
    If Len(warning) = 0 Or loaded.Count <> 0 Then Err.Raise 5, , "duplicate mapping was accepted"
    dest.Range("A13").Value2 = "001": dest.Range("A14").Value2 = "1"
    warning = ""
    Set loaded = SampleExtract_LoadCandidateTargets(dest, Array(1), warning)
    If Len(warning) = 0 Or loaded.Count <> 0 Then Err.Raise 5, , "ambiguous numeric labels were merged"
    dest.Range("A13").Value2 = "Zone A": dest.Range("A14").Value2 = "Zone A"
    dest.Range("J13").Value2 = dest.Range("J13").Value2 + 1
    Application.CalculateFull
    warning = ""
    Set loaded = SampleExtract_LoadCandidateTargets(dest, Array(1, 2, 3), warning)
    If Len(warning) = 0 Or loaded.Count <> 0 Then Err.Raise 5, , "target discrepancy was ignored"
    TestCandidateExtractLink = "PASS"
    Exit Function
Failed:
    TestCandidateExtractLink = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $classModule = $book.VBProject.VBComponents.Add(2)
    $classModule.Name = 'TestWizard'
    $classModule.CodeModule.AddFromString($testClass)
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($candidate + "`r`n" + ($blocks.Values -join "`r`n") + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestCandidateExtractLink")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'Candidate extraction target mapping and mismatch tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
