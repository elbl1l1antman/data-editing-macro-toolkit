param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$candidate = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
$candidate = $candidate -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$baseline = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleTools.bas') -Raw -Encoding UTF8
$blocks = [ordered]@{}
function Add-BaselineProcedure([string]$Name) {
    if ($blocks.Contains($Name)) { return }
    $start = [regex]::Match($baseline, "(?m)^(Public|Private) (Sub|Function) $([regex]::Escape($Name))\b.*$")
    if (-not $start.Success) { throw "Baseline procedure not found: $Name" }
    $tail = $baseline.Substring($start.Index)
    $end = [regex]::Match($tail, "(?m)^End $($start.Groups[2].Value)\s*$")
    if (-not $end.Success) { throw "Baseline procedure end not found: $Name" }
    $body = $tail.Substring(0, $end.Index + $end.Length)
    $blocks[$Name] = $body
    $withoutStrings = [regex]::Replace($body, '"(?:[^"\r\n]|"")*"', '""')
    foreach ($match in [regex]::Matches($withoutStrings, '\bSampleDesign_[A-Za-z0-9_]+\b')) { Add-BaselineProcedure $match.Value }
}
Add-BaselineProcedure 'SampleDesign_Allocate'
$test = @'
Public Function TestCompare() As String
    On Error GoTo Failed
    Dim oldRows As New Collection, newRows As New Collection, groups As New Collection, issues As New Collection
    Dim group As Object, rec As Object, targets As Object, oldResult As Collection, newResult As Collection
    Dim i As Long, p As Long, t As Double, oldTime As Double, newTime As Double
    Set group = CreateObject("Scripting.Dictionary")
    group.Add "id", "A": group.Add "target", 1000&: group.Add "sourceRow", 1&: groups.Add group
    Set targets = CreateObject("Scripting.Dictionary"): targets.Add "A", 1000&
    For i = 1 To 300
        p = 100 + ((i * 13) Mod 271)
        Set rec = CreateObject("Scripting.Dictionary")
        rec.Add "targetGroupId", "A": rec.Add "pop", CDbl(p): rec.Add "min", 0&: rec.Add "max", 0&
        rec.Add "userWeight", 0#: rec.Add "sourceRows", CStr(i)
        oldRows.Add rec: newRows.Add Array("A", CDbl(p), i, 1)
    Next i
    t = Timer
    Set oldResult = SampleDesign_Allocate(oldRows, groups, "proportional", 0.5, 0, 0, issues)
    oldTime = Timer - t
    t = Timer
    Set newResult = SampleCandidate_Allocate(newRows, targets, 0, "proportional", 0.5, 0, 0)
    newTime = Timer - t
    If oldResult.Count <> newResult.Count Or issues.Count <> 0 Then Err.Raise 5, , "record count or warnings"
    Dim oldItem As Object, newItem As Object, oldTotal As Long, newTotal As Long, differences As Long
    For i = 1 To oldResult.Count
        Set oldItem = oldResult(i): Set newItem = newResult(i)
        oldTotal = oldTotal + CLng(oldItem("final")): newTotal = newTotal + CLng(newItem("final"))
        If oldItem("final") <> newItem("final") Then differences = differences + 1
    Next i
    If oldTotal <> 1000 Or newTotal <> 1000 Then Err.Raise 5, , "target mismatch"
    TestCompare = "PASS|old=" & CStr(oldTime) & "|new=" & CStr(newTime) & "|different_cells=" & CStr(differences)
    Exit Function
Failed:
    TestCompare = "FAIL|" & CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($candidate + "`r`n" + ($blocks.Values -join "`r`n") + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestCompare")
    if (-not $result.StartsWith('PASS|')) { throw $result }
    Write-Host $result
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
