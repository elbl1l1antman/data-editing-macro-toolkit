param([string[]]$Files = @(),
      [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
if ($Files.Count -eq 0) { Write-Host 'SKIP: provide four reference workbooks to run this test.'; exit 0 }
if ($Files.Count -ne 4) { throw 'Four reference workbooks are required.' }
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$test = @'
Public Function TestRealTable(ByVal ws As Worksheet, ByVal rowsAddress As String, _
    ByVal headsAddress As String, ByVal valuesAddress As String, ByVal expectedCount As Long) As String
    On Error GoTo Failed
    Dim records As Collection, item As Variant, total As Double, i As Long
    Set records = SampleCandidate_ReadTable(ws.Range(rowsAddress), ws.Range(headsAddress), ws.Range(valuesAddress))
    If records.Count <> expectedCount Then Err.Raise 5, , "record count: " & records.Count
    For i = 1 To records.Count
        item = records(i)
        total = total + CDbl(item(UBound(item) - 2))
    Next i
    If total <> Application.Sum(ws.Range(valuesAddress)) Then Err.Raise 5, , "population sum mismatch"
    Dim grouped As Collection, result As Collection, targets As Object, outputWs As Worksheet
    Set targets = CreateObject("Scripting.Dictionary"): targets.Add "*", 20&
    Set grouped = SampleCandidate_GroupRecords(records, Array())
    Set result = SampleCandidate_Allocate(grouped, targets, 0, "proportional", 0.5, 0, 0)
    Set outputWs = ThisWorkbook.Worksheets.Add
    SampleCandidate_WriteOutput outputWs, ws.Range(rowsAddress), ws.Range(headsAddress), ws.Range(valuesAddress), result, "proportional", 0.5
    Application.CalculateFull
    Dim totalRow As Long, totalCol As Long
    totalRow = 4 + ws.Range(headsAddress).Rows.Count + ws.Range(valuesAddress).Rows.Count
    totalCol = ws.Range(rowsAddress).Columns.Count + ws.Range(valuesAddress).Columns.Count + 1
    If outputWs.Cells(totalRow, totalCol).Value2 <> 20 Then Err.Raise 5, , "output target mismatch"
    TestRealTable = "PASS"
    Exit Function
Failed:
    TestRealTable = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $hostBook = $null; $opened = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false; $excel.AutomationSecurity = 3
    $hostBook = $excel.Workbooks.Add()
    $module = $hostBook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $cases = @(
        @{Sheet=4; Rows='B4:C9'; Heads='F3:J3'; Values='F4:J9'; Count=30},
        @{Sheet=2; Rows='A6:B12'; Heads='C4:I5'; Values='C6:I12'; Count=49},
        @{Sheet=2; Rows='A4:A8'; Heads='B2:M3'; Values='B4:M8'; Count=60},
        @{Sheet=1; Rows='A5:A9'; Heads='B3:K4'; Values='B5:K9'; Count=50}
    )
    for ($i=0; $i -lt $cases.Count; $i++) {
        $book = $excel.Workbooks.Open($Files[$i], 0, $true)
        $opened += $book
        $case = $cases[$i]
        $result = [string]$excel.Run("'$($hostBook.Name)'!TestRealTable", $book.Worksheets.Item($case.Sheet), $case.Rows, $case.Heads, $case.Values, $case.Count)
        if ($result -ne 'PASS') { throw "Reference $($i+1): $result" }
        Write-Host "Reference $($i+1): $result ($($case.Count) cells)"
    }
} finally {
    foreach ($book in $opened) { $book.Close($false) }
    if ($hostBook) { $hostBook.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
