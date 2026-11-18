param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$path = Join-Path $ProjectRoot 'src\KosisOutput.bas'
if (-not (Test-Path -LiteralPath $path)) { throw 'KOSIS output module missing.' }
$source = Get-Content -LiteralPath $path -Raw -Encoding UTF8
if (-not $source.Contains('Public Function Kosis_WriteGroups(')) { throw 'Multi-period output missing.' }
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$test = @'
Public Function TestKosisOutput() As String
    On Error GoTo Failed
    Dim records As New Collection, metadata As New Collection, missing As New Collection
    Dim row As Object, item As Object, axes As New Collection, wb As Workbook, sheets As Collection, countBefore As Long, n As Long, alertsBefore As Boolean
    Dim groups As New Collection, group As Object, records2 As New Collection, review As New Collection
    Set wb = ThisWorkbook
    Set item = CreateObject("Scripting.Dictionary"): item.Add "axis", "A": item.Add "code", "00": item.Add "name", "National": metadata.Add item
    Set item = CreateObject("Scripting.Dictionary"): item.Add "axis", "B": item.Add "code", "01": item.Add "name", "Industry": metadata.Add item
    Set item = CreateObject("Scripting.Dictionary"): item.Add "axis", "E": item.Add "code", "1": item.Add "name", "Small": metadata.Add item
    Set item = CreateObject("Scripting.Dictionary"): item.Add "axis", "E": item.Add "code", "2": item.Add "name", "Large": metadata.Add item
    Set row = CreateObject("Scripting.Dictionary"): row.Add "C_A", "00": row.Add "C_B", "01": row.Add "C_E", "1": row.Add "ITEM", "T1": row.Add "period", "2024": row.Add "value", "42": row.Add "status", "OBS_STATUS=p": row.Add "updated", "2025-12-29": records.Add row
    Set row = CreateObject("Scripting.Dictionary"): row.Add "C_A", "00": row.Add "C_B", "01": row.Add "C_E", "2": row.Add "ITEM", "T1": row.Add "period", "2024": row.Add "value", "X": records.Add row
    axes.Add "A": axes.Add "B"
    Set sheets = Kosis_WriteSheets(wb, records, metadata, missing, axes, Array("A", "B", "E"), "Example", "2024", "T1")
    If sheets.Count <> 3 Then Err.Raise 5, , "wrong sheet count"
    If sheets(1).Range("A2").Value2 <> "00" Then Err.Raise 5, , "leading zero lost"
    If sheets(1).Cells(2, 10).Value2 <> "OBS_STATUS=p" Or sheets(1).Cells(2, 11).Value2 <> "2025-12-29" Then Err.Raise 5, , "status or update lost"
    If sheets(2).Range("B6").Value2 <> 42 Then Err.Raise 5, , "numeric value lost"
    If sheets(2).Range("C6").Value2 <> "X" Then Err.Raise 5, , "suppressed symbol lost"
    If sheets(3).Cells(2, 1).Value2 = "" Then Err.Raise 5, , "review entry missing"
    countBefore = wb.Worksheets.Count: alertsBefore = Application.DisplayAlerts
    records.Add row
    On Error Resume Next
    Set sheets = Kosis_WriteSheets(wb, records, metadata, missing, axes, Array("A", "B", "E"), "Example", "2024", "T1")
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    If n = 0 Or wb.Worksheets.Count <> countBefore Then Err.Raise 5, , "duplicate rollback failed"
    If Application.DisplayAlerts <> alertsBefore Then Err.Raise 5, , "Excel alert setting changed"
    records.Remove records.Count
    Set row = CreateObject("Scripting.Dictionary"): row.Add "C_A", "00": row.Add "C_B", "01": row.Add "C_E", "1": row.Add "ITEM", "T1": row.Add "period", "2023": row.Add "value", "39": records2.Add row
    Set group = CreateObject("Scripting.Dictionary"): group.Add "records", records: group.Add "missing", missing: group.Add "period", "2024": group.Add "item", "T1": groups.Add group
    Set group = CreateObject("Scripting.Dictionary"): group.Add "records", records2: group.Add "missing", missing: group.Add "period", "2023": group.Add "item", "T1": groups.Add group
    countBefore = wb.Worksheets.Count
    Set sheets = Kosis_WriteGroups(wb, groups, metadata, axes, Array("A", "B", "E"), "Example")
    If sheets.Count <> 3 Or wb.Worksheets.Count <> countBefore + 3 Then Err.Raise 5, , "multi-period sheet count wrong"
    If sheets(1).Cells(4, 7).Value2 <> "2023" Then Err.Raise 5, , "second period missing from list"
    If sheets(2).Range("A8").Value2 <> "Example" Then Err.Raise 5, , "second table block missing"
    Set records2 = New Collection
    Set row = CreateObject("Scripting.Dictionary"): row.Add "C_A", "00": row.Add "C_B", "01": row.Add "C_E", "2": row.Add "ITEM", "T1": row.Add "period", "2024": row.Add "value", "5": records2.Add row
    Set groups = New Collection: Set group = CreateObject("Scripting.Dictionary")
    review.Add "age review"
    group.Add "records", records2: group.Add "rawRecords", records: group.Add "missing", missing: group.Add "period", "2024": group.Add "item", "T1": group.Add "review", review: groups.Add group
    Set sheets = Kosis_WriteGroups(wb, groups, metadata, axes, Array("A", "B", "E"), "Example")
    If sheets(1).Cells(2, 9).Value2 <> 42 Or sheets(1).Cells(3, 9).Value2 <> "X" Then Err.Raise 5, , "raw records not preserved"
    If sheets(2).Range("B6").Value2 <> 5 Then Err.Raise 5, , "pivot data not used"
    If sheets(3).Range("A4").Value2 <> "age review" Then Err.Raise 5, , "age review not included"
    TestKosisOutput = "PASS"
    Exit Function
Failed:
    TestKosisOutput = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestKosisOutput")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'KOSIS output tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
