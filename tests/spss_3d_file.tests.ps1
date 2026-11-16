param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [string]$WorkbookPath = '',
    [switch]$AllTables
)
$ErrorActionPreference = 'Stop'
if (-not $WorkbookPath) { Write-Host 'Real SPSS workbook test skipped: WorkbookPath not provided.'; exit 0 }
$source = Get-Content -Raw -Encoding UTF8 (Join-Path $ProjectRoot 'src\SpssTools.bas')
$blocks = [ordered]@{}
function Add-Procedure([string]$name) {
    if ($blocks.Contains($name)) { return }
    $start = [regex]::Match($source, "(?m)^(Public|Private)\s+(Sub|Function)\s+$([regex]::Escape($name))\b.*$")
    if (-not $start.Success) { throw "Missing procedure: $name" }
    $tail = $source.Substring($start.Index)
    $end = [regex]::Match($tail, "(?m)^End $($start.Groups[2].Value)\s*$")
    if (-not $end.Success) { throw "Missing end: $name" }
    $body = $tail.Substring(0, $end.Index + $end.Length)
    $blocks[$name] = $body
    $unquoted = [regex]::Replace($body, '"(?:[^"\r\n]|"")*"', '""')
    foreach ($match in [regex]::Matches($unquoted, '\bSPSS_[A-Za-z0-9_]+\b')) { Add-Procedure $match.Value }
}
foreach ($name in @('SPSS_CollectTablesFromDataSheet', 'SPSS_CollectUniqueBanners', 'SPSS_CopyRebuiltTable')) { Add-Procedure $name }
$excel = New-Object -ComObject Excel.Application
$driver = $null; $inputWorkbook = $null
try {
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $driver = $excel.Workbooks.Add()
    $module = $driver.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString("Option Explicit`r`n" + ($blocks.Values -join "`r`n") + @'

Public Function VerifyActualWorkbook(ByVal inputWorkbook As Workbook) As String
    On Error GoTo Failed
    Dim ws As Worksheet, output As Worksheet, tables As Collection, onlyOne As New Collection
    Dim banners As Collection, deleted As Object, order As New Collection, info As Variant, target As String, nextRow As Long
    Set ws = inputWorkbook.Worksheets("Sheet1")
    Set tables = SPSS_CollectTablesFromDataSheet(ws)
    If tables.Count < 2 Then Err.Raise 513, , "Table detection failed"
    info = tables(1)
    If CLng(info(1)) <> 2 Then Err.Raise 513, , "First title row incorrect"
    onlyOne.Add info
    Set deleted = CreateObject("Scripting.Dictionary")
    Set banners = SPSS_CollectUniqueBanners(onlyOne, ws, deleted)
    If banners.Count < 4 Then Err.Raise 513, , "Square-marked banners missing: " & banners.Count
    target = SPSS_NormalizeBannerName(CStr(ws.Cells(23, 2).Value2))
    order.Add target
    Set output = ThisWorkbook.Worksheets.Add
    nextRow = 1
    SPSS_CopyRebuiltTable ws, output, info, nextRow, deleted, order
    If output.Cells(5, 2).Value2 <> ws.Cells(23, 2).Value2 Then Err.Raise 513, , "Banner did not move to first position"
    If output.Cells(5, 2).MergeArea.Rows.Count <> ws.Cells(23, 2).MergeArea.Rows.Count Then Err.Raise 513, , "Merged banner dimensions changed"
    If output.Cells(nextRow - 1, 2).Value2 <> ws.Cells(27, 2).Value2 Then
        Dim detail As String, r As Long
        For r = 1 To nextRow - 1
            If Len(CStr(output.Cells(r, 2).Value2)) > 0 Then detail = detail & r & ":" & CStr(output.Cells(r, 2).Value2) & ";"
        Next r
        Err.Raise 513, , "Table footer mismatch; source end=" & CLng(info(2)) & ", output last row=" & nextRow - 1 & ", rows=" & detail
    End If
    If ws.Cells(6, 2).Value2 = output.Cells(5, 2).Value2 Then Err.Raise 513, , "Source order not preserved"
    VerifyActualWorkbook = "PASS: tables=" & tables.Count & "; first-table banners=" & banners.Count
    Exit Function
Failed:
    VerifyActualWorkbook = CStr(Err.Number) & ": " & Err.Description
End Function
Public Function VerifyAllTables(ByVal inputWorkbook As Workbook) As String
    On Error GoTo Failed
    Dim ws As Worksheet, output As Worksheet, tables As Collection, info As Variant, order As New Collection
    Dim deleted As Object, i As Long, startRow As Long, nextRow As Long, lastCol As Long
    Set ws = inputWorkbook.Worksheets("Sheet1")
    Set tables = SPSS_CollectTablesFromDataSheet(ws)
    Set deleted = CreateObject("Scripting.Dictionary")
    order.Add SPSS_NormalizeBannerName(CStr(ws.Cells(23, 2).Value2))
    Set output = ThisWorkbook.Worksheets.Add
    nextRow = 1
    For i = 1 To tables.Count
        info = tables(i): startRow = nextRow
        SPSS_CopyRebuiltTable ws, output, info, nextRow, deleted, order
        If nextRow - startRow <> CLng(info(2)) - CLng(info(1)) + 1 Then Err.Raise 513, , "Table height changed: " & i
        lastCol = SPSS_TableLastCol(ws, CLng(info(1)), CLng(info(2)))
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(CLng(info(1)), 1), ws.Cells(CLng(info(2)), lastCol))) <> _
           Application.WorksheetFunction.CountA(output.Range(output.Cells(startRow, 1), output.Cells(nextRow - 1, lastCol))) Then Err.Raise 513, , "Nonempty cells changed: " & i
        nextRow = nextRow + 1
    Next i
    VerifyAllTables = "PASS: reconstructed " & tables.Count & " tables"
    Exit Function
Failed:
    VerifyAllTables = CStr(Err.Number) & ": " & Err.Description
End Function
'@)
    $inputWorkbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
    if (-not $inputWorkbook.ReadOnly) { throw 'Source workbook did not open read-only' }
    $result = [string]$excel.Run("'$($driver.Name)'!VerifyActualWorkbook", $inputWorkbook)
    if (-not $result.StartsWith('PASS:')) { throw "Real workbook test failed: $result" }
    Write-Host $result
    if ($AllTables) {
        $result = [string]$excel.Run("'$($driver.Name)'!VerifyAllTables", $inputWorkbook)
        if (-not $result.StartsWith('PASS:')) { throw "Full real workbook test failed: $result" }
        Write-Host $result
    }
}
finally {
    if ($inputWorkbook) { $inputWorkbook.Close($false) }
    if ($driver) { $driver.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
