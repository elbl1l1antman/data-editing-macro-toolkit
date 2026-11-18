param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path, [string]$KeyFile = '')
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\KosisTools.bas') -Raw -Encoding UTF8
if (-not $source.Contains('Public Function Kosis_ParsePeriods(')) { throw 'KOSIS period parser missing.' }
if (-not $source.Contains('Public Function Kosis_CacheTimestamp(')) { throw 'KOSIS metadata cache missing.' }
$source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$source = $source -replace 'Option Explicit', "Option Explicit`r`nPublic Const BIZ_APP_NAME As String = `"KosisFoundationTest`"`r`nPublic Const TOOLKIT_CANCEL_ERROR_NUMBER As Long = 12345"
$test = @'
Public Sub Toolkit_RaiseIfCancelled()
End Sub
Public Sub Toolkit_ProgressUpdate(ByVal current As Long, ByVal total As Long, ByVal message As String)
End Sub
Public Function TestKosisMetadata() As String
    On Error GoTo Failed
    Dim xml As Object, rows As Collection, row As Object, options As Collection, n As Long, bands As New Collection, band As Object, uncovered As Long
    Dim selection As New Collection, overlaps As Collection
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.LoadXML "<response><Structures><MetaRow><objId>B</objId><objNm>Industry</objNm><itmId>01</itmId><itmNm>Agriculture</itmNm><upItmId>00</upItmId></MetaRow></Structures></response>"
    Set rows = Kosis_ParseMetadata(xml)
    Set row = rows(1)
    If rows.Count <> 1 Or row("code") <> "01" Or row("parent") <> "00" Then Err.Raise 5, , "metadata mismatch"
    If Not Kosis_ValidKey("abcdef123=") Or Kosis_ValidKey("invalid key") Then Err.Raise 5, , "key validation mismatch"
    xml.LoadXML "<response><err>20</err></response>"
    On Error Resume Next
    Set rows = Kosis_ParseMetadata(xml)
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    If n = 0 Then Err.Raise 5, , "error response accepted"
    xml.LoadXML "<m:GenericData xmlns:m='http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message' xmlns:g='http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'><m:DataSet><g:Series><g:SeriesKey><g:Value id='C_A' value='00'/><g:Value id='ITEM' value='T1'/></g:SeriesKey><g:Obs><g:ObsDimension value='2024'/><g:ObsValue value='X'/><g:Attributes><g:Value id='OBS_STATUS' value='p'/></g:Attributes></g:Obs></g:Series></m:DataSet></m:GenericData>"
    Set rows = Kosis_ParseData(xml): Set row = rows(1)
    If row("C_A") <> "00" Or row("period") <> "2024" Or row("value") <> "X" Then Err.Raise 5, , "data code or symbol lost"
    If row("status") <> "OBS_STATUS=p" Then Err.Raise 5, , "status attribute lost"
    Set band = CreateObject("Scripting.Dictionary"): band.Add "name", "young": band.Add "from", 0: band.Add "to", 19: bands.Add band
    Set band = CreateObject("Scripting.Dictionary"): band.Add "name", "adult": band.Add "from", 20: band.Add "to", 100: bands.Add band
    If Not Kosis_ValidateAgeBands(bands, 100, uncovered) Or uncovered <> 0 Then Err.Raise 5, , "valid age bands rejected"
    band("from") = 19
    On Error Resume Next
    Call Kosis_ValidateAgeBands(bands, 100, uncovered)
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    If n = 0 Then Err.Raise 5, , "overlapping age bands accepted"
    band("from") = 21
    If Not Kosis_ValidateAgeBands(bands, 100, uncovered) Or uncovered <> 1 Then Err.Raise 5, , "missing age not detected"
    xml.LoadXML "<response><Structures><prdSe>" & ChrW(&HC6D4) & "</prdSe><prdDe>2024.01</prdDe><prdDe>2024.02</prdDe><prdSe>" & ChrW(&HB144) & "</prdSe><prdDe>2024</prdDe></Structures></response>"
    Set rows = Kosis_ParsePeriods(xml): Set row = rows(1)
    If rows.Count <> 3 Or row("period") <> "202401" Or row("type") <> "M" Then Err.Raise 5, , "period metadata mismatch"
    Set row = rows(3): If row("period") <> "2024" Or row("type") <> "Y" Then Err.Raise 5, , "mixed frequency period mismatch"
    Set options = Kosis_PeriodChoices(rows): Set row = options(1)
    If row("code") <> "M202401" Or row("axis") <> "period" Then Err.Raise 5, , "period picker mismatch"
    xml.LoadXML "<response><Structures><MetaRow><objId>B</objId><itmId>A</itmId><itmNm>Parent</itmNm></MetaRow><MetaRow><objId>B</objId><itmId>A01</itmId><itmNm>Child</itmNm><upItmId>A</upItmId></MetaRow></Structures></response>"
    Set options = Kosis_ParseMetadata(xml): selection.Add "A": selection.Add "A01"
    Set overlaps = Kosis_FindOverlaps(selection, options, "B")
    If overlaps.Count <> 1 Then Err.Raise 5, , "parent child overlap missed"
    xml.LoadXML "<response><Structures><MetaRow><objId>A</objId><itmId>00</itmId><itmNm>National</itmNm></MetaRow><MetaRow><objId>A</objId><itmId>11</itmId><itmNm>Seoul</itmNm></MetaRow></Structures></response>"
    Set options = Kosis_ParseMetadata(xml): Set selection = New Collection: selection.Add "00": selection.Add "11"
    Set overlaps = Kosis_FindOverlaps(selection, options, "A")
    If overlaps.Count <> 1 Then Err.Raise 5, , "national overlap missed"
    TestKosisMetadata = "PASS"
    Exit Function
Failed:
    TestKosisMetadata = CStr(Err.Number) & ": " & Err.Description
End Function
Public Function TestKosisLive(ByVal apiKey As String, ByVal cacheDir As String) As String
    On Error GoTo Failed
    Dim rows As Collection, row As Object, missing As Collection, foundSuppressed As Boolean, errorNumber As Long, stage As String
    Dim axis1 As New Collection, axis2 As New Collection, axis3 As New Collection
    SaveSetting BIZ_APP_NAME, "KOSIS", "ApiKey", apiKey
    stage = "metadata": Set rows = Kosis_ReadMetadata("DT_1K52F03", True, cacheDir)
    If rows.Count < 100 Then Err.Raise 5, , "metadata count too small"
    If Len(Kosis_CacheTimestamp("DT_1K52F03", cacheDir)) = 0 Then Err.Raise 5, , "cache timestamp missing"
    Set rows = Kosis_ReadMetadata("DT_1K52F03", False, cacheDir)
    If rows.Count < 100 Then Err.Raise 5, , "cached metadata missing"
    stage = "year periods": Set rows = Kosis_ReadPeriods("DT_1K52F03", True, cacheDir)
    If rows.Count < 1 Then Err.Raise 5, , "year metadata missing"
    Set row = rows(1): If row("type") <> "Y" Then Err.Raise 5, , "wrong year frequency"
    stage = "monthly periods": Set rows = Kosis_ReadPeriods("DT_1B04006", True, cacheDir)
    If rows.Count < 12 Then Err.Raise 5, , "monthly metadata missing"
    Set row = rows(1): If row("type") <> "M" Then Err.Raise 5, , "wrong month frequency"
    stage = "data": Set rows = Kosis_ReadData("DT_1K52F03", "2024", "Y", "T1", "00", "0", "0")
    If rows.Count <> 1 Then Err.Raise 5, , "unexpected result count"
    Set row = rows(1)
    If row("C_A") <> "00" Or row("ITEM") <> "T1" Or row("period") <> "2024" Or Not IsNumeric(row("value")) Then Err.Raise 5, , "wrong data response"
    stage = "business type data": Set rows = Kosis_ReadData("DT_1K52F01", "2024", "Y", "T1", "00", "0", "0")
    If rows.Count <> 1 Then Err.Raise 5, , "unexpected business type result count"
    Set row = rows(1)
    If row("C_A") <> "00" Or row("C_B") <> "0" Or row("C_C") <> "0" Or Not IsNumeric(row("value")) Then Err.Raise 5, , "wrong business type response"
    stage = "population data": Set rows = Kosis_ReadData("DT_1B04006", "2024", "Y", "T2", "00", "0401", "")
    If rows.Count <> 1 Then Err.Raise 5, , "unexpected population result count"
    Set row = rows(1)
    If row("C_A") <> "00" Or row("C_YRE") <> "0401" Or row("ITEM") <> "T2" Or Not IsNumeric(row("value")) Then Err.Raise 5, , "wrong population response"
    axis1.Add "00": axis2.Add "0": axis2.Add "A": axis3.Add "0": axis3.Add "8"
    stage = "batches": Set rows = Kosis_FetchBatches("DT_1K52F03", "2024", "Y", "T2", axis1, axis2, axis3, "A", "B", "E", missing)
    If rows.Count + missing.Count <> 4 Then Err.Raise 5, , "batch coverage mismatch"
    For Each row In rows
        If row("C_B") = "A" And row("C_E") = "8" And row("value") = "X" Then foundSuppressed = True
    Next row
    If Not foundSuppressed Then Err.Raise 5, , "suppressed value was lost"
    axis2.Add "A"
    On Error Resume Next
    Set rows = Kosis_FetchBatches("DT_1K52F03", "2024", "Y", "T2", axis1, axis2, axis3, "A", "B", "E", missing)
    errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    If errorNumber = 0 Then Err.Raise 5, , "duplicate selection accepted"
    DeleteSetting BIZ_APP_NAME, "KOSIS"
    TestKosisLive = "PASS"
    Exit Function
Failed:
    stage = stage & ": " & Err.Description
    On Error Resume Next
    DeleteSetting BIZ_APP_NAME, "KOSIS"
    TestKosisLive = "KOSIS live request failed: " & stage
End Function
'@
$excel = $null; $book = $null; $cacheDir = ''
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($source + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestKosisMetadata")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'KOSIS metadata parser tests passed.'
    if ($KeyFile) {
        $cacheDir = Join-Path ([IO.Path]::GetTempPath()) ('KosisCacheTest_' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $cacheDir | Out-Null
        $key = (Get-Content -LiteralPath $KeyFile -Raw).Trim()
        $result = [string]$excel.Run("'$($book.Name)'!TestKosisLive", $key, $cacheDir)
        if ($result -ne 'PASS') { throw $result }
        Write-Host 'KOSIS official metadata request passed.'
    }
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    if ($cacheDir) {
        foreach ($name in @('DT_1K52F03_ITM.xml', 'DT_1K52F03_PRD.xml', 'DT_1B04006_PRD.xml')) {
            $path = Join-Path $cacheDir $name
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        }
        if (Test-Path -LiteralPath $cacheDir) { Remove-Item -LiteralPath $cacheDir -Force }
    }
}
