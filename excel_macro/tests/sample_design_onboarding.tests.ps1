param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)

$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleTools.bas') -Raw -Encoding UTF8
$create = [regex]::Match($source, '(?ms)^Public Sub SampleDesign_CreateTemplate\(\).*?^End Sub').Value
if ($create.Contains('SampleDesign_PrepareSheet(')) { throw 'Template creation still overwrites existing sheets.' }
if (-not $source.Contains('vbQuestion + vbYesNo + vbDefaultButton2')) { throw 'Practice overwrite must default to No.' }
# Replace only interactive boundaries; execute the actual sheet creation and cancellation paths.
$source = $source.Replace('Set setup = SampleDesign_GetSetupFromUser()', 'Set setup = TestSetup()').Replace('MsgBox', 'TestMessage')
$blocks = [ordered]@{}
function Add-VbaProcedure([string]$Name) {
    if ($blocks.Contains($Name)) { return }
    $start = [regex]::Match($source, "(?m)^(Public|Private) (Sub|Function) $([regex]::Escape($Name))\b.*$")
    if (-not $start.Success) { throw "Procedure not found: $Name" }
    $tail = $source.Substring($start.Index)
    $end = [regex]::Match($tail, "(?m)^End $($start.Groups[2].Value)\s*$")
    if (-not $end.Success) { throw "Procedure end not found: $Name" }
    $body = $tail.Substring(0, $end.Index + $end.Length)
    $blocks[$Name] = $body
    $withoutStrings = [regex]::Replace($body, '"(?:[^"\r\n]|"")*"', '""')
    foreach ($match in [regex]::Matches($withoutStrings, '\bSampleDesign_[A-Za-z0-9_]+\b')) {
        Add-VbaProcedure $match.Value
    }
}
foreach ($name in @('SampleDesign_CreateTemplate', 'SampleDesign_UseExample', 'SampleDesign_WriteExampleSheet', 'SampleDesign_WritePracticeInput',
    'SampleDesign_WriteTableInputTemplate', 'SampleDesign_ReadInputSchema',
    'SampleDesign_ReadTargetGroups', 'SampleDesign_ReadPopulation', 'SampleDesign_Allocate')) {
    Add-VbaProcedure $name
}

$wrapper = @'
Private Function TestSetup() As Object
    Dim setup As Object
    Set setup = SampleDesign_DefaultSetup()
    setup("inputMode") = "table"
    Set TestSetup = setup
End Function
Private Function TestMessage(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    TestMessage = mPromptAnswer
End Function
Public Function TestSampleOnboarding() As String
    On Error GoTo Failed
    Dim wb As Workbook, example As Worksheet, inputWs As Worksheet, tableWs As Worksheet
    Set wb = ThisWorkbook
    Set example = wb.Worksheets.Add
    example.Name = "표본설계_예시"
    Set inputWs = wb.Worksheets.Add
    inputWs.Name = "표본설계_입력"
    Set tableWs = wb.Worksheets.Add
    tableWs.Name = "표본설계_집계표입력"
    Dim setup As Object, schema As Object
    Set setup = SampleDesign_DefaultSetup()
    setup("targetN") = 245
    Dim headers As Variant
    headers = setup("quotaHeaders")
    headers(0) = "조사산업"
    setup("quotaHeaders") = headers
    SampleDesign_WriteInputTemplate inputWs, setup, False
    If Application.CountA(inputWs.Range("A21:O23")) <> 0 Then Err.Raise 513, , "Live input contains example records"
    If inputWs.Range("B4").Value2 <> 245 Or inputWs.Range("B20").Value2 <> "조사산업" Then Err.Raise 513, , "Custom settings lost"
    SampleDesign_WriteTableInputTemplate tableWs, False
    If Application.Count(tableWs.UsedRange) <> 0 Then Err.Raise 513, , "Live table contains example counts"
    SampleDesign_WriteExampleSheet example
    If example.Range("J21").Value2 <> 1000 Or example.Range("J23").Value2 <> 300 Then Err.Raise 513, , "Example records missing"
    If example.Range("B4").Value2 <> 180 Then Err.Raise 513, , "Practice target invalid"
    If InStr(example.Range("A2").Value2, "연습용") = 0 Then Err.Raise 513, , "Practice warning missing"
    If example.Hyperlinks.Count < 2 Then Err.Raise 513, , "Navigation links missing"
    Dim button As Shape
    Set button = example.Shapes("sdPractice")
    If InStr(button.OnAction, "SampleDesign_UseExample") = 0 Or InStr(button.OnAction, wb.Name) = 0 Then Err.Raise 513, , "Practice action not qualified"
    If button.Left + button.Width > example.Range("G9").Left Then Err.Raise 513, , "Practice button overlaps navigation"
    If button.Top + button.Height > example.Range("A11").Top Then Err.Raise 513, , "Practice button overlaps group table"
    Dim resultWs As Worksheet, countBefore As Long
    Set resultWs = wb.Worksheets.Add
    resultWs.Name = "표본설계_할당결과"
    resultWs.Range("A1").Value2 = "KEEP_RESULT"
    mPromptAnswer = vbNo
    SampleDesign_CreateTemplate
    countBefore = wb.Worksheets.Count
    SampleDesign_CreateTemplate
    If wb.Worksheets.Count <> countBefore Then Err.Raise 513, , "Repeated setup duplicates sheets"
    If resultWs.Range("A1").Value2 <> "KEEP_RESULT" Or inputWs.Range("B4").Value2 <> 245 Then Err.Raise 513, , "Existing design overwritten"
    If InStr(example.Range("G9").Hyperlinks(1).SubAddress, "집계표입력") = 0 Then Err.Raise 513, , "Table navigation wrong"
    SampleDesign_UseExample
    If inputWs.Range("B4").Value2 <> 245 Then Err.Raise 513, , "Cancelled practice changed settings"
    If Application.CountA(inputWs.Range("A21:O23")) <> 0 Then Err.Raise 513, , "Cancelled practice inserted data"
    mPromptAnswer = vbYes
    SampleDesign_UseExample
    If resultWs.Range("A1").Value2 <> "KEEP_RESULT" Then Err.Raise 513, , "Practice changed result sheet"
    SampleDesign_WritePracticeInput inputWs
    If inputWs.Shapes.Count <> 1 Then Err.Raise 513, , "Repeated setup duplicates controls"
    If InStr(inputWs.Range("A2").Value2, "연습자료") = 0 Then Err.Raise 513, , "Practice input warning missing"
    If inputWs.Range("B4").Value2 <> 180 Then Err.Raise 513, , "Practice settings not applied"
    Dim issues As New Collection, records As New Collection, groups As Collection, result As Collection
    Set schema = SampleDesign_ReadInputSchema(inputWs)
    Set groups = SampleDesign_ReadTargetGroups(inputWs, schema, 180, issues)
    SampleDesign_ReadPopulation inputWs, schema, groups, records, issues
    Set result = SampleDesign_Allocate(records, groups, "단순 비례", 0.5, 0, 0, issues)
    If result.Count <> 3 Then Err.Raise 513, , "Practice allocation count wrong"
    If result(1)("final") <> 100 Or result(2)("final") <> 50 Or result(3)("final") <> 30 Then Err.Raise 513, , "Practice allocation differs from guide"
    If issues.Count <> 0 Then Err.Raise 513, , "Unexpected practice warning"
    inputWs.Delete
    tableWs.Delete
    SampleDesign_CreateTemplate
    Set inputWs = wb.Worksheets("표본설계_입력")
    Set tableWs = wb.Worksheets("표본설계_집계표입력")
    If Application.CountA(inputWs.Range("A21:O23")) <> 0 Then Err.Raise 513, , "New input was not blank"
    If Application.Count(tableWs.UsedRange) <> 0 Then Err.Raise 513, , "New table contains demo numbers"
    If resultWs.Range("A1").Value2 <> "KEEP_RESULT" Then Err.Raise 513, , "Creating missing sheets erased results"
    TestSampleOnboarding = "PASS"
    Exit Function
Failed:
    TestSampleOnboarding = CStr(Err.Number) & ": " & Err.Description
End Function
'@

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString("Option Explicit`r`nPrivate mPromptAnswer As Long`r`n" + ($blocks.Values -join "`r`n") + "`r`n" + $wrapper)
    $result = [string]$excel.Run("'$($workbook.Name)'!TestSampleOnboarding")
    if ($result -ne 'PASS') { throw "Sample onboarding failed: $result" }
    Write-Host 'Sample design onboarding sheets and 100/50/30 allocation tests passed.'
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
