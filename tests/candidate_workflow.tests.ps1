param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$core = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidate.bas') -Raw -Encoding UTF8
$ui = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SampleDesignCandidateUI.bas') -Raw -Encoding UTF8
$core = $core -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$ui = $ui -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
$ui = $ui.Replace('Option Explicit', "Option Explicit`r`nPrivate mRangeCall As Long, mCancel As Boolean")
$ui = [regex]::Replace($ui, '(?ms)^Private Function SampleCandidate_PickRange\(.*?^End Function\s*', '')
$ui = $ui.Replace('If Not SampleCandidate_PickRange(', 'If Not Toolkit_GetRangeFromUserForm(')
$ui = $ui.Replace('If Not SampleCandidate_Preview(records,', 'If Not TestPreview(records,')
$ui = $ui.Replace('Application.InputBox(', 'TestInputBox(').Replace('Type:=1', 'argType:=1').Replace('MsgBox', 'TestMsgBox')
$test = @'
Public Function Toolkit_GetRangeFromUserForm(ByVal title As String, ByVal guide As String, ByRef target As Range) As Boolean
    If mCancel Then Exit Function
    mRangeCall = mRangeCall + 1
    Select Case mRangeCall
        Case 1: Set target = ThisWorkbook.Worksheets(1).Range("A3:A4")
        Case 2: Set target = ThisWorkbook.Worksheets(1).Range("B1:E2")
        Case 3: Set target = ThisWorkbook.Worksheets(1).Range("B3:E4")
    End Select
    Toolkit_GetRangeFromUserForm = True
End Function
Public Function Toolkit_CheckListMenu(ByVal title As String, ByVal guide As String, ByVal options As String, ByVal defaults As String) As String
    Toolkit_CheckListMenu = "0"
End Function
Public Function Toolkit_ButtonMenu(ByVal title As String, ByVal guide As String, ByVal options As String, ByVal defaults As String) As String
    Toolkit_ButtonMenu = "1"
End Function
Public Function TestInputBox(ByVal Prompt As String, ByVal Title As String, ByVal defaultValue As Variant, Optional ByVal argType As Variant) As Variant
    TestInputBox = 20
End Function
Public Function TestMsgBox(ByVal Prompt As String, Optional ByVal Buttons As Variant, Optional ByVal Title As Variant) As Long
    TestMsgBox = vbYes
End Function
Private Function TestPreview(ByVal records As Collection, ByVal rowAxes As Range, ByVal columnAxes As Range, _
    ByVal populations As Range, ByVal sumPopulation As Double, ByRef outputName As String) As Boolean
    TestPreview = True
End Function
Public Function TestCandidateWorkflow() As String
    On Error GoTo Failed
    Dim inputWs As Worksheet, outputWs As Worksheet, initialCount As Long
    Set inputWs = ThisWorkbook.Worksheets(1)
    inputWs.Range("B1:C1").Merge: inputWs.Range("B1").Value = "Male"
    inputWs.Range("D1:E1").Merge: inputWs.Range("D1").Value = "Female"
    inputWs.Range("B2").Value = "20s": inputWs.Range("C2").Value = "30s"
    inputWs.Range("D2").Value = "20s": inputWs.Range("E2").Value = "30s"
    inputWs.Range("A3").Value = "Zone A": inputWs.Range("A4").Value = "Zone B"
    inputWs.Range("B3:E3").Value = Array(10, 20, 30, 40)
    inputWs.Range("B4:E4").Value = Array(0, 5, 6, 7)
    initialCount = ThisWorkbook.Worksheets.Count
    mRangeCall = 0: mCancel = True
    SampleCandidate_Start
    If ThisWorkbook.Worksheets.Count <> initialCount Then Err.Raise 5, , "cancel created a worksheet"
    mCancel = False: mRangeCall = 0
    SampleCandidate_Start
    If ThisWorkbook.Worksheets.Count <> initialCount + 1 Then Err.Raise 5, , "workflow did not create output"
    Set outputWs = ActiveSheet
    Application.CalculateFull
    If outputWs.Range("F8").Value2 <> 20 Then Err.Raise 5, , "workflow total differs from target"
    If inputWs.Range("B3").Value2 <> 10 Then Err.Raise 5, , "workflow modified source"
    inputWs.Range("B3").Value2 = -1
    inputWs.Activate
    mRangeCall = 0: Application.ScreenUpdating = False
    SampleCandidate_Start
    If Application.ScreenUpdating Then Err.Raise 5, , "error changed screen updating"
    If ThisWorkbook.Worksheets.Count <> initialCount + 1 Then Err.Raise 5, , "error created output"
    Application.ScreenUpdating = True
    TestCandidateWorkflow = "PASS"
    Exit Function
Failed:
    Application.ScreenUpdating = True
    TestCandidateWorkflow = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $coreModule = $book.VBProject.VBComponents.Add(1)
    $coreModule.CodeModule.AddFromString($core)
    $uiModule = $book.VBProject.VBComponents.Add(1)
    $uiModule.CodeModule.AddFromString($ui + "`r`n" + $test)
    $result = [string]$excel.Run("'$($book.Name)'!TestCandidateWorkflow")
    if ($result -ne 'PASS') { throw $result }
    Write-Host 'Candidate workflow, cancellation and source preservation tests passed.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
