param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SpssTools.bas') -Raw -Encoding UTF8
$source = $source.Replace('MsgBox', 'TestMessage').Replace('InputBox(', 'TestInput(')
$blocks = [ordered]@{}
function Add-Procedure([string]$name) {
 if ($blocks.Contains($name)) { return }
 $start=[regex]::Match($source,"(?m)^(Public|Private)\s+(Sub|Function)\s+$([regex]::Escape($name))\b.*$")
 if(-not $start.Success){throw "Missing procedure: $name"}
 $tail=$source.Substring($start.Index);$end=[regex]::Match($tail,"(?m)^End $($start.Groups[2].Value)\s*$")
 if(-not $end.Success){throw "Missing end: $name"}
 $body=$tail.Substring(0,$end.Index+$end.Length);$blocks[$name]=$body
 $body=[regex]::Replace($body,'"(?:[^"\r\n]|"")*"','""')
 foreach($match in [regex]::Matches($body,'\bSPSS_[A-Za-z0-9_]+\b')){Add-Procedure $match.Value}
}
foreach($name in @('SPSS_DeleteOptionColumns','SPSS_ParseOptionLabels','SPSS_CollectOptionColumnDeletes','SPSS_ApplyOptionColumnDeletes','SPSS_CollectTablesFromDataSheet')){Add-Procedure $name}
$wrapper=@'
Private Function TestMessage(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    mLastMessage = prompt
    TestMessage = mConfirm
End Function
Private Function TestInput(ByVal prompt As String, Optional ByVal title As String = "", Optional ByVal defaultValue As String = "") As String
    mLastDefault = defaultValue
    TestInput = mLabelInput
End Function
Private Function FindSheet(ByVal wb As Workbook, ByVal prefix As String) As Worksheet
    Dim sheet As Worksheet
    For Each sheet In wb.Worksheets
        If Left$(sheet.Name, Len(prefix)) = prefix Then Set FindSheet = sheet
    Next sheet
End Function
Public Function TestSpssOptionColumns() As String
    On Error GoTo Failed
    Dim wb As Workbook, ws As Worksheet, labels As Object
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1): ws.Name = "Sheet1"

    Set labels = SPSS_ParseOptionLabels(" 계, 합계 ,,계" & vbLf & "TOTAL")
    If labels.Count <> 3 Then Err.Raise 513, , "Label parsing count: " & labels.Count
    If Not labels.Exists("total") Then Err.Raise 513, , "Label match should ignore case"

    ' 표 1: 계 열 뒤에 평균 열이 있고 제목은 A:G 병합
    ws.Range("A1:G1").Merge: ws.Range("A1").Value = "[표 1 > 홈페이지 여부]"
    ws.Range("A2").Value = "BASE:전체": ws.Range("D2").Value = "있음": ws.Range("E2").Value = "없음": ws.Range("F2").Value = "계": ws.Range("G2").Value = "▶평균◀(년)"
    ws.Range("C3").Value = "사례수": ws.Range("D3").Value = "%": ws.Range("E3").Value = "%": ws.Range("F3").Value = "%"
    ws.Range("A4").Value = "■ 전체": ws.Range("C4").Value = "(10)": ws.Range("D4").Value = 60: ws.Range("E4").Value = 40: ws.Range("F4").Value = 100: ws.Range("G4").Value = 7.5
    ws.Range("A5").Value = "▣ 지역별": ws.Range("B5").Value = "계": ws.Range("C5").Value = "(10)": ws.Range("D5").Value = 60: ws.Range("E5").Value = 40: ws.Range("F5").Value = 100: ws.Range("G5").Value = 7.5
    ws.Range("F4:F5").NumberFormat = "0.0"

    ' 표 2: 총계/▶합계◀만 있고 계 열은 없음, 가로 병합 머리글 포함
    ws.Range("A8:G8").Merge: ws.Range("A8").Value = "[표 2 > 지국 수]"
    ws.Range("A9").Value = "BASE:일간": ws.Range("D9:E9").Merge: ws.Range("D9").Value = "1) 직영": ws.Range("F9:G9").Merge: ws.Range("F9").Value = "총계"
    ws.Range("C10").Value = "사례수": ws.Range("D10").Value = "▶평균◀(개)": ws.Range("E10").Value = "▶합계◀(개)": ws.Range("F10").Value = "▶평균◀(개)": ws.Range("G10").Value = "▶합계◀(개)"
    ws.Range("A11").Value = "■ 전체": ws.Range("C11").Value = "(5)": ws.Range("D11").Value = 2: ws.Range("E11").Value = 10: ws.Range("F11").Value = 2: ws.Range("G11").Value = 10

    ' 표 3: 가로 배너 병합(D:F) 안에 계가 있고 오른쪽에 다른 배너 병합(G:H)
    ws.Range("A14:H14").Merge: ws.Range("A14").Value = "[표 3 > 배너 표]"
    ws.Range("D15:F15").Merge: ws.Range("D15").Value = "◐ 성별": ws.Range("G15:H15").Merge: ws.Range("G15").Value = "◐ 연령"
    ws.Range("D16").Value = "남": ws.Range("E16").Value = "여": ws.Range("F16").Value = "계": ws.Range("G16").Value = "20대": ws.Range("H16").Value = "30대"
    ws.Range("A17").Value = "■ 전체": ws.Range("C17").Value = "(4)": ws.Range("D17").Value = 50: ws.Range("E17").Value = 50: ws.Range("F17").Value = 100: ws.Range("G17").Value = 25: ws.Range("H17").Value = 75

    ' 표 4: 계 셀이 다음 표 제목행까지 세로 병합되어 있어 건너뛰어야 함
    ws.Range("A20:E20").Merge: ws.Range("A20").Value = "[표 4 > 건너뛰기]"
    ws.Range("D21").Value = "예": ws.Range("E21").Value = "아니오"
    ws.Range("F21:F24").Merge: ws.Range("F21").Value = "계"
    ws.Range("A22").Value = "■ 전체": ws.Range("D22").Value = 30: ws.Range("E22").Value = 70
    ws.Range("A24").Value = "[표 5 > 마지막]"
    ws.Range("D25").Value = "예"
    ws.Range("A26").Value = "■ 전체": ws.Range("D26").Value = 1

    ' 표 6: 계 머리글이 N/% 두 열에 가로 병합되어 있어 한 열만 지우지 않고 건너뛰어야 함
    ws.Range("A29:G29").Merge: ws.Range("A29").Value = "[표 6 > 묶음 머리글]"
    ws.Range("D30:E30").Merge: ws.Range("D30").Value = "있음": ws.Range("F30:G30").Merge: ws.Range("F30").Value = "계"
    ws.Range("D31").Value = "N": ws.Range("E31").Value = "%": ws.Range("F31").Value = "N": ws.Range("G31").Value = "%"
    ws.Range("A32").Value = "■ 전체": ws.Range("D32").Value = 6: ws.Range("E32").Value = 60: ws.Range("F32").Value = 10: ws.Range("G32").Value = 100

    Dim tables As Collection, plan As Collection
    Set tables = SPSS_CollectTablesFromDataSheet(ws)
    If tables.Count <> 6 Then Err.Raise 513, , "Expected 6 tables, got " & tables.Count
    Set plan = SPSS_CollectOptionColumnDeletes(ws, tables, SPSS_ParseOptionLabels("계"))
    If plan.Count <> 4 Then Err.Raise 513, , "Expected 4 planned columns, got " & plan.Count

    mLabelInput = "": mConfirm = vbYes: ws.Activate
    SPSS_DeleteOptionColumns
    If mLastDefault <> "계" Then Err.Raise 513, , "Default label should be 계"
    If wb.Worksheets.Count <> 1 Then Err.Raise 513, , "Blank input should cancel"

    mLabelInput = "계": mConfirm = vbNo: ws.Activate
    SPSS_DeleteOptionColumns
    If wb.Worksheets.Count <> 1 Then Err.Raise 513, , "Declined confirmation should not create sheets"

    mConfirm = vbYes: ws.Activate
    SPSS_DeleteOptionColumns
    Dim outWs As Worksheet, outIndex As Worksheet
    Set outWs = FindSheet(wb, "SPSS_보기삭제")
    Set outIndex = FindSheet(wb, "SPSS_목차")
    If outWs Is Nothing Or outIndex Is Nothing Then Err.Raise 513, , "Output sheets missing: " & mLastMessage

    ' 표 1
    If outWs.Range("F2").Value2 <> "▶평균◀(년)" Or outWs.Range("F4").Value2 <> 7.5 Then Err.Raise 513, , "Columns right of 계 not shifted"
    If Not IsEmpty(outWs.Range("G2").Value2) Then Err.Raise 513, , "Old last column not cleared"
    If outWs.Range("E2").Value2 <> "없음" Or outWs.Range("D4").Value2 <> 60 Then Err.Raise 513, , "Other options changed"
    If outWs.Range("B5").Value2 <> "계" Then Err.Raise 513, , "Row label 계 must remain"
    If outWs.Range("A1").MergeArea.Columns.Count <> 6 Then Err.Raise 513, , "Title merge not shrunk: " & outWs.Range("A1").MergeArea.Address
    ' 표 2
    If outWs.Range("F9").Value2 <> "총계" Or outWs.Range("G10").Value2 <> "▶합계◀(개)" Or outWs.Range("D9").MergeArea.Columns.Count <> 2 Then Err.Raise 513, , "Table without 계 changed"
    If outWs.Range("A8").MergeArea.Columns.Count <> 7 Then Err.Raise 513, , "Unrelated title merge changed"
    ' 표 3
    If outWs.Range("D15").MergeArea.Address(False, False) <> "D15:E15" Or outWs.Range("D15").Value2 <> "◐ 성별" Then Err.Raise 513, , "Banner merge not shrunk: " & outWs.Range("D15").MergeArea.Address
    If outWs.Range("F15").MergeArea.Address(False, False) <> "F15:G15" Or outWs.Range("F15").Value2 <> "◐ 연령" Then Err.Raise 513, , "Right banner merge not shifted"
    If outWs.Range("F16").Value2 <> "20대" Or outWs.Range("G17").Value2 <> 75 Then Err.Raise 513, , "Banner options not shifted"
    ' 표 4
    If outWs.Range("F21").Value2 <> "계" Or outWs.Range("F21").MergeArea.Rows.Count <> 4 Or outWs.Range("E22").Value2 <> 70 Then Err.Raise 513, , "Unsafe merged column should be skipped"
    ' 표 6
    If outWs.Range("F30").Value2 <> "계" Or outWs.Range("F30").MergeArea.Columns.Count <> 2 Or outWs.Range("G32").Value2 <> 100 Then Err.Raise 513, , "Merged option header should be skipped"
    If InStr(mLastMessage, "건너뛴 열 2개") = 0 Or InStr(mLastMessage, "[표 4") = 0 Or InStr(mLastMessage, "[표 6") = 0 Then Err.Raise 513, , "Skip not reported: " & mLastMessage
    If InStr(mLastMessage, "삭제: 표 2개 / 열 2개") = 0 Then Err.Raise 513, , "Deleted table count should exclude skipped tables: " & mLastMessage
    ' 원본과 목차
    If ws.Range("F2").Value2 <> "계" Or ws.Range("G4").Value2 <> 7.5 Or ws.Range("A1").MergeArea.Columns.Count <> 7 Then Err.Raise 513, , "Original sheet changed"
    If outIndex.Hyperlinks.Count <> 6 Then Err.Raise 513, , "Index link count: " & outIndex.Hyperlinks.Count
    If InStr(outIndex.Hyperlinks(3).SubAddress, outWs.Name) = 0 Then Err.Raise 513, , "Index does not point at output sheet"

    Dim beforeCount As Long
    beforeCount = wb.Worksheets.Count
    mLabelInput = "없는보기": ws.Activate
    SPSS_DeleteOptionColumns
    If wb.Worksheets.Count <> beforeCount Or InStr(mLastMessage, "찾지 못했습니다") = 0 Then Err.Raise 513, , "Missing label should stop without output"

    TestSpssOptionColumns = "PASS"
    Exit Function
Failed:
    TestSpssOptionColumns = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
 $book=$excel.Workbooks.Add()
 while($book.Worksheets.Count -gt 1){$book.Worksheets.Item($book.Worksheets.Count).Delete()}
 $module=$book.VBProject.VBComponents.Add(1)
 $module.CodeModule.AddFromString("Option Explicit`r`nPrivate mLastMessage As String`r`nPrivate mLastDefault As String`r`nPrivate mLabelInput As String`r`nPrivate mConfirm As Long`r`n"+($blocks.Values -join "`r`n")+"`r`n"+$wrapper)
 $result=[string]$excel.Run("'$($book.Name)'!TestSpssOptionColumns")
 if($result -ne 'PASS'){throw $result}
 Write-Host 'SPSS option column delete tests passed.'
} finally {if($book){$book.Close($false)};if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}}
