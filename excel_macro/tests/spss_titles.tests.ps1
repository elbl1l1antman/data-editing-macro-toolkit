param([string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path)
$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $ProjectRoot 'src\SpssTools.bas') -Raw -Encoding UTF8
$source = $source.Replace('MsgBox', 'TestMessage')
$source = $source.Replace('Dim startRow As Long: startRow = CLng(tableInfo(1))', "If mFailCopy Then Err.Raise 513, , ""Injected copy failure""`r`n    Dim startRow As Long: startRow = CLng(tableInfo(1))")
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
foreach($name in @('SPSS_CollectTablesFromDataSheet','SPSS_CollectTablesFromIndex','SPSS_RebuildIndexFromDataSheet','SPSS_RebuildTablesFromIndex','SPSS_DeleteVerticalBannersFromIndex','SPSS_ApplyVerticalBannerDeletes','SPSS_ApplyVerticalLikeMarkerDeletes')){Add-Procedure $name}
$blocks['SPSS_ShowRebuildOptionsForm'] = @'
Private Function SPSS_ShowRebuildOptionsForm(ByVal tables As Collection, ByVal banners As Collection, ByRef deleteTableInput As String, ByRef deleteBannerInput As String, ByRef tableOrderInput As String, ByRef orderInput As String) As Boolean
    If mCancelForm Then Err.Raise vbObjectError + 8301, , "Cancelled"
    deleteBannerInput = mDeleteBanners
    orderInput = mOrderText
    SPSS_ShowRebuildOptionsForm = True
End Function
'@
$wrapper=@'
Private Function TestMessage(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    mLastMessage = prompt
    TestMessage = vbYes
End Function
Private Function Toolkit_CheckListMenu(ByVal title As String, ByVal prompt As String, ByVal options As String, ByVal defaults As String) As String
    Toolkit_CheckListMenu = mVerticalSelection
End Function
Private Function Toolkit_OptionsFromNumberedPrompt(ByVal prompt As String) As String
End Function
Private Function Toolkit_CanDeleteRows(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal rowCount As Long, ByVal operation As String) As Boolean
    Toolkit_CanDeleteRows = True
End Function
Private Function Toolkit_CanInsertRows(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal rowCount As Long, ByVal operation As String) As Boolean
    Toolkit_CanInsertRows = True
End Function
Public Function TestSpssTitles() As String
    On Error GoTo Failed
    Dim wb As Workbook, idx As Worksheet, one As Worksheet, two As Worksheet, rebuilt As Worksheet
    Set wb = ThisWorkbook
    Set idx = wb.Worksheets.Add: idx.Name = "목차"
    Set one = wb.Worksheets.Add: one.Name = "표 1"
    Set two = wb.Worksheets.Add: two.Name = "표 2"
    Set rebuilt = wb.Worksheets.Add: rebuilt.Name = "새목차"
    one.Cells(1, 1).Value = "[표 1] 첫 표"
    one.Cells(2, 3).Value = "사례수": one.Cells(2, 4).Value = "비율"
    one.Cells(4, 1).Value = "【 표 A01 】두 번째 표"
    one.Cells(5, 3).Value = "사례수": one.Cells(5, 4).Value = "비율"
    one.Cells(7, 1).Value = "Q01. 세 번째 표"
    one.Cells(8, 3).Value = "사례수": one.Cells(8, 4).Value = "비율"
    one.Cells(10, 1).Value = "<일반현황> 네 번째 표"
    one.Cells(11, 3).Value = "빈도수": one.Cells(11, 4).Value = "비율"
    one.Cells(12, 1).Value = "Q03 데이터 값"
    one.Cells(13, 3).Value = 1: one.Cells(13, 4).Value = 2
    one.Cells(14, 1).Value = "지역": one.Cells(14, 2).Value = "서울": one.Cells(14, 3).Value = 7
    one.Cells(15, 2).Value = "대전": one.Cells(15, 3).Value = 3
    one.Cells(16, 1).Value = "성별": one.Cells(16, 2).Value = "남성": one.Cells(16, 3).Value = 6
    one.Cells(17, 2).Value = "여성": one.Cells(17, 3).Value = 4
    two.Cells(3, 2).Value = "Q3"
    two.Cells(4, 3).Value = "세종": two.Cells(4, 4).Value = "공주"
    two.Cells(5, 2).Value = "전체": two.Cells(5, 3).Value = 10
    two.Cells(5, 4).Value = CVErr(xlErrNA)
    two.Cells(6, 1).Value = "성별": two.Cells(6, 2).Value = "남성": two.Cells(6, 3).Value = 6
    two.Cells(7, 2).Value = "여성": two.Cells(7, 3).Value = 4
    two.Cells(8, 1).Value = "지역": two.Cells(8, 2).Value = "서울": two.Cells(8, 3).Value = 7
    two.Cells(9, 2).Value = "대전": two.Cells(9, 3).Value = 3
    Dim tables As Collection, info As Variant
    Set tables = SPSS_CollectTablesFromDataSheet(one)
    If tables.Count <> 4 Then Err.Raise 513, , "Expected four distinct titles, got " & tables.Count
    info = tables(4): If CLng(info(1)) <> 10 Then Err.Raise 513, , "False title in data row"
    Set tables = SPSS_CollectTablesFromDataSheet(two)
    If tables.Count <> 1 Then Err.Raise 513, , "Second-column title missing"
    info = tables(1): If CLng(info(1)) <> 3 Then Err.Raise 513, , "Second-column title row wrong"
    idx.Cells(1, 2).Value = "[표 1] 첫 표"
    idx.Hyperlinks.Add Anchor:=idx.Cells(1, 2), Address:="", SubAddress:="'표 1'!A1"
    idx.Cells(2, 2).Value = "Q3 지역별 표"
    idx.Hyperlinks.Add Anchor:=idx.Cells(2, 1), Address:="", SubAddress:="'표 2'!B3", TextToDisplay:="바로가기"
    idx.Hyperlinks.Add Anchor:=idx.Cells(3, 1), Address:="", SubAddress:="'표 1'!A1", TextToDisplay:="중복"
    idx.Hyperlinks.Add Anchor:=idx.Cells(4, 1), Address:="", SubAddress:="'없는 시트'!A1", TextToDisplay:="연결 오류"
    If SPSS_CountInvalidIndexLinks(idx) <> 1 Then Err.Raise 513, , "Broken link not detected"
    Set tables = SPSS_CollectTablesFromIndex(idx, one)
    If tables.Count <> 2 Then Err.Raise 513, , "Multi-sheet index count wrong: " & tables.Count
    info = tables(1): If CStr(info(3)) <> "표 1" Or CLng(info(2)) <> 17 Then Err.Raise 513, , "First target wrong"
    info = tables(2): If CStr(info(3)) <> "표 2" Or CLng(info(1)) <> 3 Or CLng(info(2)) <> 9 Or CStr(info(0)) <> "Q3 지역별 표" Then Err.Raise 513, , "Second target wrong"
    If SPSS_RowFromSubAddress(one, "'표 2'!B3") <> 0 Then Err.Raise 513, , "Cross-sheet link accepted as same sheet"
    Dim bannerNames As New Collection, emptySelection As Object, order As Collection
    bannerNames.Add "성별": bannerNames.Add "지역"
    Set emptySelection = CreateObject("Scripting.Dictionary")
    Set order = SPSS_ParseBannerOrder("1,2", bannerNames, emptySelection)
    If order.Count <> 0 Then Err.Raise 513, , "Untouched banner order is not preserved"
    Set order = SPSS_ParseBannerOrder("2,1", bannerNames, emptySelection)
    If order.Count <> 2 Or CStr(order(1)) <> "지역" Then Err.Raise 513, , "Explicit banner reorder ignored"
    SPSS_RebuildIndexFromDataSheet rebuilt, one
    If rebuilt.Hyperlinks.Count <> 4 Then Err.Raise 513, , "New index missing titles"
    If InStr(rebuilt.Hyperlinks(4).SubAddress, "A10") = 0 Then Err.Raise 513, , "New index points at wrong row"
    idx.Move Before:=wb.Worksheets(1)
    idx.Activate
    mOrderText = "1,2"
    SPSS_RebuildTablesFromIndex
    Dim output As Worksheet, outputIndex As Worksheet, target As Range
    Dim sheet As Worksheet
    For Each sheet In wb.Worksheets
        If Left$(sheet.Name, Len("SPSS_표정리")) = "SPSS_표정리" Then Set output = sheet
        If Left$(sheet.Name, Len("SPSS_목차")) = "SPSS_목차" Then Set outputIndex = sheet
    Next sheet
    If output Is Nothing Or outputIndex Is Nothing Then Err.Raise 513, , "Rebuild failed: " & mLastMessage
    If outputIndex.Hyperlinks.Count <> 2 Then Err.Raise 513, , "Rebuilt index count wrong"
    Set target = output.Range(Mid$(outputIndex.Hyperlinks(2).SubAddress, InStrRev(outputIndex.Hyperlinks(2).SubAddress, "!") + 1))
    If target.Offset(0, 1).Value2 <> "Q3" Then Err.Raise 513, , "Rebuilt link misses second sheet title"
    If target.Offset(3, 0).Value2 <> "성별" Or target.Offset(5, 0).Value2 <> "지역" Then Err.Raise 513, , "Unchanged banner order was altered"
    If one.Cells(1, 1).Value2 <> "[표 1] 첫 표" Or two.Cells(3, 2).Value2 <> "Q3" Then Err.Raise 513, , "Original sheet changed"
    mDeleteBanners = "2"
    idx.Activate
    SPSS_RebuildTablesFromIndex
    Dim filtered As Worksheet
    For Each sheet In wb.Worksheets
        If Left$(sheet.Name, Len("SPSS_표정리")) = "SPSS_표정리" Then
            If Not sheet Is output Then Set filtered = sheet
        End If
    Next sheet
    If filtered Is Nothing Then Err.Raise 513, , "Filtered output missing: " & mLastMessage
    If Application.WorksheetFunction.CountIf(filtered.Columns(1), "성별") <> 0 Then Err.Raise 513, , "Plain banner not deleted"
    If Application.WorksheetFunction.CountIf(filtered.Columns(1), "지역") <> 2 Then Err.Raise 513, , "Other banner not preserved"
    If two.Cells(6, 1).Value2 <> "성별" Then Err.Raise 513, , "Source banner changed"
    Dim beforeFailure As Long
    beforeFailure = wb.Worksheets.Count
    mDeleteBanners = "": mFailCopy = True: idx.Activate
    Application.DisplayAlerts = True
    SPSS_RebuildTablesFromIndex
    If Not Application.DisplayAlerts Then Err.Raise 513, , "Display alerts not restored"
    Application.DisplayAlerts = False
    If wb.Worksheets.Count <> beforeFailure Then Err.Raise 513, , "Failed rebuild left output sheets"
    If one.Cells(1, 1).Value2 <> "[표 1] 첫 표" Or two.Cells(6, 1).Value2 <> "성별" Then Err.Raise 513, , "Failed rebuild changed source"
    mFailCopy = False: mCancelForm = True: mLastMessage = "": idx.Activate
    SPSS_RebuildTablesFromIndex
    If wb.Worksheets.Count <> beforeFailure Then Err.Raise 513, , "Cancelled rebuild changed sheets"
    If InStr(mLastMessage, "오류가 발생했습니다") > 0 Then Err.Raise 513, , "Cancellation shown as error"
    mCancelForm = False
    two.Copy After:=two
    Dim copied As Worksheet, selected As Object, oneTable As New Collection
    Set copied = ActiveSheet
    Set selected = CreateObject("Scripting.Dictionary")
    selected.Add "성별", True
    oneTable.Add Array("Q3", 3, 9, copied.Name)
    SPSS_ApplyVerticalBannerDeletes copied, selected, oneTable
    SPSS_ApplyVerticalLikeMarkerDeletes copied, selected
    If Application.WorksheetFunction.CountIf(copied.Columns(1), "성별") <> 0 Then Err.Raise 513, , "Plain vertical banner not deleted"
    If Application.WorksheetFunction.CountIf(copied.Columns(1), "지역") <> 1 Then Err.Raise 513, , "Other vertical banner not preserved"
    mVerticalSelection = "1": two.Activate
    SPSS_DeleteVerticalBannersFromIndex
    Dim verticalOutput As Worksheet
    For Each sheet In wb.Worksheets
        If Left$(sheet.Name, Len("SPSS_세로배너삭제")) = "SPSS_세로배너삭제" Then Set verticalOutput = sheet
    Next sheet
    If verticalOutput Is Nothing Then Err.Raise 513, , "Vertical workflow failed: " & mLastMessage
    If Application.WorksheetFunction.CountIf(verticalOutput.Columns(1), "성별") <> 0 Then Err.Raise 513, , "Vertical workflow kept deleted banner"
    If Application.WorksheetFunction.CountIf(verticalOutput.Columns(1), "지역") <> 1 Then Err.Raise 513, , "Vertical workflow removed another banner"
    If Not IsError(verticalOutput.Cells(5, 4).Value2) Then Err.Raise 513, , "Vertical workflow did not preserve error value"
    If two.Cells(6, 1).Value2 <> "성별" Then Err.Raise 513, , "Vertical workflow changed original"
    Dim reverseIndex As Worksheet
    Set reverseIndex = wb.Worksheets.Add
    reverseIndex.Name = "목차역순"
    reverseIndex.Hyperlinks.Add Anchor:=reverseIndex.Cells(3, 1), Address:="", SubAddress:="'표 1'!A1", TextToDisplay:="늦게 있는 중복"
    reverseIndex.Hyperlinks.Add Anchor:=reverseIndex.Cells(2, 1), Address:="", SubAddress:="'표 2'!B3", TextToDisplay:="두 번째"
    reverseIndex.Hyperlinks.Add Anchor:=reverseIndex.Cells(1, 1), Address:="", SubAddress:="'표 1'!A1", TextToDisplay:="첫 번째"
    Set tables = SPSS_CollectTablesFromIndex(reverseIndex, one)
    info = tables(1)
    If CStr(info(3)) <> "표 1" Then Err.Raise 513, , "Index insertion order differs from worksheet order"
    If CStr(info(0)) <> "첫 번째" Then Err.Raise 513, , "Earlier duplicate title was not preferred"
    TestSpssTitles = "PASS"
    Exit Function
Failed:
    TestSpssTitles = CStr(Err.Number) & ": " & Err.Description
End Function
'@
$excel=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
 $book=$excel.Workbooks.Add();$module=$book.VBProject.VBComponents.Add(1)
 $module.CodeModule.AddFromString("Option Explicit`r`nPrivate mLastMessage As String`r`nPrivate mDeleteBanners As String`r`nPrivate mOrderText As String`r`nPrivate mFailCopy As Boolean`r`nPrivate mCancelForm As Boolean`r`nPrivate mVerticalSelection As String`r`n"+($blocks.Values -join "`r`n")+"`r`n"+$wrapper)
 $result=[string]$excel.Run("'$($book.Name)'!TestSpssTitles")
 if($result -ne 'PASS'){throw $result}
 Write-Host 'SPSS title and multi-sheet link tests passed.'
} finally {if($book){$book.Close($false)};if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}}
