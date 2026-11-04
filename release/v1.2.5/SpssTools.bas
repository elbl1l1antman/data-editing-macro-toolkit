Attribute VB_Name = "SpssTools"
Option Explicit


' ============================================================
' SPSS table banner tools
' 설  명  : SPSS 내보내기 통계표에서 선택한 가로/세로 배너 블록을 삭제하고
'           남은 열/행을 자동으로 당겨 붙인다.
' ============================================================
Public Sub SPSS_DeleteSelectedBanner()
    If TypeName(Selection) <> "Range" Then
        MsgBox "삭제할 배너 셀을 먼저 선택하세요." & vbCrLf & _
               "예: ◐ 지역별, ◐ 성별, ▣ 지역별", vbExclamation, "SPSS 배너 삭제"
        Exit Sub
    End If

    Dim selectedCell As Range
    Set selectedCell = Selection.Cells(1, 1)

    Dim bannerCell As Range
    Set bannerCell = SPSS_FindBannerCell(selectedCell)
    If bannerCell Is Nothing Then
        MsgBox "선택한 셀 주변에서 배너를 찾지 못했습니다." & vbCrLf & vbCrLf & _
               "가로 배너는 '◐ 지역별'처럼 ◐ 표시가 있는 셀을 선택하세요." & vbCrLf & _
               "세로 배너는 '▣ 지역별'처럼 ▣ 표시가 있는 셀을 선택하세요.", _
               vbExclamation, "SPSS 배너 삭제"
        Exit Sub
    End If

    Dim bannerText As String
    bannerText = Trim$(CStr(bannerCell.Value))

    If Left$(bannerText, 1) = "◐" Then
        SPSS_DeleteHorizontalBannerBlocks bannerCell.Worksheet, SPSS_NormalizeBannerName(bannerText)
    ElseIf Left$(bannerText, 1) = "▣" Then
        SPSS_DeleteVerticalBannerBlocks bannerCell.Worksheet, SPSS_NormalizeBannerName(bannerText)
    Else
        MsgBox "지원하는 배너 표식이 아닙니다: " & bannerText, vbExclamation, "SPSS 배너 삭제"
    End If
End Sub

Private Function SPSS_FindBannerCell(ByVal selectedCell As Range) As Range
    Dim textValue As String
    textValue = Trim$(CStr(selectedCell.Value))

    If Left$(textValue, 1) = "◐" Or Left$(textValue, 1) = "▣" Then
        Set SPSS_FindBannerCell = selectedCell
        Exit Function
    End If

    ' 하위 항목 셀을 선택한 경우: 바로 위 같은 열의 가로 배너를 우선 탐색
    Dim r As Long
    For r = selectedCell.Row To Application.WorksheetFunction.Max(1, selectedCell.Row - 5) Step -1
        textValue = Trim$(CStr(selectedCell.Worksheet.Cells(r, selectedCell.Column).Value))
        If Left$(textValue, 1) = "◐" Then
            Set SPSS_FindBannerCell = selectedCell.Worksheet.Cells(r, selectedCell.Column)
            Exit Function
        End If
    Next r

    ' 행 항목을 선택한 경우: 같은 행 왼쪽의 세로 배너를 탐색
    Dim c As Long
    For c = selectedCell.Column To 1 Step -1
        textValue = Trim$(CStr(selectedCell.Worksheet.Cells(selectedCell.Row, c).Value))
        If Left$(textValue, 1) = "▣" Then
            Set SPSS_FindBannerCell = selectedCell.Worksheet.Cells(selectedCell.Row, c)
            Exit Function
        End If
    Next c
End Function

Private Sub SPSS_DeleteHorizontalBannerBlocks(ByVal ws As Worksheet, ByVal targetName As String)
    Dim usedRange As Range
    Set usedRange = ws.UsedRange

    Dim ranges As Object
    Set ranges = CreateObject("Scripting.Dictionary")

    Dim r As Long, c As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        For c = usedRange.Column To usedRange.Column + usedRange.Columns.Count - 1
            Dim cellText As String
            cellText = Trim$(CStr(ws.Cells(r, c).Value))
            If Left$(cellText, 1) = "◐" And SPSS_NormalizeBannerName(cellText) = targetName Then
                Dim endCol As Long
                endCol = SPSS_FindHorizontalBannerEnd(ws, r, c, usedRange.Column + usedRange.Columns.Count - 1)
                ranges(CStr(c) & ":" & CStr(endCol)) = Array(c, endCol)
            End If
        Next c
    Next r

    If ranges.Count = 0 Then
        MsgBox "삭제할 가로 배너를 찾지 못했습니다: " & targetName, vbExclamation, "SPSS 배너 삭제"
        Exit Sub
    End If

    Dim msg As String
    msg = "가로 배너 '" & targetName & "' 열 구간 " & ranges.Count & "개를 삭제합니다." & vbCrLf & _
          "삭제 후 오른쪽 열은 자동으로 당겨 붙습니다." & vbCrLf & vbCrLf & _
          "계속하시겠습니까?"
    If MsgBox(msg, vbQuestion + vbYesNo, "SPSS 배너 삭제") = vbNo Then Exit Sub

    Dim items As Variant
    items = ranges.Items
    SPSS_SortColumnRangesDescending items

    Application.ScreenUpdating = False
    For r = LBound(items) To UBound(items)
        ws.Range(ws.Columns(CLng(items(r)(0))), ws.Columns(CLng(items(r)(1)))).Delete Shift:=xlToLeft
    Next r
    Application.ScreenUpdating = True

    MsgBox "완료: 가로 배너 '" & targetName & "' 삭제", vbInformation, "SPSS 배너 삭제"
End Sub

Private Function SPSS_FindHorizontalBannerEnd(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal startCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = startCol + 1 To lastCol
        Dim textValue As String
        textValue = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If Left$(textValue, 1) = "◐" Then
            SPSS_FindHorizontalBannerEnd = c - 1
            Exit Function
        End If
    Next c
    SPSS_FindHorizontalBannerEnd = lastCol
End Function

Private Sub SPSS_DeleteVerticalBannerBlocks(ByVal ws As Worksheet, ByVal targetName As String)
    Dim usedRange As Range
    Set usedRange = ws.UsedRange

    Dim ranges As Object
    Set ranges = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        Dim cellText As String
        cellText = Trim$(CStr(ws.Cells(r, usedRange.Column).Value))
        If Left$(cellText, 1) = "▣" And SPSS_NormalizeBannerName(cellText) = targetName Then
            Dim endRow As Long
            endRow = SPSS_FindVerticalBannerEnd(ws, r, usedRange.Row + usedRange.Rows.Count - 1, usedRange.Column)
            ranges(CStr(r) & ":" & CStr(endRow)) = Array(r, endRow)
        End If
    Next r

    If ranges.Count = 0 Then
        MsgBox "삭제할 세로 배너를 찾지 못했습니다: " & targetName, vbExclamation, "SPSS 배너 삭제"
        Exit Sub
    End If

    Dim msg As String
    msg = "세로 배너 '" & targetName & "' 행 구간 " & ranges.Count & "개를 삭제합니다." & vbCrLf & _
          "삭제 후 아래 행은 자동으로 당겨 붙습니다." & vbCrLf & vbCrLf & _
          "계속하시겠습니까?"
    If MsgBox(msg, vbQuestion + vbYesNo, "SPSS 배너 삭제") = vbNo Then Exit Sub

    Dim items As Variant
    items = ranges.Items
    SPSS_SortRowRangesDescending items

    For r = LBound(items) To UBound(items)
        If Not Toolkit_CanDeleteRows(ws, CLng(items(r)(0)), CLng(items(r)(1)) - CLng(items(r)(0)) + 1, "SPSS 세로 배너 삭제") Then Exit Sub
    Next r

    Application.ScreenUpdating = False
    For r = LBound(items) To UBound(items)
        ws.Rows(CLng(items(r)(0)) & ":" & CLng(items(r)(1))).Delete Shift:=xlUp
    Next r
    Application.ScreenUpdating = True

    MsgBox "완료: 세로 배너 '" & targetName & "' 삭제", vbInformation, "SPSS 배너 삭제"
End Sub

Private Function SPSS_FindVerticalBannerEnd(ByVal ws As Worksheet, ByVal startRow As Long, ByVal lastRow As Long, ByVal markerCol As Long) As Long
    Dim r As Long
    For r = startRow + 1 To lastRow
        Dim textValue As String
        textValue = Trim$(CStr(ws.Cells(r, markerCol).Value))
        If Left$(textValue, 1) = "▣" Or Left$(textValue, 1) = "[" Then
            SPSS_FindVerticalBannerEnd = r - 1
            Exit Function
        End If
    Next r
    SPSS_FindVerticalBannerEnd = lastRow
End Function

Private Function SPSS_NormalizeBannerName(ByVal bannerText As String) As String
    Dim result As String
    result = Trim$(bannerText)
    If Len(result) > 0 Then
        If Left$(result, 1) = "◐" Or Left$(result, 1) = "▣" Then result = Mid$(result, 2)
    End If
    SPSS_NormalizeBannerName = Trim$(result)
End Function

Private Sub SPSS_SortColumnRangesDescending(ByRef items As Variant)
    Dim i As Long, j As Long
    For i = LBound(items) To UBound(items) - 1
        For j = i + 1 To UBound(items)
            If CLng(items(i)(0)) < CLng(items(j)(0)) Then
                Dim temp As Variant
                temp = items(i)
                items(i) = items(j)
                items(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub SPSS_SortRowRangesDescending(ByRef items As Variant)
    Dim i As Long, j As Long
    For i = LBound(items) To UBound(items) - 1
        For j = i + 1 To UBound(items)
            If CLng(items(i)(0)) < CLng(items(j)(0)) Then
                Dim temp As Variant
                temp = items(i)
                items(i) = items(j)
                items(j) = temp
            End If
        Next j
    Next i
End Sub

Public Sub SPSS_RebuildTablesFromIndex()
    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Or ActiveWorkbook.Worksheets.Count < 1 Then
        MsgBox "집계표 시트가 있는 통합문서에서 실행하세요.", vbExclamation, "SPSS 표 재구성"
        Exit Sub
    End If

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim indexWs As Worksheet
    Dim dataWs As Worksheet
    SPSS_ResolveIndexAndDataSheets wb, indexWs, dataWs

    Dim tables As Collection
    If indexWs Is Nothing Then
        Set tables = SPSS_CollectTablesFromDataSheet(dataWs)
    Else
        Set tables = SPSS_CollectTablesFromIndex(indexWs, dataWs)
    End If
    If tables.Count = 0 Then
        MsgBox "목차 링크 또는 집계표 제목 행에서 표를 찾지 못했습니다.", vbExclamation, "SPSS 표 재구성"
        Exit Sub
    End If

    Dim deleteTables As Object
    Dim deleteTableInput As String
    Dim deleteBannerInput As String
    Dim tableOrderInput As String
    Dim orderInput As String

    Dim banners As Collection
    Set deleteTables = CreateObject("Scripting.Dictionary")
    Set banners = SPSS_CollectUniqueBanners(tables, dataWs, deleteTables)

    If SPSS_ShowRebuildOptionsForm(tables, banners, deleteTableInput, deleteBannerInput, tableOrderInput, orderInput) Then
        Set deleteTables = SPSS_ParseIndexSelection(deleteTableInput, tables.Count)
    Else
        Dim tableList As String
        tableList = SPSS_FormatTableList(tables)

        deleteTableInput = Toolkit_CheckListMenu( _
            "SPSS 표 재구성 - 삭제할 표", _
            "삭제할 표를 선택하세요. 선택하지 않으면 모든 표를 남깁니다.", _
            Toolkit_OptionsFromNumberedPrompt(tableList), "")
        Set deleteTables = SPSS_ParseIndexSelection(deleteTableInput, tables.Count)

        Set banners = SPSS_CollectUniqueBanners(tables, dataWs, deleteTables)
        If banners.Count = 0 Then
            If MsgBox("선택된 표에서 배너를 찾지 못했습니다." & vbCrLf & _
                      "표 삭제만 반영해 신규 표/목차 시트를 생성할까요?", _
                      vbQuestion + vbYesNo, "SPSS 표 재구성") = vbNo Then Exit Sub
        End If

        Dim bannerList As String
        bannerList = SPSS_FormatBannerList(banners)

        If banners.Count > 0 Then
            deleteBannerInput = Toolkit_CheckListMenu( _
                "SPSS 표 재구성 - 삭제할 배너", _
                "삭제할 배너를 선택하세요. 선택하지 않으면 모든 배너를 남깁니다.", _
                Toolkit_OptionsFromNumberedPrompt(bannerList), "")

        End If
    End If

    Dim deleteBanners As Object
    Set deleteBanners = SPSS_ParseBannerDeleteSelection(deleteBannerInput, banners)
    Dim bannerOrder As Collection
    Set bannerOrder = SPSS_ParseBannerOrder(orderInput, banners, deleteBanners)
    Dim tableOrder As Collection
    Set tableOrder = SPSS_ParseTableOrder(tableOrderInput, tables, deleteTables)

    If tableOrder.Count = 0 Then
        MsgBox "남길 표가 없습니다. 삭제할 표 선택을 다시 확인하세요.", vbExclamation, "SPSS 표 재구성"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim outData As Worksheet
    Set outData = wb.Worksheets.Add(After:=dataWs)
    outData.Name = SPSS_UniqueSheetName(wb, "SPSS_표정리")

    Dim dstStartRow As Long: dstStartRow = 1
    Dim orderIdx As Long
    For orderIdx = 1 To tableOrder.Count
        Dim tableNo As Long
        tableNo = CLng(tableOrder(orderIdx))
        SPSS_CopyRebuiltTable dataWs, outData, tables(tableNo), dstStartRow, deleteBanners, bannerOrder
        dstStartRow = dstStartRow + 1
    Next orderIdx

    SPSS_EnsureOneBlankRowBetweenTables outData

    Dim outIndex As Worksheet
    Set outIndex = wb.Worksheets.Add(Before:=outData)
    outIndex.Name = SPSS_UniqueSheetName(wb, "SPSS_목차")
    SPSS_RebuildIndexFromDataSheet outIndex, outData

    outIndex.Columns("A:D").AutoFit
    outData.UsedRange.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "신규 표 시트와 목차 시트를 생성했습니다." & vbCrLf & _
           "목차: " & outIndex.Name & vbCrLf & _
           "표: " & outData.Name, vbInformation, "SPSS 표 재구성 완료"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "SPSS 표 재구성 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "SPSS 표 재구성"
End Sub

Public Sub SPSS_DeleteVerticalBannersFromIndex()
    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Or ActiveWorkbook.Worksheets.Count < 1 Then
        MsgBox "집계표 시트가 있는 통합문서에서 실행하세요.", vbExclamation, "SPSS 세로 배너 삭제"
        Exit Sub
    End If

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim indexWs As Worksheet
    Dim dataWs As Worksheet
    SPSS_ResolveIndexAndDataSheets wb, indexWs, dataWs

    Dim tables As Collection
    If indexWs Is Nothing Then
        Set tables = SPSS_CollectTablesFromDataSheet(dataWs)
    Else
        Set tables = SPSS_CollectTablesFromIndex(indexWs, dataWs)
    End If
    If tables.Count = 0 Then
        MsgBox "목차 링크 또는 집계표 제목 행에서 표를 찾지 못했습니다.", vbExclamation, "SPSS 세로 배너 삭제"
        Exit Sub
    End If

    Dim verticalBanners As Collection
    Set verticalBanners = SPSS_CollectVerticalBanners(tables, dataWs)
    If verticalBanners.Count = 0 Then
        MsgBox "삭제할 세로 배너를 찾지 못했습니다.", vbInformation, "SPSS 세로 배너 삭제"
        Exit Sub
    End If

    Dim deleteInput As String
    deleteInput = Toolkit_CheckListMenu( _
        "SPSS 세로 배너 삭제", _
        "삭제할 세로 배너를 선택하세요. BOT/TOP/MID 계열은 세로 배너로 분류합니다.", _
        Toolkit_OptionsFromNumberedPrompt(SPSS_FormatBannerList(verticalBanners)), "")
    If Len(Trim$(deleteInput)) = 0 Then Exit Sub

    Dim deleteBanners As Object
    Set deleteBanners = SPSS_ParseBannerDeleteSelection(deleteInput, verticalBanners)

    Application.ScreenUpdating = False
    Dim outData As Worksheet
    dataWs.Copy After:=dataWs
    Set outData = ActiveSheet
    outData.Name = SPSS_UniqueSheetName(wb, "SPSS_세로배너삭제")

    SPSS_ApplyVerticalBannerDeletes outData, deleteBanners
    SPSS_ApplyVerticalLikeMarkerDeletes outData, deleteBanners
    SPSS_EnsureOneBlankRowBetweenTables outData

    Dim outIndex As Worksheet
    Set outIndex = wb.Worksheets.Add(Before:=outData)
    outIndex.Name = SPSS_UniqueSheetName(wb, "SPSS_목차")
    SPSS_RebuildIndexFromDataSheet outIndex, outData

    outIndex.Columns("A:D").AutoFit
    outData.UsedRange.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "세로 배너 삭제 결과 시트와 목차를 생성했습니다." & vbCrLf & _
           "목차: " & outIndex.Name & vbCrLf & _
           "표: " & outData.Name, vbInformation, "SPSS 세로 배너 삭제 완료"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "세로 배너 삭제 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "SPSS 세로 배너 삭제"
End Sub

Public Sub SPSS_CreateIndexForActiveSheet()
    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Or TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "목차를 만들 집계표 시트를 먼저 선택하세요.", vbExclamation, "SPSS 목차 생성"
        Exit Sub
    End If

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim dataWs As Worksheet: Set dataWs = ActiveSheet

    Dim tables As Collection
    Set tables = SPSS_CollectTablesFromDataSheet(dataWs)
    If tables.Count = 0 Then
        MsgBox "현재 시트에서 표 제목 행을 찾지 못했습니다." & vbCrLf & _
               "표 제목이 [표 ...]처럼 '['로 시작하는지 확인하세요.", vbExclamation, "SPSS 목차 생성"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim outIndex As Worksheet
    Set outIndex = wb.Worksheets.Add(Before:=dataWs)
    outIndex.Name = SPSS_UniqueSheetName(wb, "SPSS_목차")
    SPSS_RebuildIndexFromDataSheet outIndex, dataWs
    outIndex.Columns("A:D").AutoFit
    Application.ScreenUpdating = True

    MsgBox "목차 시트를 생성했습니다." & vbCrLf & _
           "목차: " & outIndex.Name & vbCrLf & _
           "표: " & dataWs.Name, vbInformation, "SPSS 목차 생성 완료"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "SPSS 목차 생성 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "SPSS 목차 생성"
End Sub

Private Function SPSS_CollectTablesFromIndex(ByVal indexWs As Worksheet, ByVal dataWs As Worksheet) As Collection
    Dim tables As Collection
    Set tables = New Collection

    Dim seenRows As Object
    Set seenRows = CreateObject("Scripting.Dictionary")

    Dim usedRange As Range
    Set usedRange = indexWs.UsedRange

    Dim cell As Range
    For Each cell In usedRange.Cells
        If cell.Hyperlinks.Count > 0 Then
            Dim subAddress As String
            subAddress = cell.Hyperlinks(1).SubAddress
            If Len(subAddress) > 0 Then
                Dim targetRow As Long
                targetRow = SPSS_RowFromSubAddress(dataWs, subAddress)
                If targetRow > 0 And Not seenRows.Exists(CStr(targetRow)) Then
                    seenRows(CStr(targetRow)) = True
                    tables.Add Array(SPSS_FindIndexTitle(indexWs, cell.Row, cell.Column, dataWs.Cells(targetRow, 1).Value), targetRow, 0)
                End If
            End If
        End If
    Next cell

    If tables.Count = 0 Then
        Set tables = SPSS_CollectTablesFromDataSheet(dataWs)
    Else
        SPSS_SortTablesByStartRow tables
        SPSS_SetTableEndRows tables, dataWs.UsedRange.Row + dataWs.UsedRange.Rows.Count - 1
    End If

    Set SPSS_CollectTablesFromIndex = tables
End Function

Private Function SPSS_CollectTablesFromDataSheet(ByVal dataWs As Worksheet) As Collection
    Dim tables As Collection
    Set tables = New Collection

    Dim usedRange As Range
    Set usedRange = dataWs.UsedRange

    Dim r As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        Dim titleText As String
        titleText = Trim$(CStr(dataWs.Cells(r, 1).Value))
        If Left$(titleText, 1) = "[" Then
            tables.Add Array(titleText, r, 0)
        End If
    Next r

    SPSS_SortTablesByStartRow tables
    SPSS_SetTableEndRows tables, usedRange.Row + usedRange.Rows.Count - 1
    Set SPSS_CollectTablesFromDataSheet = tables
End Function

Private Sub SPSS_ResolveIndexAndDataSheets(ByVal wb As Workbook, ByRef indexWs As Worksheet, ByRef dataWs As Worksheet)
    Set indexWs = Nothing
    Set dataWs = Nothing

    If wb.Worksheets.Count >= 2 And SPSS_WorksheetHasHyperlinks(wb.Worksheets(1)) Then
        Set indexWs = wb.Worksheets(1)
        Set dataWs = wb.Worksheets(2)
    ElseIf TypeName(ActiveSheet) = "Worksheet" Then
        Set dataWs = ActiveSheet
    Else
        Set dataWs = wb.Worksheets(1)
    End If
End Sub

Private Function SPSS_WorksheetHasHyperlinks(ByVal ws As Worksheet) As Boolean
    SPSS_WorksheetHasHyperlinks = (ws.Hyperlinks.Count > 0)
End Function

Private Function SPSS_RowFromSubAddress(ByVal dataWs As Worksheet, ByVal subAddress As String) As Long
    Dim addr As String
    addr = subAddress
    If InStr(1, addr, "!", vbTextCompare) > 0 Then addr = Mid$(addr, InStrRev(addr, "!") + 1)
    addr = Replace(addr, "'", "")

    On Error Resume Next
    SPSS_RowFromSubAddress = dataWs.Range(addr).Row
    On Error GoTo 0
End Function

Private Function SPSS_FindIndexTitle(ByVal indexWs As Worksheet, ByVal rowNo As Long, ByVal linkCol As Long, ByVal fallbackTitle As String) As String
    Dim c As Long
    For c = linkCol - 1 To 1 Step -1
        Dim textValue As String
        textValue = Trim$(CStr(indexWs.Cells(rowNo, c).Value))
        If Len(textValue) > 0 And textValue <> "GO" Then
            SPSS_FindIndexTitle = textValue
            Exit Function
        End If
    Next c
    SPSS_FindIndexTitle = Trim$(CStr(fallbackTitle))
End Function

Private Sub SPSS_SortTablesByStartRow(ByRef tables As Collection)
    Dim arr() As Variant
    Dim i As Long, j As Long
    ReDim arr(1 To tables.Count)
    For i = 1 To tables.Count
        arr(i) = tables(i)
    Next i

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)(1)) > CLng(arr(j)(1)) Then
                Dim temp As Variant
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    Set tables = New Collection
    For i = 1 To UBound(arr)
        tables.Add arr(i)
    Next i
End Sub

Private Sub SPSS_SetTableEndRows(ByRef tables As Collection, ByVal lastUsedRow As Long)
    Dim rebuilt As Collection
    Set rebuilt = New Collection

    Dim i As Long
    For i = 1 To tables.Count
        Dim item As Variant
        item = tables(i)
        Dim endRow As Long
        If i < tables.Count Then
            Dim nextItem As Variant
            nextItem = tables(i + 1)
            endRow = CLng(nextItem(1)) - 1
        Else
            endRow = lastUsedRow
        End If
        rebuilt.Add Array(CStr(item(0)), CLng(item(1)), endRow)
    Next i

    Set tables = rebuilt
End Sub

Private Function SPSS_FormatTableList(ByVal tables As Collection) As String
    Dim result As String
    Dim i As Long
    For i = 1 To tables.Count
        Dim item As Variant
        item = tables(i)
        result = result & i & ". " & CStr(item(0)) & "  (행 " & CLng(item(1)) & "-" & CLng(item(2)) & ")" & vbCrLf
    Next i
    SPSS_FormatTableList = result
End Function

Private Function SPSS_ShowRebuildOptionsForm(ByVal tables As Collection, _
                                             ByVal banners As Collection, _
                                             ByRef deleteTableInput As String, _
                                             ByRef deleteBannerInput As String, _
                                             ByRef tableOrderInput As String, _
                                             ByRef orderInput As String) As Boolean
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmSpssRebuildOptions")
    form.InitializeOptions tables, banners
    form.Show

    If form.Cancelled Then
        Unload form
        Err.Raise vbObjectError + 8301, "SPSS_ShowRebuildOptionsForm", "사용자가 작업을 취소했습니다."
    End If

    deleteTableInput = CStr(form.DeleteTablesText)
    deleteBannerInput = CStr(form.DeleteBannersText)
    tableOrderInput = CStr(form.TableOrderText)
    orderInput = CStr(form.BannerOrderText)

    Unload form
    SPSS_ShowRebuildOptionsForm = True
    Exit Function

Fallback:
    If Err.Number = vbObjectError + 8301 Then Err.Raise Err.Number, Err.Source, Err.Description
    Err.Clear
    SPSS_ShowRebuildOptionsForm = False
End Function

Private Function SPSS_CollectUniqueBanners(ByVal tables As Collection, ByVal dataWs As Worksheet, ByVal deleteTables As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To tables.Count
        If Not deleteTables.Exists(CStr(i)) Then
            Dim tableInfo As Variant
            tableInfo = tables(i)
            SPSS_AddTableBanners result, seen, dataWs, CLng(tableInfo(1)), CLng(tableInfo(2))
        End If
    Next i

    Set SPSS_CollectUniqueBanners = result
End Function

Private Function SPSS_CollectVerticalBanners(ByVal tables As Collection, ByVal dataWs As Worksheet) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To tables.Count
        Dim tableInfo As Variant
        tableInfo = tables(i)
        Dim lastCol As Long
        lastCol = SPSS_TableLastCol(dataWs, CLng(tableInfo(1)), CLng(tableInfo(2)))

        Dim r As Long, c As Long
        For r = CLng(tableInfo(1)) To CLng(tableInfo(2))
            Dim firstText As String
            firstText = Trim$(CStr(dataWs.Cells(r, 1).Value))
            If Left$(firstText, 1) = "▣" Then SPSS_AddBannerName result, seen, firstText

            For c = 1 To lastCol
                Dim textValue As String
                textValue = Trim$(CStr(dataWs.Cells(r, c).Value))
                If Left$(textValue, 1) = "◐" And SPSS_IsVerticalBannerLikeName(SPSS_NormalizeBannerName(textValue)) Then
                    SPSS_AddBannerName result, seen, textValue
                End If
            Next c
        Next r
    Next i

    Set SPSS_CollectVerticalBanners = result
End Function

Private Sub SPSS_AddTableBanners(ByVal result As Collection, ByVal seen As Object, ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim lastCol As Long
    lastCol = SPSS_TableLastCol(ws, startRow, endRow)

    Dim r As Long, c As Long
    For r = startRow To endRow
        Dim firstText As String
        firstText = Trim$(CStr(ws.Cells(r, 1).Value))
        If Left$(firstText, 1) = "▣" Then SPSS_AddBannerName result, seen, firstText

        For c = 1 To lastCol
            Dim textValue As String
            textValue = Trim$(CStr(ws.Cells(r, c).Value))
            If Left$(textValue, 1) = "◐" Then SPSS_AddBannerName result, seen, textValue
        Next c
    Next r
End Sub

Private Sub SPSS_AddBannerName(ByVal result As Collection, ByVal seen As Object, ByVal bannerText As String)
    Dim name As String
    name = SPSS_NormalizeBannerName(bannerText)
    If Len(name) = 0 Then Exit Sub
    If Not seen.Exists(name) Then
        seen(name) = True
        result.Add name
    End If
End Sub

Private Function SPSS_FormatBannerList(ByVal banners As Collection) As String
    Dim result As String
    Dim i As Long
    For i = 1 To banners.Count
        result = result & i & ". " & CStr(banners(i)) & vbCrLf
    Next i
    SPSS_FormatBannerList = result
End Function

Public Function SPSS_ParseIndexSelection(ByVal inputText As String, ByVal maxIndex As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    inputText = Replace(Trim$(inputText), " ", "")
    If Len(inputText) = 0 Then
        Set SPSS_ParseIndexSelection = result
        Exit Function
    End If

    Dim parts As Variant
    parts = Split(inputText, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim token As String
        token = CStr(parts(i))
        If InStr(1, token, "-", vbTextCompare) > 0 Then
            Dim pair As Variant
            pair = Split(token, "-")
            If UBound(pair) = 1 And IsNumeric(pair(0)) And IsNumeric(pair(1)) Then
                Dim a As Long: a = CLng(pair(0))
                Dim b As Long: b = CLng(pair(1))
                Dim n As Long
                For n = a To b
                    If n >= 1 And n <= maxIndex Then result(CStr(n)) = True
                Next n
            End If
        ElseIf IsNumeric(token) Then
            Dim idx As Long: idx = CLng(token)
            If idx >= 1 And idx <= maxIndex Then result(CStr(idx)) = True
        End If
    Next i

    Set SPSS_ParseIndexSelection = result
End Function

Private Function SPSS_ParseBannerDeleteSelection(ByVal inputText As String, ByVal banners As Collection) As Object
    Dim indexSelection As Object
    Set indexSelection = SPSS_ParseIndexSelection(inputText, banners.Count)

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In indexSelection.Keys
        result(CStr(banners(CLng(key)))) = True
    Next key

    Set SPSS_ParseBannerDeleteSelection = result
End Function

Private Function SPSS_ParseTableOrder(ByVal inputText As String, ByVal tables As Collection, ByVal deleteTables As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim added As Object
    Set added = CreateObject("Scripting.Dictionary")

    inputText = Replace(Trim$(inputText), " ", "")
    If Len(inputText) > 0 Then
        Dim parts As Variant
        parts = Split(inputText, ",")
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            If IsNumeric(parts(i)) Then
                Dim idx As Long
                idx = CLng(parts(i))
                If idx >= 1 And idx <= tables.Count Then
                    If Not deleteTables.Exists(CStr(idx)) And Not added.Exists(CStr(idx)) Then
                        result.Add idx
                        added(CStr(idx)) = True
                    End If
                End If
            End If
        Next i
    End If

    Dim tableNo As Long
    For tableNo = 1 To tables.Count
        If Not deleteTables.Exists(CStr(tableNo)) And Not added.Exists(CStr(tableNo)) Then
            result.Add tableNo
            added(CStr(tableNo)) = True
        End If
    Next tableNo

    Set SPSS_ParseTableOrder = result
End Function

Private Function SPSS_ParseBannerOrder(ByVal inputText As String, ByVal banners As Collection, ByVal deleteBanners As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim added As Object
    Set added = CreateObject("Scripting.Dictionary")

    Dim indexSelection As Object
    Set indexSelection = SPSS_ParseIndexSelection(inputText, banners.Count)
    Dim parts As Variant
    inputText = Replace(Trim$(inputText), " ", "")
    If Len(inputText) > 0 Then
        parts = Split(inputText, ",")
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            If IsNumeric(parts(i)) Then
                Dim idx As Long: idx = CLng(parts(i))
                If idx >= 1 And idx <= banners.Count Then
                    Dim name As String: name = CStr(banners(idx))
                    If Not deleteBanners.Exists(name) And Not added.Exists(name) Then
                        result.Add name
                        added(name) = True
                    End If
                End If
            End If
        Next i
    End If

    Set SPSS_ParseBannerOrder = result
End Function

Private Sub SPSS_InitRebuiltIndex(ByVal ws As Worksheet)
    ws.Range("A1").Value = "재구성 표 목차"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "번호"
    ws.Range("B2").Value = "표 제목"
    ws.Range("C2").Value = "이동"
    ws.Range("A2:C2").Font.Bold = True
End Sub

Private Sub SPSS_WriteRebuiltIndexRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal title As String, ByVal dataSheetName As String, ByVal targetRow As Long)
    ws.Cells(rowNo, 1).Value = rowNo - 2
    ws.Cells(rowNo, 2).Value = title
    ws.Cells(rowNo, 3).Value = "GO"
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowNo, 3), Address:="", _
                      SubAddress:="'" & dataSheetName & "'!A" & targetRow, _
                      TextToDisplay:="GO"
End Sub

Private Sub SPSS_ApplyTableDeletes(ByVal ws As Worksheet, ByVal tables As Collection, ByVal deleteTables As Object)
    Dim ranges As Collection
    Set ranges = New Collection

    Dim i As Long
    For i = 1 To tables.Count
        If deleteTables.Exists(CStr(i)) Then
            Dim item As Variant
            item = tables(i)
            ranges.Add Array(CLng(item(1)), CLng(item(2)))
        End If
    Next i

    SPSS_DeleteRowRanges ws, ranges
End Sub

Private Sub SPSS_ApplyVerticalBannerDeletes(ByVal ws As Worksheet, ByVal deleteBanners As Object)
    If deleteBanners.Count = 0 Then Exit Sub

    Dim usedRange As Range
    Set usedRange = ws.UsedRange

    Dim ranges As Collection
    Set ranges = New Collection

    Dim r As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        Dim textValue As String
        textValue = Trim$(CStr(ws.Cells(r, usedRange.Column).Value))
        If Left$(textValue, 1) = "▣" Then
            Dim bannerName As String
            bannerName = SPSS_NormalizeBannerName(textValue)
            If deleteBanners.Exists(bannerName) Then
                ranges.Add Array(r, SPSS_FindVerticalBannerEnd(ws, r, usedRange.Row + usedRange.Rows.Count - 1, usedRange.Column))
            End If
        End If
    Next r

    SPSS_DeleteRowRanges ws, ranges
End Sub

Private Sub SPSS_ApplyVerticalLikeMarkerDeletes(ByVal ws As Worksheet, ByVal deleteBanners As Object)
    If deleteBanners.Count = 0 Then Exit Sub

    Dim usedRange As Range
    Set usedRange = ws.UsedRange

    Dim ranges As Collection
    Set ranges = New Collection

    Dim r As Long, c As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        For c = usedRange.Column To usedRange.Column + usedRange.Columns.Count - 1
            Dim textValue As String
            textValue = Trim$(CStr(ws.Cells(r, c).Value))
            If Left$(textValue, 1) = "◐" Then
                Dim bannerName As String
                bannerName = SPSS_NormalizeBannerName(textValue)
                If deleteBanners.Exists(bannerName) And SPSS_IsVerticalBannerLikeName(bannerName) Then
                    ranges.Add Array(r, r)
                    Exit For
                End If
            End If
        Next c
    Next r

    SPSS_DeleteRowRanges ws, ranges
End Sub

Private Sub SPSS_ApplyHorizontalBannerDeletes(ByVal ws As Worksheet, ByVal deleteBanners As Object)
    If deleteBanners.Count = 0 Then Exit Sub

    Dim usedRange As Range
    Set usedRange = ws.UsedRange
    Dim ranges As Object
    Set ranges = CreateObject("Scripting.Dictionary")

    Dim r As Long, c As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        For c = usedRange.Column To usedRange.Column + usedRange.Columns.Count - 1
            Dim textValue As String
            textValue = Trim$(CStr(ws.Cells(r, c).Value))
            If Left$(textValue, 1) = "◐" Then
                Dim bannerName As String
                bannerName = SPSS_NormalizeBannerName(textValue)
                If deleteBanners.Exists(bannerName) And Not SPSS_IsVerticalBannerLikeName(bannerName) Then
                    Dim endCol As Long
                    endCol = SPSS_FindHorizontalBannerEnd(ws, r, c, usedRange.Column + usedRange.Columns.Count - 1)
                    ranges(CStr(c) & ":" & CStr(endCol)) = Array(c, endCol)
                End If
            End If
        Next c
    Next r

    If ranges.Count = 0 Then Exit Sub
    Dim items As Variant
    items = ranges.Items
    SPSS_SortColumnRangesDescending items

    For r = LBound(items) To UBound(items)
        If Not Toolkit_CanDeleteColumns(ws, CLng(items(r)(0)), CLng(items(r)(1)) - CLng(items(r)(0)) + 1, "SPSS 가로 배너 삭제") Then Exit Sub
    Next r

    For r = LBound(items) To UBound(items)
        ws.Range(ws.Columns(CLng(items(r)(0))), ws.Columns(CLng(items(r)(1)))).Delete Shift:=xlToLeft
    Next r
End Sub

Private Sub SPSS_DeleteRowRanges(ByVal ws As Worksheet, ByVal ranges As Collection)
    If ranges.Count = 0 Then Exit Sub

    Dim items() As Variant
    Dim i As Long
    ReDim items(0 To ranges.Count - 1)
    For i = 1 To ranges.Count
        items(i - 1) = ranges(i)
    Next i
    SPSS_SortRowRangesDescending items

    For i = LBound(items) To UBound(items)
        If Not Toolkit_CanDeleteRows(ws, CLng(items(i)(0)), CLng(items(i)(1)) - CLng(items(i)(0)) + 1, "SPSS 표 재구성") Then Exit Sub
    Next i

    For i = LBound(items) To UBound(items)
        ws.Rows(CLng(items(i)(0)) & ":" & CLng(items(i)(1))).Delete Shift:=xlUp
    Next i
End Sub

Private Sub SPSS_EnsureOneBlankRowBetweenTables(ByVal ws As Worksheet)
    Dim usedRange As Range
    Set usedRange = ws.UsedRange

    Dim titleRows As Collection
    Set titleRows = New Collection

    Dim r As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        Dim titleText As String
        titleText = Trim$(CStr(ws.Cells(r, 1).Value))
        If Left$(titleText, 1) = "[" Then titleRows.Add r
    Next r

    If titleRows.Count < 2 Then Exit Sub

    Dim i As Long
    For i = titleRows.Count To 2 Step -1
        Dim titleRow As Long
        titleRow = CLng(titleRows(i))

        Dim blankCount As Long
        Dim scanRow As Long
        scanRow = titleRow - 1
        Do While scanRow > 1 And SPSS_IsRowValueBlank(ws, scanRow)
            blankCount = blankCount + 1
            scanRow = scanRow - 1
        Loop

        If blankCount = 0 Then
            If Not Toolkit_CanInsertRows(ws, titleRow, 1, "SPSS 표 간격 정리") Then Exit Sub
            ws.Rows(titleRow).Insert Shift:=xlDown
        ElseIf blankCount > 1 Then
            If Not Toolkit_CanDeleteRows(ws, titleRow - blankCount, blankCount - 1, "SPSS 표 간격 정리") Then Exit Sub
            ws.Rows((titleRow - blankCount) & ":" & (titleRow - 2)).Delete Shift:=xlUp
        End If
    Next i
End Sub

Private Function SPSS_IsRowValueBlank(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    SPSS_IsRowValueBlank = (Application.WorksheetFunction.CountA(ws.Rows(rowIndex)) = 0)
End Function

Private Sub SPSS_RebuildIndexFromDataSheet(ByVal indexWs As Worksheet, ByVal dataWs As Worksheet)
    SPSS_InitRebuiltIndex indexWs

    Dim indexRow As Long: indexRow = 3
    Dim usedRange As Range
    Set usedRange = dataWs.UsedRange

    Dim r As Long
    For r = usedRange.Row To usedRange.Row + usedRange.Rows.Count - 1
        Dim titleText As String
        titleText = Trim$(CStr(dataWs.Cells(r, 1).Value))
        If Left$(titleText, 1) = "[" Then
            SPSS_WriteRebuiltIndexRow indexWs, indexRow, titleText, dataWs.Name, r
            indexRow = indexRow + 1
        End If
    Next r
End Sub

Private Function SPSS_IsVerticalBannerLikeName(ByVal bannerName As String) As Boolean
    Dim textValue As String
    textValue = UCase$(Trim$(bannerName))
    SPSS_IsVerticalBannerLikeName = (InStr(1, textValue, "BOT", vbTextCompare) > 0 _
                                  Or InStr(1, textValue, "BOTTOM", vbTextCompare) > 0 _
                                  Or InStr(1, textValue, "TOP", vbTextCompare) > 0 _
                                  Or InStr(1, textValue, "MID", vbTextCompare) > 0)
End Function

Private Sub SPSS_CopyRebuiltTable(ByVal srcWs As Worksheet, _
                                  ByVal dstWs As Worksheet, _
                                  ByVal tableInfo As Variant, _
                                  ByRef dstStartRow As Long, _
                                  ByVal deleteBanners As Object, _
                                  ByVal orderNames As Collection)
    Dim startRow As Long: startRow = CLng(tableInfo(1))
    Dim endRow As Long: endRow = CLng(tableInfo(2))
    Dim lastCol As Long: lastCol = SPSS_TableLastCol(srcWs, startRow, endRow)

    Dim rowSegments As Collection
    Set rowSegments = SPSS_BuildRowSegments(srcWs, startRow, endRow)
    Dim colSegments As Collection
    Set colSegments = SPSS_BuildColumnSegments(srcWs, startRow, endRow, lastCol)

    Dim orderedRows As Collection
    Set orderedRows = SPSS_OrderSegments(rowSegments, deleteBanners, orderNames)
    Dim orderedCols As Collection
    Set orderedCols = SPSS_OrderSegments(colSegments, deleteBanners, orderNames)

    Dim dstRow As Long: dstRow = dstStartRow
    Dim rIdx As Long, cIdx As Long
    For rIdx = 1 To orderedRows.Count
        Dim rowSeg As Variant
        rowSeg = orderedRows(rIdx)
        Dim dstCol As Long: dstCol = 1
        For cIdx = 1 To orderedCols.Count
            Dim colSeg As Variant
            colSeg = orderedCols(cIdx)
            srcWs.Range(srcWs.Cells(CLng(rowSeg(0)), CLng(colSeg(0))), srcWs.Cells(CLng(rowSeg(1)), CLng(colSeg(1)))).Copy _
                Destination:=dstWs.Cells(dstRow, dstCol)
            dstCol = dstCol + (CLng(colSeg(1)) - CLng(colSeg(0)) + 1)
        Next cIdx
        dstRow = dstRow + (CLng(rowSeg(1)) - CLng(rowSeg(0)) + 1)
    Next rIdx

    dstStartRow = dstRow
End Sub

Private Function SPSS_TableLastCol(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim lastCol As Long: lastCol = 1
    Dim r As Long
    For r = startRow To endRow
        Dim c As Long
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If c > lastCol Then lastCol = c
    Next r
    SPSS_TableLastCol = lastCol
End Function

Private Function SPSS_BuildRowSegments(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim bannerRows As Collection
    Set bannerRows = New Collection
    Dim r As Long
    For r = startRow To endRow
        If Left$(Trim$(CStr(ws.Cells(r, 1).Value)), 1) = "▣" Then bannerRows.Add r
    Next r

    If bannerRows.Count = 0 Then
        result.Add Array(startRow, endRow, "", "fixed")
        Set SPSS_BuildRowSegments = result
        Exit Function
    End If

    If CLng(bannerRows(1)) > startRow Then result.Add Array(startRow, CLng(bannerRows(1)) - 1, "", "fixed")
    Dim i As Long
    For i = 1 To bannerRows.Count
        Dim segStart As Long: segStart = CLng(bannerRows(i))
        Dim segEnd As Long
        If i < bannerRows.Count Then segEnd = CLng(bannerRows(i + 1)) - 1 Else segEnd = endRow
        result.Add Array(segStart, segEnd, SPSS_NormalizeBannerName(CStr(ws.Cells(segStart, 1).Value)), "banner")
    Next i

    Set SPSS_BuildRowSegments = result
End Function

Private Function SPSS_BuildColumnSegments(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim bannerRow As Long: bannerRow = 0
    Dim starts As Collection
    Set starts = New Collection

    Dim r As Long, c As Long
    For r = startRow To endRow
        Set starts = New Collection
        For c = 1 To lastCol
            If Left$(Trim$(CStr(ws.Cells(r, c).Value)), 1) = "◐" Then starts.Add c
        Next c
        If starts.Count > 0 Then
            bannerRow = r
            Exit For
        End If
    Next r

    If bannerRow = 0 Then
        result.Add Array(1, lastCol, "", "fixed")
        Set SPSS_BuildColumnSegments = result
        Exit Function
    End If

    If CLng(starts(1)) > 1 Then result.Add Array(1, CLng(starts(1)) - 1, "", "fixed")
    Dim i As Long
    For i = 1 To starts.Count
        Dim segStart As Long: segStart = CLng(starts(i))
        Dim segEnd As Long
        If i < starts.Count Then segEnd = CLng(starts(i + 1)) - 1 Else segEnd = lastCol
        result.Add Array(segStart, segEnd, SPSS_NormalizeBannerName(CStr(ws.Cells(bannerRow, segStart).Value)), "banner")
    Next i

    Set SPSS_BuildColumnSegments = result
End Function

Private Function SPSS_OrderSegments(ByVal sourceSegments As Collection, ByVal deleteBanners As Object, ByVal orderNames As Collection) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim added As Object
    Set added = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To sourceSegments.Count
        Dim seg As Variant
        seg = sourceSegments(i)
        If CStr(seg(3)) <> "banner" Then
            result.Add seg
            added(CStr(i)) = True
        End If
    Next i

    Dim orderIdx As Long
    For orderIdx = 1 To orderNames.Count
        Dim targetName As String
        targetName = CStr(orderNames(orderIdx))
        For i = 1 To sourceSegments.Count
            seg = sourceSegments(i)
            If CStr(seg(3)) = "banner" And CStr(seg(2)) = targetName Then
                If Not deleteBanners.Exists(CStr(seg(2))) And Not added.Exists(CStr(i)) Then
                    result.Add seg
                    added(CStr(i)) = True
                End If
            End If
        Next i
    Next orderIdx

    For i = 1 To sourceSegments.Count
        seg = sourceSegments(i)
        If Not added.Exists(CStr(i)) Then
            If CStr(seg(3)) <> "banner" Or Not deleteBanners.Exists(CStr(seg(2))) Then
                result.Add seg
            End If
        End If
    Next i

    Set SPSS_OrderSegments = result
End Function

Public Function SPSS_UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim name As String
    name = Left$(baseName & "_" & Format(Now, "mmdd_hhnn"), 31)
    Dim candidate As String: candidate = name
    Dim n As Long: n = 1
    Do While SPSS_WorksheetExists(wb, candidate)
        n = n + 1
        candidate = Left$(name, 28) & "_" & n
    Loop
    SPSS_UniqueSheetName = candidate
End Function

Private Function SPSS_WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SPSS_WorksheetExists = Not ws Is Nothing
End Function

Public Sub SPSS_ExtractTableMetadata()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Dim indexWs As Worksheet, dataWs As Worksheet
    SPSS_ResolveIndexAndDataSheets wb, indexWs, dataWs
    If dataWs Is Nothing Then
        MsgBox "표 메타를 추출할 집계표 시트를 찾지 못했습니다.", vbExclamation, "SPSS 표 메타"
        Exit Sub
    End If

    Dim tables As Collection
    If indexWs Is Nothing Then Set tables = SPSS_CollectTablesFromDataSheet(dataWs) Else Set tables = SPSS_CollectTablesFromIndex(indexWs, dataWs)
    Dim outWs As Worksheet
    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = SPSS_UniqueSheetName(wb, "SPSS_표메타")
    Roster_WriteArray outWs, 1, 1, Array("순번", "표제목", "시작행", "끝행", "배너목록", "링크")

    Dim i As Long, tableInfo As Variant, banners As Collection
    For i = 1 To tables.Count
        tableInfo = tables(i)
        Set banners = SPSS_CollectBannersForTable(dataWs, CLng(tableInfo(1)), CLng(tableInfo(2)))
        Roster_WriteArray outWs, i + 1, 1, Array(i, CStr(tableInfo(0)), CLng(tableInfo(1)), CLng(tableInfo(2)), SPSS_JoinCollection(banners), "이동")
        outWs.Hyperlinks.Add Anchor:=outWs.Cells(i + 1, 6), Address:="", SubAddress:="'" & dataWs.Name & "'!A" & CLng(tableInfo(1)), TextToDisplay:="이동"
    Next i
    outWs.Columns.AutoFit
    MsgBox "SPSS 표 메타 시트를 생성했습니다.", vbInformation, "SPSS 표 메타"
End Sub

Private Function SPSS_CollectBannersForTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Collection
    Dim result As New Collection, seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = SPSS_TableLastCol(ws, startRow, endRow)
    Dim r As Long, c As Long, textValue As String, name As String
    For r = startRow To endRow
        For c = 1 To lastCol
            textValue = Trim$(CStr(ws.Cells(r, c).Value))
            If Left$(textValue, 1) = "◐" Or Left$(textValue, 1) = "▣" Then
                name = SPSS_NormalizeBannerName(textValue)
                If Len(name) > 0 And Not seen.Exists(name) Then seen(name) = True: result.Add name
            End If
        Next c
    Next r
    Set SPSS_CollectBannersForTable = result
End Function

Private Function SPSS_JoinCollection(ByVal items As Collection) As String
    Dim i As Long, result As String
    For i = 1 To items.Count
        If Len(result) > 0 Then result = result & ", "
        result = result & CStr(items(i))
    Next i
    SPSS_JoinCollection = result
End Function
