Attribute VB_Name = "KosisOutput"
Option Explicit

Public Function Kosis_WriteGroups(ByVal wb As Workbook, ByVal groups As Collection, ByVal metadata As Collection, _
    ByVal rowAxes As Collection, ByVal axisIds As Variant, ByVal title As String) As Collection
    Dim output As Collection, current As Collection, original As Collection, group As Object, source As Worksheet, target As Worksheet, issue As Variant, tableId As String
    Dim oldAlerts As Boolean, oldScreen As Boolean, savedNumber As Long, savedMessage As String
    Dim i As Long, lastRow As Long, nextRow As Long
    oldAlerts = Application.DisplayAlerts: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    If groups Is Nothing Then Err.Raise 5, , "출력할 자료가 없습니다."
    If groups.Count = 0 Then Err.Raise 5, , "출력할 자료가 없습니다."
    Application.ScreenUpdating = False
    For Each group In groups
        Set original = Nothing
        If group.Exists("rawRecords") Then Set original = group("rawRecords")
        tableId = "": If group.Exists("tableId") Then tableId = CStr(group("tableId"))
        Set current = Kosis_WriteSheets(wb, group("records"), metadata, group("missing"), rowAxes, axisIds, title, CStr(group("period")), CStr(group("item")), original, tableId)
        Set target = current(3)
        nextRow = target.Cells(target.Rows.Count, 1).End(xlUp).Row + 1
        If group.Exists("review") Then
            If group("review").Count > 0 And CStr(target.Cells(2, 1).Value2) = "검토 사항 없음" Then nextRow = 2
            For Each issue In group("review")
                target.Cells(nextRow, 1).Value2 = CStr(issue)
                nextRow = nextRow + 1
            Next issue
        End If
        target.Cells(nextRow, 1).Value2 = "조회: " & title & " / " & CStr(group("period")) & " / " & CStr(group("item"))
        If output Is Nothing Then
            Set output = current
        Else
            For i = 1 To 3
                Set source = current(i): Set target = output(i)
                lastRow = source.Cells(source.Rows.Count, 1).End(xlUp).Row
                nextRow = target.Cells(target.Rows.Count, 1).End(xlUp).Row + 1
                If i = 2 Then
                    nextRow = nextRow + 1
                    source.UsedRange.Copy Destination:=target.Cells(nextRow, 1)
                ElseIf lastRow >= 2 Then
                    source.Range(source.Cells(2, 1), source.Cells(lastRow, source.UsedRange.Columns.Count)).Copy Destination:=target.Cells(nextRow, 1)
                End If
            Next i
            Application.DisplayAlerts = False
            For i = 3 To 1 Step -1: current(i).Delete: Next i
            Application.DisplayAlerts = oldAlerts
        End If
        Set current = Nothing
    Next group
    For i = 1 To 3: output(i).Columns.AutoFit: Next i
    Application.ScreenUpdating = oldScreen
    Set Kosis_WriteGroups = output
    Exit Function
Failed:
    savedNumber = Err.Number: savedMessage = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not current Is Nothing Then
        For i = current.Count To 1 Step -1: current(i).Delete: Next i
    End If
    If Not output Is Nothing Then
        For i = output.Count To 1 Step -1: output(i).Delete: Next i
    End If
    Application.DisplayAlerts = oldAlerts: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise savedNumber, , savedMessage
End Function

Public Function Kosis_WriteSheets(ByVal wb As Workbook, ByVal records As Collection, ByVal metadata As Collection, _
    ByVal missing As Collection, ByVal rowAxes As Collection, ByVal axisIds As Variant, _
    ByVal title As String, ByVal period As String, ByVal item As String, Optional ByVal original As Collection = Nothing, Optional ByVal tableId As String = "") As Collection
    Dim labels As Object, units As Object, rowMap As Object, colMap As Object, cellMap As Object, rowLabels As New Collection, colLabels As New Collection
    Dim selected As Object, entry As Object, row As Object, axis As Variant, i As Long, j As Long
    Dim rowKey As String, colKey As String, rowLabel As String, colLabel As String, cellKey As String, value As Variant
    Dim listData() As Variant, grid() As Variant, reviewData() As Variant, reviewCount As Long
    Dim listWs As Worksheet, gridWs As Worksheet, reviewWs As Worksheet, sheets As New Collection
    Dim oldAlerts As Boolean, savedNumber As Long, savedMessage As String, allAxes As Long, totalCells As Double
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    If wb Is Nothing Or records Is Nothing Or metadata Is Nothing Or missing Is Nothing Or rowAxes Is Nothing Then Err.Raise 5, , "출력 자료가 없습니다."
    If wb.ReadOnly Or wb.ProtectStructure Then Err.Raise 5, , "읽기 전용 또는 구조 보호된 문서에는 결과를 만들 수 없습니다."
    allAxes = UBound(axisIds) - LBound(axisIds) + 1
    If allAxes < 2 Or allAxes > 3 Or rowAxes.Count = 0 Or rowAxes.Count >= allAxes Then Err.Raise 5, , "행과 열 구분을 각각 선택하세요."
    If records.Count = 0 Then Err.Raise 5, , "출력할 통계자료가 없습니다."
    If original Is Nothing Then Set original = records
    If original.Count = 0 Then Err.Raise 5, , "원자료가 없습니다."
    Set selected = CreateObject("Scripting.Dictionary")
    For Each axis In rowAxes
        If selected.Exists(CStr(axis)) Then Err.Raise 5, , "행 구분이 중복되었습니다."
        selected.Add CStr(axis), True
    Next axis
    For Each axis In selected.Keys
        Dim found As Boolean: found = False
        For i = LBound(axisIds) To UBound(axisIds)
            If CStr(axisIds(i)) = CStr(axis) Then found = True
        Next i
        If Not found Then Err.Raise 5, , "선택하지 않은 구분이 있습니다."
    Next axis
    Set labels = CreateObject("Scripting.Dictionary")
    Set units = CreateObject("Scripting.Dictionary")
    For Each entry In metadata
        cellKey = Kosis_OutputKey(CStr(entry("axis")), CStr(entry("code")))
        If labels.Exists(cellKey) Then Err.Raise 5, , "메타정보 코드가 중복되었습니다."
        labels.Add cellKey, CStr(entry("name"))
        If entry.Exists("unit") Then units.Add cellKey, CStr(entry("unit"))
    Next entry
    Set rowMap = CreateObject("Scripting.Dictionary"): Set colMap = CreateObject("Scripting.Dictionary"): Set cellMap = CreateObject("Scripting.Dictionary")
    ReDim listData(1 To original.Count, 1 To allAxes * 2 + 10)
    reviewCount = 0: i = 0
    For Each row In original
        i = i + 1
        If CStr(row("period")) <> period Or CStr(row("ITEM")) <> item Then Err.Raise 5, , "서로 다른 시점 또는 항목이 섞였습니다."
        For j = LBound(axisIds) To UBound(axisIds)
            Dim id As String, code As String, label As String, index As Long
            id = CStr(axisIds(j)): index = j - LBound(axisIds)
            If Not row.Exists("C_" & id) Then Err.Raise 5, , "원자료에 분류 코드가 없습니다."
            code = CStr(row("C_" & id)): cellKey = Kosis_OutputKey(id, code)
            If Not labels.Exists(cellKey) Then Err.Raise 5, , "원자료에 대응하는 분류명이 없습니다."
            listData(i, index * 2 + 1) = code: listData(i, index * 2 + 2) = CStr(labels(cellKey))
        Next j
        listData(i, allAxes * 2 + 1) = period: listData(i, allAxes * 2 + 2) = item
        value = Kosis_OutputValue(CStr(row("value")))
        listData(i, allAxes * 2 + 3) = Kosis_CellValue(value)
        If row.Exists("status") Then listData(i, allAxes * 2 + 4) = CStr(row("status"))
        If row.Exists("updated") Then listData(i, allAxes * 2 + 5) = CStr(row("updated"))
        listData(i, allAxes * 2 + 6) = tableId
        If Len(tableId) > 0 Then listData(i, allAxes * 2 + 7) = "https://kosis.kr/statHtml/statHtml.do?orgId=101&tblId=" & tableId
        listData(i, allAxes * 2 + 8) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        cellKey = Kosis_OutputKey("ITEM", item)
        If labels.Exists(cellKey) Then listData(i, allAxes * 2 + 9) = CStr(labels(cellKey))
        If units.Exists(cellKey) Then listData(i, allAxes * 2 + 10) = CStr(units(cellKey))
        If VarType(value) = vbString Then reviewCount = reviewCount + 1
        If row.Exists("status") Then
            If Len(CStr(row("status"))) > 0 Then reviewCount = reviewCount + 1
        End If
    Next row
    i = 0
    For Each row In records
        i = i + 1: rowKey = "": colKey = "": rowLabel = "": colLabel = ""
        If CStr(row("period")) <> period Or CStr(row("ITEM")) <> item Then Err.Raise 5, , "서로 다른 시점 또는 항목이 섞였습니다."
        For j = LBound(axisIds) To UBound(axisIds)
            id = CStr(axisIds(j)): index = j - LBound(axisIds)
            If Not row.Exists("C_" & id) Then Err.Raise 5, , "결과에 분류 코드가 없습니다."
            code = CStr(row("C_" & id)): cellKey = Kosis_OutputKey(id, code)
            If Not labels.Exists(cellKey) Then Err.Raise 5, , "결과에 대응하는 분류명이 없습니다."
            label = code & " " & CStr(labels(cellKey))
            If selected.Exists(id) Then
                rowKey = rowKey & Kosis_OutputKey(id, code)
                If Len(rowLabel) > 0 Then rowLabel = rowLabel & " / "
                rowLabel = rowLabel & label
            Else
                colKey = colKey & Kosis_OutputKey(id, code)
                If Len(colLabel) > 0 Then colLabel = colLabel & " / "
                colLabel = colLabel & label
            End If
        Next j
        value = Kosis_OutputValue(CStr(row("value")))
        If Not rowMap.Exists(rowKey) Then rowMap.Add rowKey, rowMap.Count + 1: rowLabels.Add rowLabel
        If Not colMap.Exists(colKey) Then colMap.Add colKey, colMap.Count + 1: colLabels.Add colLabel
        cellKey = Kosis_OutputKey(rowKey, colKey)
        If cellMap.Exists(cellKey) Then Err.Raise 5, , "행렬 변환 중 동일한 조합이 중복되었습니다."
        cellMap.Add cellKey, value
        If Not original Is records Then
            If VarType(value) = vbString Then reviewCount = reviewCount + 1
        End If
    Next row
    totalCells = CDbl(rowMap.Count) * CDbl(colMap.Count)
    If rowMap.Count + 5 > wb.Worksheets(1).Rows.Count Or colMap.Count + 1 > wb.Worksheets(1).Columns.Count Or totalCells > 1000000# Then Err.Raise 5, , "행렬표가 너무 큽니다. 조건을 나누어 조회하세요."
    ReDim grid(1 To rowMap.Count + 1, 1 To colMap.Count + 1)
    grid(1, 1) = "구분"
    For i = 1 To rowLabels.Count: grid(i + 1, 1) = rowLabels(i): Next i
    For j = 1 To colLabels.Count: grid(1, j + 1) = colLabels(j): Next j
    For Each axis In rowMap.Keys
        For Each value In colMap.Keys
            cellKey = Kosis_OutputKey(CStr(axis), CStr(value))
            If cellMap.Exists(cellKey) Then
                grid(CLng(rowMap(axis)) + 1, CLng(colMap(value)) + 1) = Kosis_CellValue(cellMap(cellKey))
            Else
                grid(CLng(rowMap(axis)) + 1, CLng(colMap(value)) + 1) = Kosis_CellValue("자료 없음")
                reviewCount = reviewCount + 1
            End If
        Next value
    Next axis
    Set listWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): listWs.Name = Kosis_OutputName(wb, "KOSIS_리스트")
    Set gridWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): gridWs.Name = Kosis_OutputName(wb, "KOSIS_집계표")
    Set reviewWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): reviewWs.Name = Kosis_OutputName(wb, "KOSIS_검토")
    For j = LBound(axisIds) To UBound(axisIds)
        index = j - LBound(axisIds): listWs.Cells(1, index * 2 + 1).Value2 = CStr(axisIds(j)) & " 코드"
        listWs.Cells(1, index * 2 + 2).Value2 = CStr(axisIds(j)) & " 명칭"
    Next j
    listWs.Cells(1, allAxes * 2 + 1).Resize(1, 10).Value = Array("시점", "항목코드", "값", "상태부호", "최종수정일", "통계표ID", "출처", "조회일시", "항목명", "단위")
    ' 코드·명칭·시점은 앞자리 0을 지키도록 텍스트로 두고, 값 열만 숫자로 저장한다.
    listWs.Range(listWs.Cells(2, 1), listWs.Cells(original.Count + 1, allAxes * 2 + 10)).NumberFormat = "@"
    listWs.Range(listWs.Cells(2, allAxes * 2 + 3), listWs.Cells(original.Count + 1, allAxes * 2 + 3)).NumberFormat = "General"
    listWs.Cells(2, 1).Resize(original.Count, allAxes * 2 + 10).Value2 = listData
    gridWs.Range("A1:A3").NumberFormat = "@"
    gridWs.Range("A1").Value2 = title: gridWs.Range("A2").Value2 = period: gridWs.Range("A3").Value2 = item
    ' 행·열 머리글만 텍스트로 두고 본문 값은 숫자로 저장한다.
    gridWs.Cells(5, 1).Resize(1, UBound(grid, 2)).NumberFormat = "@"
    gridWs.Cells(5, 1).Resize(UBound(grid, 1), 1).NumberFormat = "@"
    gridWs.Cells(5, 1).Resize(UBound(grid, 1), UBound(grid, 2)).Value2 = grid
    reviewWs.Range("A1").Value2 = "검토 사항"
    ReDim reviewData(1 To IIf(reviewCount + missing.Count > 0, reviewCount + missing.Count, 1), 1 To 1)
    i = 0
    j = 1
    For Each row In original
        j = j + 1
        If VarType(Kosis_OutputValue(CStr(row("value")))) = vbString Then
            i = i + 1: reviewData(i, 1) = "리스트 " & CStr(j) & "행: 통계부호 또는 숫자가 아닌 값 " & CStr(row("value"))
        End If
        If row.Exists("status") Then
            If Len(CStr(row("status"))) > 0 Then
                i = i + 1: reviewData(i, 1) = "리스트 " & CStr(j) & "행: 상태부호 " & CStr(row("status"))
            End If
        End If
    Next row
    If Not original Is records Then
        For Each row In records
            If VarType(Kosis_OutputValue(CStr(row("value")))) = vbString Then
                i = i + 1: reviewData(i, 1) = "집계 결과 확인 필요: " & CStr(row("value"))
            End If
        Next row
    End If
    For Each axis In rowMap.Keys
        For Each value In colMap.Keys
            If Not cellMap.Exists(Kosis_OutputKey(CStr(axis), CStr(value))) Then
                i = i + 1: reviewData(i, 1) = "자료 없음: " & rowLabels(CLng(rowMap(axis))) & " / " & colLabels(CLng(colMap(value)))
            End If
        Next value
    Next axis
    For Each value In missing
        i = i + 1: reviewData(i, 1) = "요청한 조합 누락: " & Join(value, " / ")
    Next value
    If i = 0 Then reviewData(1, 1) = "검토 사항 없음"
    reviewWs.Cells(2, 1).Resize(UBound(reviewData, 1), 1).Value2 = reviewData
    listWs.Rows(1).Font.Bold = True: gridWs.Rows(5).Font.Bold = True: reviewWs.Rows(1).Font.Bold = True
    listWs.Columns.AutoFit: gridWs.Columns.AutoFit: reviewWs.Columns.AutoFit
    sheets.Add listWs: sheets.Add gridWs: sheets.Add reviewWs
    Set Kosis_WriteSheets = sheets
    Exit Function
Failed:
    savedNumber = Err.Number: savedMessage = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not reviewWs Is Nothing Then reviewWs.Delete
    If Not gridWs Is Nothing Then gridWs.Delete
    If Not listWs Is Nothing Then listWs.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Err.Raise savedNumber, , savedMessage
End Function

Private Function Kosis_OutputValue(ByVal text As String) As Variant
    Dim number As Double
    If IsNumeric(text) Then
        number = CDbl(text)
        If Abs(number) <= 9007199254740991# Then Kosis_OutputValue = number: Exit Function
    End If
    Kosis_OutputValue = text
End Function

' 숫자는 그대로, 통계부호 등 문자열은 앞에 '를 붙여 일반 서식 셀에서도 텍스트로 저장되게 한다.
Private Function Kosis_CellValue(ByVal value As Variant) As Variant
    If VarType(value) = vbString Then Kosis_CellValue = "'" & value Else Kosis_CellValue = value
End Function

Private Function Kosis_OutputKey(ByVal first As String, ByVal second As String) As String
    Kosis_OutputKey = Len(first) & ":" & first & Len(second) & ":" & second
End Function

Private Function Kosis_OutputName(ByVal wb As Workbook, ByVal base As String) As String
    Dim i As Long, candidate As String, ws As Object
    candidate = base: i = 1
    Do
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Sheets(candidate)
        On Error GoTo 0
        If ws Is Nothing Then Kosis_OutputName = candidate: Exit Function
        i = i + 1: candidate = Left$(base, 24) & "_" & CStr(i)
    Loop
End Function
