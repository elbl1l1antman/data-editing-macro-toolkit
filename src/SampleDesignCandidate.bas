Attribute VB_Name = "SampleDesignCandidate"
Option Explicit

' Return one record per source cell: row axes, column axes, population, source row, source column.
Public Function SampleCandidate_ReadTable(ByVal rowAxes As Range, ByVal columnAxes As Range, _
                                          ByVal populations As Range) As Collection
    If rowAxes Is Nothing Or columnAxes Is Nothing Or populations Is Nothing Then Err.Raise 5, , "선택하지 않은 범위가 있습니다."
    If Not rowAxes.Worksheet Is populations.Worksheet Or Not columnAxes.Worksheet Is populations.Worksheet Then Err.Raise 5, , "세 범위는 같은 시트에 있어야 합니다."
    If rowAxes.Areas.Count <> 1 Or columnAxes.Areas.Count <> 1 Or populations.Areas.Count <> 1 Then Err.Raise 5, , "각 범위는 떨어지지 않은 한 영역이어야 합니다."
    If rowAxes.Row <> populations.Row Or rowAxes.Rows.Count <> populations.Rows.Count Then Err.Raise 5, , "세로 구분과 모집단 값의 시작 행 또는 행 수가 다릅니다."
    If columnAxes.Column <> populations.Column Or columnAxes.Columns.Count <> populations.Columns.Count Then Err.Raise 5, , "가로 머리글과 모집단 값의 시작 열 또는 열 수가 다릅니다."
    If CDbl(populations.Rows.Count) * CDbl(populations.Columns.Count) > 10000# Then _
        Err.Raise 5, , "모집단 값은 한 번에 10,000셀까지 선택할 수 있습니다. 표를 나누거나 기존 구성표 리스트화 기능을 사용하세요."

    Dim records As New Collection, r As Long, c As Long, axis As Long, index As Long
    Dim labels() As Variant, item() As Variant, value As Variant, rowIsTotal As Boolean, columnIsTotal As Boolean
    ReDim labels(0 To rowAxes.Columns.Count + columnAxes.Rows.Count - 1)
    For r = 1 To populations.Rows.Count
        rowIsTotal = False
        For axis = 1 To rowAxes.Columns.Count
            labels(axis - 1) = SampleCandidate_CellLabel(rowAxes.Cells(r, axis))
            If SampleCandidate_IsTotalLabel(CStr(labels(axis - 1))) Then rowIsTotal = True
        Next axis
        For c = 1 To populations.Columns.Count
            columnIsTotal = False
            For axis = 1 To columnAxes.Rows.Count
                labels(rowAxes.Columns.Count + axis - 1) = SampleCandidate_CellLabel(columnAxes.Cells(axis, c))
                If SampleCandidate_IsTotalLabel(CStr(labels(rowAxes.Columns.Count + axis - 1))) Then columnIsTotal = True
            Next axis
            If Not rowIsTotal And Not columnIsTotal Then
                value = populations.Cells(r, c).Value2
                If IsError(value) Then Err.Raise 5, , "모집단 셀에 수식 오류가 있습니다: " & populations.Cells(r, c).Address(False, False)
                If Len(Trim$(CStr(value))) > 0 Then
                If Not IsNumeric(value) Then Err.Raise 5, , "모집단이 숫자가 아닙니다: " & populations.Cells(r, c).Address(False, False)
                If CDbl(value) < 0 Or CDbl(value) <> Fix(CDbl(value)) Then Err.Raise 5, , "모집단은 0 이상의 정수여야 합니다: " & populations.Cells(r, c).Address(False, False)
                For axis = LBound(labels) To UBound(labels)
                    If Len(labels(axis)) = 0 Then Err.Raise 5, , "구분 이름이 비어 있습니다. 다음 값에 대응하는 행·열을 확인하세요: " & populations.Cells(r, c).Address(False, False)
                Next axis
                ReDim item(0 To UBound(labels) + 3)
                For index = LBound(labels) To UBound(labels)
                    item(index) = labels(index)
                Next index
                item(UBound(labels) + 1) = CDbl(value)
                item(UBound(labels) + 2) = populations.Row + r - 1
                item(UBound(labels) + 3) = populations.Column + c - 1
                    records.Add item
                End If
            End If
        Next c
    Next r
    Set SampleCandidate_ReadTable = records
End Function

Public Function SampleCandidate_GroupKey(ByVal record As Variant, ByVal axisIndexes As Variant) As String
    Dim axis As Variant, label As String, key As String
    For Each axis In axisIndexes
        If CLng(axis) < 0 Or CLng(axis) > UBound(record) - 3 Then Err.Raise 5, , "Invalid group axis"
        label = CStr(record(CLng(axis)))
        key = key & Len(label) & ":" & label
    Next axis
    If Len(key) = 0 Then key = "*"
    SampleCandidate_GroupKey = key
End Function

Public Function SampleCandidate_GroupRecords(ByVal records As Collection, ByVal axisIndexes As Variant) As Collection
    Dim grouped As New Collection, rec As Variant, item() As Variant, i As Long, j As Long
    For i = 1 To records.Count
        rec = records(i)
        ReDim item(0 To UBound(rec) + 1)
        item(0) = SampleCandidate_GroupKey(rec, axisIndexes)
        For j = LBound(rec) To UBound(rec)
            item(j + 1) = rec(j)
        Next j
        grouped.Add item
    Next i
    Set SampleCandidate_GroupRecords = grouped
End Function

Public Function SampleCandidate_Allocate(ByVal records As Collection, ByVal targets As Object, _
    ByVal groupAxis As Long, ByVal method As String, ByVal alpha As Double, _
    ByVal minimum As Long, ByVal maximum As Long) As Collection
    If records Is Nothing Or targets Is Nothing Then Err.Raise 5, , "Missing allocation inputs"
    If records.Count = 0 Or targets.Count = 0 Then Err.Raise 5, , "Empty allocation inputs"
    If minimum < 0 Or maximum < 0 Or (maximum > 0 And minimum > maximum) Then Err.Raise 5, , "최소·최대 표본수 조건을 확인하세요."
    If method <> "proportional" And method <> "sqrt" And method <> "power" And method <> "equal" Then Err.Raise 5, , "Unknown allocation method"
    If method = "power" And alpha < 0 Then Err.Raise 5, , "멱배분 지수는 0 이상이어야 합니다."

    Dim groups As Object, group As Collection, rec As Variant, key As Variant, groupName As String
    Set groups = CreateObject("Scripting.Dictionary")
    Dim i As Long, n As Long, population As Double
    n = records.Count
    Dim weights() As Double, lower() As Long, upper() As Long, ideals() As Double, finals() As Long
    ReDim weights(1 To n): ReDim lower(1 To n): ReDim upper(1 To n)
    ReDim ideals(1 To n): ReDim finals(1 To n)
    For i = 1 To n
        rec = records(i)
        If groupAxis < 0 Or groupAxis > UBound(rec) - 3 Then Err.Raise 5, , "Invalid group axis"
        groupName = CStr(rec(groupAxis))
        If Not targets.Exists(groupName) Then Err.Raise 5, , "No target for group: " & groupName
        If Not groups.Exists(groupName) Then
            Set group = New Collection
            groups.Add groupName, group
        End If
        Set group = groups(groupName)
        group.Add i
        population = CDbl(rec(UBound(rec) - 2))
        If population < 0 Or population <> Fix(population) Or population > 2147483647# Then Err.Raise 5, , "Invalid population"
        If population > 0 Then
            lower(i) = minimum
            upper(i) = CLng(population)
            If maximum > 0 And maximum < upper(i) Then upper(i) = maximum
            If lower(i) > upper(i) Then Err.Raise 5, , "최소 표본수가 모집단 또는 최대 표본수보다 큽니다."
            Select Case method
                Case "proportional": weights(i) = population
                Case "sqrt": weights(i) = Sqr(population)
                Case "power": weights(i) = population ^ alpha
                Case "equal": weights(i) = 1
            End Select
        End If
    Next i

    For Each key In targets.Keys
        If Not groups.Exists(CStr(key)) Then Err.Raise 5, , "Target group has no population: " & CStr(key)
        If Not IsNumeric(targets(key)) Then Err.Raise 5, , "Invalid group target"
        If CDbl(targets(key)) < 0 Or CDbl(targets(key)) <> Fix(CDbl(targets(key))) Or CDbl(targets(key)) > 2147483647# Then Err.Raise 5, , "Invalid group target"
        Set group = groups(CStr(key))
        SampleCandidate_AllocateGroup group, CLng(targets(key)), weights, lower, upper, ideals, finals
    Next key

    Dim output As New Collection, result As Object
    For i = 1 To n
        rec = records(i)
        Set result = CreateObject("Scripting.Dictionary")
        result.Add "source", rec
        result.Add "group", CStr(rec(groupAxis))
        result.Add "target", CLng(targets(CStr(rec(groupAxis))))
        result.Add "weight", weights(i)
        result.Add "ideal", ideals(i)
        result.Add "final", finals(i)
        output.Add result
    Next i
    Set SampleCandidate_Allocate = output
End Function

Private Sub SampleCandidate_AllocateGroup(ByVal group As Collection, ByVal target As Long, _
    ByRef weights() As Double, ByRef lower() As Long, ByRef upper() As Long, _
    ByRef ideals() As Double, ByRef finals() As Long)
    Dim i As Long, idx As Long, sumLower As Double, sumUpper As Double
    Dim lowScale As Double, highScale As Double, middleScale As Double, scaledTotal As Double
    Dim proposed As Double, iteration As Long
    For i = 1 To group.Count
        idx = CLng(group(i))
        sumLower = sumLower + lower(idx): sumUpper = sumUpper + upper(idx)
        If weights(idx) > 0 Then
            If CDbl(upper(idx)) / weights(idx) > highScale Then highScale = CDbl(upper(idx)) / weights(idx)
        End If
    Next i
    If target < sumLower Or target > sumUpper Then Err.Raise 5, , "그룹 목표가 최소·최대 조건 또는 모집단 합계를 벗어납니다. 목표와 제약을 확인하세요."
    ' Solve all lower and upper bounds together, independent of record order.
    For iteration = 1 To 80
        middleScale = (lowScale + highScale) / 2
        scaledTotal = 0
        For i = 1 To group.Count
            idx = CLng(group(i))
            proposed = middleScale * weights(idx)
            If proposed < lower(idx) Then proposed = lower(idx)
            If proposed > upper(idx) Then proposed = upper(idx)
            scaledTotal = scaledTotal + proposed
        Next i
        If scaledTotal <= target Then lowScale = middleScale Else highScale = middleScale
    Next iteration
    For i = 1 To group.Count
        idx = CLng(group(i))
        proposed = lowScale * weights(idx)
        If proposed < lower(idx) Then proposed = lower(idx)
        If proposed > upper(idx) Then proposed = upper(idx)
        ideals(idx) = proposed
        finals(idx) = CLng(Fix(proposed))
    Next i
    Dim allocated As Long, best As Long, bestFraction As Double, fraction As Double
    For i = 1 To group.Count
        idx = CLng(group(i)): allocated = allocated + finals(idx)
    Next i
    Do While allocated < target
        best = 0: bestFraction = -1
        For i = 1 To group.Count
            idx = CLng(group(i))
            If weights(idx) > 0 And finals(idx) < upper(idx) Then
                fraction = ideals(idx) - finals(idx)
                If fraction > bestFraction Then best = idx: bestFraction = fraction
            End If
        Next i
        If best = 0 Then Err.Raise 5, , "Cannot complete allocation"
        finals(best) = finals(best) + 1: allocated = allocated + 1
    Loop
End Sub

Public Sub SampleCandidate_WriteOutput(ByVal outputWs As Worksheet, ByVal rowAxes As Range, _
    ByVal columnAxes As Range, ByVal populations As Range, ByVal allocation As Collection, _
    ByVal method As String, ByVal alpha As Double)
    If outputWs Is populations.Worksheet Then Err.Raise 5, , "Output cannot overwrite the source sheet"
    Dim rowCount As Long, columnCount As Long, rowAxisCount As Long, columnAxisCount As Long
    rowCount = populations.Rows.Count: columnCount = populations.Columns.Count
    rowAxisCount = rowAxes.Columns.Count: columnAxisCount = columnAxes.Rows.Count
    If 5 + columnAxisCount + rowCount + allocation.Count > outputWs.Rows.Count Then Err.Raise 5, , "Output exceeds worksheet rows"
    If 2 * (rowAxisCount + columnCount) + 4 > outputWs.Columns.Count Then Err.Raise 5, , "Output exceeds worksheet columns"

    Dim data() As Variant, populationData() As Variant, r As Long, c As Long, i As Long, j As Long, source As Variant, entry As Object
    ReDim data(1 To columnAxisCount + rowCount, 1 To rowAxisCount + columnCount)
    ReDim populationData(1 To columnAxisCount + rowCount, 1 To rowAxisCount + columnCount)
    For r = 1 To columnAxisCount
        For c = 1 To columnCount
            data(r, rowAxisCount + c) = SampleCandidate_CellLabel(columnAxes.Cells(r, c))
            populationData(r, rowAxisCount + c) = data(r, rowAxisCount + c)
        Next c
    Next r
    For r = 1 To rowCount
        For c = 1 To rowAxisCount
            data(columnAxisCount + r, c) = SampleCandidate_CellLabel(rowAxes.Cells(r, c))
            populationData(columnAxisCount + r, c) = data(columnAxisCount + r, c)
        Next c
    Next r
    For i = 1 To allocation.Count
        Set entry = allocation(i): source = entry("source")
        r = CLng(source(UBound(source) - 1)) - populations.Row + 1
        c = CLng(source(UBound(source))) - populations.Column + 1
        If r < 1 Or r > rowCount Or c < 1 Or c > columnCount Then Err.Raise 5, , "Allocation does not match source range"
        data(columnAxisCount + r, rowAxisCount + c) = CLng(entry("final"))
        populationData(columnAxisCount + r, rowAxisCount + c) = CDbl(source(UBound(source) - 2))
    Next i
    outputWs.Range("A1").Value = "표본설계 결과"
    outputWs.Range("A2").Value = "원본: " & populations.Worksheet.Name & "!" & populations.Address(False, False)
    Dim firstDataRow As Long, lastDataRow As Long, lastValueCol As Long, totalRow As Long
    firstDataRow = 4 + columnAxisCount: lastDataRow = firstDataRow + rowCount - 1
    lastValueCol = rowAxisCount + columnCount: totalRow = lastDataRow + 1
    Dim populationCol As Long, populationLastCol As Long
    populationCol = lastValueCol + 3: populationLastCol = populationCol + lastValueCol - 1
    outputWs.Cells(3, 1).Value = "최종 표본"
    outputWs.Cells(3, populationCol).Value = "모집단"
    outputWs.Cells(4, 1).Resize(columnAxisCount, lastValueCol).NumberFormat = "@"
    outputWs.Cells(firstDataRow, 1).Resize(rowCount, rowAxisCount).NumberFormat = "@"
    outputWs.Cells(4, populationCol).Resize(columnAxisCount, lastValueCol).NumberFormat = "@"
    outputWs.Cells(firstDataRow, populationCol).Resize(rowCount, rowAxisCount).NumberFormat = "@"
    outputWs.Cells(4, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
    outputWs.Cells(4, populationCol).Resize(UBound(populationData, 1), UBound(populationData, 2)).Value2 = populationData
    outputWs.Cells(totalRow, 1).Value = "합계"
    outputWs.Cells(firstDataRow - 1, lastValueCol + 1).Value = "합계"
    outputWs.Cells(totalRow, populationCol).Value = "합계"
    outputWs.Cells(firstDataRow - 1, populationLastCol + 1).Value = "합계"
    For r = firstDataRow To lastDataRow
        outputWs.Cells(r, lastValueCol + 1).FormulaR1C1 = "=SUM(RC" & rowAxisCount + 1 & ":RC" & lastValueCol & ")"
        outputWs.Cells(r, populationLastCol + 1).FormulaR1C1 = "=SUM(RC" & populationCol + rowAxisCount & ":RC" & populationLastCol & ")"
    Next r
    For c = rowAxisCount + 1 To lastValueCol + 1
        outputWs.Cells(totalRow, c).FormulaR1C1 = "=SUM(R" & firstDataRow & "C:R" & lastDataRow & "C)"
    Next c
    For c = populationCol + rowAxisCount To populationLastCol + 1
        outputWs.Cells(totalRow, c).FormulaR1C1 = "=SUM(R" & firstDataRow & "C:R" & lastDataRow & "C)"
    Next c
    With outputWs.Range(outputWs.Cells(4, 1), outputWs.Cells(firstDataRow - 1, lastValueCol + 1))
        .Font.Bold = True: .Interior.Color = RGB(220, 232, 246)
    End With
    outputWs.Range(outputWs.Cells(totalRow, 1), outputWs.Cells(totalRow, lastValueCol + 1)).Font.Bold = True
    outputWs.Range(outputWs.Cells(4, populationCol), outputWs.Cells(firstDataRow - 1, populationLastCol + 1)).Interior.Color = RGB(226, 240, 228)
    outputWs.Range(outputWs.Cells(totalRow, populationCol), outputWs.Cells(totalRow, populationLastCol + 1)).Font.Bold = True

    Dim detailRow As Long, axisCount As Long, popCol As Long, groupCol As Long, targetCol As Long
    Dim weightCol As Long, groupWeightCol As Long, idealCol As Long, finalCol As Long
    Dim codeCol As Long, summaryCol As Long, groupIds As Object, groupTargets As Object, groupId As Long, groupKey As String
    detailRow = totalRow + 4
    If allocation.Count = 0 Then Err.Raise 5, , "No allocation records"
    Set entry = allocation(1): source = entry("source")
    axisCount = UBound(source) - 3
    popCol = axisCount + 1: groupCol = popCol + 1: targetCol = popCol + 2
    weightCol = popCol + 3: groupWeightCol = popCol + 4: idealCol = popCol + 5: finalCol = popCol + 6
    codeCol = finalCol + 3: summaryCol = codeCol + 2
    If detailRow + allocation.Count > outputWs.Rows.Count Or summaryCol + 5 > outputWs.Columns.Count Then Err.Raise 5, , "Output detail exceeds worksheet limits"
    For j = 1 To axisCount
        outputWs.Cells(detailRow, j).Value = "구분" & j
    Next j
    outputWs.Cells(detailRow, popCol).Value = "모집단수"
    outputWs.Cells(detailRow, groupCol).Value = "목표그룹"
    outputWs.Cells(detailRow, targetCol).Value = "그룹 목표"
    outputWs.Cells(detailRow, weightCol).Value = "배분 가중치"
    outputWs.Cells(detailRow, groupWeightCol).Value = "그룹 가중치 합"
    outputWs.Cells(detailRow, idealCol).Value = "이론 배분(제약 전)"
    outputWs.Cells(detailRow, finalCol).Value = "최종 표본수"
    outputWs.Cells(detailRow, finalCol + 1).Value = "원본 행"
    outputWs.Cells(detailRow, finalCol + 2).Value = "원본 열"
    outputWs.Cells(detailRow + 1, 1).Resize(allocation.Count, axisCount).NumberFormat = "@"
    outputWs.Cells(detailRow + 1, groupCol).Resize(allocation.Count, 1).NumberFormat = "@"
    ReDim data(1 To allocation.Count, 1 To codeCol)
    Set groupIds = CreateObject("Scripting.Dictionary")
    Set groupTargets = CreateObject("Scripting.Dictionary")
    For i = 1 To allocation.Count
        Set entry = allocation(i): source = entry("source")
        groupKey = CStr(entry("group"))
        If Not groupIds.Exists(groupKey) Then
            groupIds.Add groupKey, groupIds.Count + 1
            groupTargets.Add groupKey, CLng(entry("target"))
        End If
        For j = 1 To axisCount
            data(i, j) = source(j)
        Next j
        data(i, popCol) = source(UBound(source) - 2)
        data(i, groupCol) = SampleCandidate_GroupLabel(groupKey)
        data(i, targetCol) = entry("target")
        data(i, finalCol) = entry("final")
        data(i, finalCol + 1) = source(UBound(source) - 1)
        data(i, finalCol + 2) = source(UBound(source))
        data(i, codeCol) = groupIds(groupKey)
    Next i
    outputWs.Cells(detailRow + 1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
    Dim tableFormulas() As Variant
    ReDim tableFormulas(1 To rowCount, 1 To columnCount)
    For i = 1 To allocation.Count
        Set entry = allocation(i): source = entry("source")
        r = CLng(source(UBound(source) - 1)) - populations.Row + 1
        c = CLng(source(UBound(source))) - populations.Column + 1
        tableFormulas(r, c) = "=R" & detailRow + i & "C" & finalCol
    Next i
    outputWs.Cells(firstDataRow, rowAxisCount + 1).Resize(rowCount, columnCount).FormulaR1C1 = tableFormulas
    Dim weightFormula As String
    Select Case method
        Case "proportional": weightFormula = "=RC" & popCol
        Case "sqrt": weightFormula = "=SQRT(RC" & popCol & ")"
        Case "power": weightFormula = "=IF(RC" & popCol & "=0,0,RC" & popCol & "^" & Trim$(Str$(alpha)) & ")"
        Case "equal": weightFormula = "=IF(RC" & popCol & ">0,1,0)"
        Case Else: Err.Raise 5, , "Unknown allocation method"
    End Select
    outputWs.Cells(detailRow + 1, weightCol).Resize(allocation.Count, 1).FormulaR1C1 = weightFormula
    Dim firstDetail As Long, lastDetail As Long
    firstDetail = detailRow + 1: lastDetail = detailRow + allocation.Count
    outputWs.Cells(detailRow, summaryCol).Value = "그룹 번호"
    outputWs.Cells(detailRow, summaryCol + 1).Value = "목표그룹"
    outputWs.Cells(detailRow, summaryCol + 2).Value = "그룹 가중치 합"
    outputWs.Cells(detailRow, summaryCol + 3).Value = "목표 표본수"
    outputWs.Cells(detailRow, summaryCol + 4).Value = "최종 표본 합계"
    outputWs.Cells(detailRow, summaryCol + 5).Value = "목표와 차이"
    outputWs.Cells(firstDetail, summaryCol + 1).Resize(groupIds.Count, 1).NumberFormat = "@"
    Dim summaryKey As Variant, groupWeightFormulas() As Variant
    For Each summaryKey In groupIds.Keys
        groupId = CLng(groupIds(summaryKey))
        outputWs.Cells(detailRow + groupId, summaryCol).Value = groupId
        outputWs.Cells(detailRow + groupId, summaryCol + 1).Value = SampleCandidate_GroupLabel(CStr(summaryKey))
        outputWs.Cells(detailRow + groupId, summaryCol + 3).Value = groupTargets(summaryKey)
    Next summaryKey
    outputWs.Cells(firstDetail, summaryCol + 2).Resize(groupIds.Count, 1).FormulaR1C1 = _
        "=SUMIF(R" & firstDetail & "C" & codeCol & ":R" & lastDetail & "C" & codeCol & ",RC" & summaryCol & ",R" & firstDetail & "C" & weightCol & ":R" & lastDetail & "C" & weightCol & ")"
    outputWs.Cells(firstDetail, summaryCol + 4).Resize(groupIds.Count, 1).FormulaR1C1 = _
        "=SUMIF(R" & firstDetail & "C" & codeCol & ":R" & lastDetail & "C" & codeCol & ",RC" & summaryCol & ",R" & firstDetail & "C" & finalCol & ":R" & lastDetail & "C" & finalCol & ")"
    outputWs.Cells(firstDetail, summaryCol + 5).Resize(groupIds.Count, 1).FormulaR1C1 = _
        "=RC" & summaryCol + 4 & "-RC" & summaryCol + 3
    With outputWs.Cells(firstDetail, summaryCol + 5).Resize(groupIds.Count, 1)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 220, 220)
    End With
    ReDim groupWeightFormulas(1 To allocation.Count, 1 To 1)
    For i = 1 To allocation.Count
        Set entry = allocation(i)
        groupId = CLng(groupIds(CStr(entry("group"))))
        groupWeightFormulas(i, 1) = "=R" & detailRow + groupId & "C" & summaryCol + 2
    Next i
    outputWs.Cells(firstDetail, groupWeightCol).Resize(allocation.Count, 1).FormulaR1C1 = groupWeightFormulas
    outputWs.Cells(firstDetail, idealCol).Resize(allocation.Count, 1).FormulaR1C1 = _
        "=IF(RC" & groupWeightCol & "=0,0,RC" & targetCol & "*RC" & weightCol & "/RC" & groupWeightCol & ")"
    With outputWs.Cells(firstDetail, finalCol).Resize(allocation.Count, 1).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="0", Formula2:="=" & outputWs.Cells(firstDetail, popCol).Address(False, False)
        .IgnoreBlank = False
        .ErrorTitle = "표본수 확인"
        .ErrorMessage = "0 이상 모집단수 이하의 정수를 입력하세요."
        .ShowError = True
    End With
    outputWs.Rows(detailRow).Font.Bold = True
    outputWs.Columns(codeCol).Hidden = True
    outputWs.Columns.AutoFit
End Sub

Private Function SampleCandidate_GroupLabel(ByVal groupKey As String) As String
    If groupKey = "*" Then
        SampleCandidate_GroupLabel = "전체"
        Exit Function
    End If
    Dim position As Long, separator As Long, labelLength As Long, label As String
    position = 1
    Do While position <= Len(groupKey)
        separator = InStr(position, groupKey, ":")
        If separator = 0 Or Not IsNumeric(Mid$(groupKey, position, separator - position)) Then Err.Raise 5, , "Invalid group key"
        labelLength = CLng(Mid$(groupKey, position, separator - position))
        label = Mid$(groupKey, separator + 1, labelLength)
        If Len(label) <> labelLength Then Err.Raise 5, , "Invalid group key"
        If Len(SampleCandidate_GroupLabel) > 0 Then SampleCandidate_GroupLabel = SampleCandidate_GroupLabel & " / "
        SampleCandidate_GroupLabel = SampleCandidate_GroupLabel & label
        position = separator + labelLength + 1
    Loop
End Function

Private Function SampleCandidate_IsTotalLabel(ByVal value As String) As Boolean
    Dim label As String
    label = UCase$(Replace(Trim$(value), " ", ""))
    Select Case label
        Case "계", "합계", "총계", "소계", "전체", "TOTAL", "SUBTOTAL"
            SampleCandidate_IsTotalLabel = True
        Case Else
            If Len(label) > 2 Then
                SampleCandidate_IsTotalLabel = (Right$(label, 2) = "소계" Or Right$(label, 2) = "합계" Or Left$(label, 3) = "전체(")
            End If
    End Select
End Function

Private Function SampleCandidate_CellLabel(ByVal cell As Range) As String
    Dim value As Variant
    If cell.MergeCells Then
        value = cell.MergeArea.Cells(1, 1).Value2
    Else
        value = cell.Value2
    End If
    If IsError(value) Then Err.Raise 5, , "구분 셀에 수식 오류가 있습니다: " & cell.Address(False, False)
    SampleCandidate_CellLabel = Trim$(CStr(value))
End Function
