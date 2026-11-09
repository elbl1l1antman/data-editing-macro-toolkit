Attribute VB_Name = "AddressTools"
Option Explicit


Public Sub Address_ConvertOldToRoad()
    Dim addressRange As Range
    If Not Toolkit_GetRangeFromUserForm("도로명주소 변환 - 원본 주소", _
        "변환할 구주소 범위를 선택하세요. 결과는 오른쪽에 우편번호, 도로명주소, 변환결과 열로 생성됩니다.", _
        addressRange) Then Exit Sub

    Dim mapRange As Range
    Set mapRange = Address_FindMappingRange(ActiveWorkbook)
    If mapRange Is Nothing Then
        If Not Toolkit_GetRangeFromUserForm("도로명주소 변환 - 변환표", _
            "구주소, 도로명주소, 우편번호 헤더가 포함된 변환표 범위를 선택하세요.", _
            mapRange) Then Exit Sub
    End If

    Address_ConvertRangeWithMapping addressRange, mapRange
End Sub

Public Sub Address_ClassifyAddressType()
    Dim addressRange As Range
    If Not Toolkit_GetRangeFromUserForm("주소 유형 판별", _
        "주소가 들어 있는 범위를 선택하세요. 결과는 원본 오른쪽에 도로명주소/법정·행정동 기준 주소 여부로 표시됩니다.", _
        addressRange) Then Exit Sub

    Address_ClassifyRangeType addressRange
End Sub

Public Sub Address_ConvertProvinceNotation()
    Dim addressRange As Range
    If Not Toolkit_GetRangeFromUserForm("광역시도 표기 변환", _
        "주소가 들어 있는 한 개 열을 선택하세요. 첫 어절의 광역시도만 변환하며 결과는 오른쪽 새 열에 생성됩니다.", _
        addressRange) Then Exit Sub

    Dim direction As String
    direction = Toolkit_ButtonMenu( _
        "광역시도 표기 변환", _
        "변환 방향을 선택하세요. 주소 중간에 같은 지명이 있어도 첫 어절만 바뀝니다.", _
        "formal" & vbTab & "축약형 → 정규형" & vbLf & _
        "short" & vbTab & "정규형 → 축약형", _
        "formal")
    If Len(direction) = 0 Then Exit Sub

    Dim formalStyle As String
    formalStyle = "latest"
    If direction = "formal" Then
        formalStyle = Toolkit_ButtonMenu( _
            "정규형 명칭 선택", _
            "강원·전북·제주의 정규형 표기를 선택하세요. 두 형식 모두 입력값으로 인정됩니다.", _
            "latest" & vbTab & "최신 명칭: 강원특별자치도·전북특별자치도·제주특별자치도" & vbLf & _
            "legacy" & vbTab & "기존 명칭: 강원도·전라북도·제주도", _
            "latest")
        If Len(formalStyle) = 0 Then Exit Sub
    End If

    Address_ConvertProvinceRange addressRange, direction, formalStyle
End Sub

Private Sub Address_ConvertProvinceRange(ByVal addressRange As Range, ByVal direction As String, ByVal formalStyle As String)
    If addressRange.Areas.Count > 1 Or addressRange.Columns.Count <> 1 Then
        MsgBox "광역시도 표기 변환은 한 번에 한 개의 연속된 주소 열만 처리합니다.", vbExclamation, "광역시도 표기 변환"
        Exit Sub
    End If

    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(addressRange)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 변환할 주소를 찾지 못했습니다.", vbInformation, "광역시도 표기 변환"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = workRange.Worksheet

    Dim firstRow As Long, lastRow As Long, dataStartRow As Long, headerRow As Long
    firstRow = workRange.Row
    lastRow = workRange.Row + workRange.Rows.Count - 1
    dataStartRow = firstRow
    If Address_LooksLikeHeader(CStr(workRange.Cells(1, 1).Value)) Then
        headerRow = firstRow
        dataStartRow = firstRow + 1
    ElseIf firstRow > 1 Then
        headerRow = firstRow - 1
    End If

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStartRow + 1)
    If totalRows = 0 Then
        MsgBox "헤더를 제외하면 변환할 주소가 없습니다.", vbInformation, "광역시도 표기 변환"
        Exit Sub
    End If

    Dim outputCol As Long
    outputCol = workRange.Column + 1
    If Not Toolkit_CanInsertColumns(ws, outputCol, 1, "광역시도 표기 변환") Then Exit Sub

    Dim sourceValues As Variant, resultValues() As Variant
    sourceValues = ws.Range(ws.Cells(dataStartRow, workRange.Column), ws.Cells(lastRow, workRange.Column)).Value2
    ReDim resultValues(1 To totalRows, 1 To 1)

    Dim i As Long, sourceText As String, convertedText As String, changedCount As Long
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "광역시도 표기 변환", totalRows, "광역시도 표기 변환 준비 중"
    For i = 1 To totalRows
        If totalRows = 1 Then
            sourceText = CStr(sourceValues)
        Else
            sourceText = CStr(sourceValues(i, 1))
        End If
        convertedText = Address_ConvertProvinceFirstToken(sourceText, direction, formalStyle)
        resultValues(i, 1) = convertedText
        If convertedText <> sourceText Then changedCount = changedCount + 1
        Toolkit_RowProgressUpdate i, totalRows, "광역시도 표기 변환 중: " & CStr(i) & "/" & CStr(totalRows) & "행"
    Next i

    ws.Columns(outputCol).Insert Shift:=xlToRight
    If headerRow > 0 Then
        If direction = "formal" Then
            ws.Cells(headerRow, outputCol).Value = "주소_광역시도정규형"
        Else
            ws.Cells(headerRow, outputCol).Value = "주소_광역시도축약형"
        End If
        ws.Cells(headerRow, outputCol).Font.Bold = True
    End If
    ws.Range(ws.Cells(dataStartRow, outputCol), ws.Cells(lastRow, outputCol)).Value = resultValues
    ws.Columns(outputCol).AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True

    MsgBox "광역시도 표기 변환을 완료했습니다." & vbCrLf & _
           "변환: " & changedCount & "건" & vbCrLf & _
           "변경 없음: " & (totalRows - changedCount) & "건", vbInformation, "광역시도 표기 변환"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("광역시도 표기 변환") Then Exit Sub
    MsgBox "광역시도 표기 변환 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "광역시도 표기 변환"
End Sub

Public Function Address_ProvinceNotationForTest(ByVal addressText As String, ByVal direction As String, ByVal formalStyle As String) As String
    Address_ProvinceNotationForTest = Address_ConvertProvinceFirstToken(addressText, direction, formalStyle)
End Function

Private Function Address_ConvertProvinceFirstToken(ByVal addressText As String, ByVal direction As String, ByVal formalStyle As String) As String
    Dim textValue As String
    textValue = Trim$(CStr(addressText))
    If Len(textValue) = 0 Then Exit Function

    Dim separatorPosition As Long, firstToken As String, remainder As String
    separatorPosition = Address_FirstWhitespacePosition(textValue)
    If separatorPosition > 0 Then
        firstToken = Left$(textValue, separatorPosition - 1)
        remainder = Mid$(textValue, separatorPosition)
    Else
        firstToken = textValue
    End If

    Dim shortName As String, replacement As String
    shortName = Address_ProvinceShortName(firstToken)
    If Len(shortName) = 0 Then
        Address_ConvertProvinceFirstToken = addressText
        Exit Function
    End If

    Select Case LCase$(direction)
        Case "short"
            replacement = shortName
        Case "normalize"
            replacement = Address_ProvinceFormalFromShort(shortName, "latest")
        Case Else
            If Address_IsProvinceFormalToken(firstToken) Then
                replacement = firstToken
            Else
                replacement = Address_ProvinceFormalFromShort(shortName, formalStyle)
            End If
    End Select

    Address_ConvertProvinceFirstToken = replacement & remainder
End Function

Private Function Address_FirstWhitespacePosition(ByVal textValue As String) As Long
    Dim i As Long, charCode As Long
    For i = 1 To Len(textValue)
        charCode = AscW(Mid$(textValue, i, 1))
        If charCode = 32 Or charCode = 9 Or charCode = 10 Or charCode = 13 Then
            Address_FirstWhitespacePosition = i
            Exit Function
        End If
    Next i
End Function

Private Function Address_ProvinceShortName(ByVal provinceName As String) As String
    Select Case Trim$(provinceName)
        Case "서울", "서울시", "서울특별시": Address_ProvinceShortName = "서울"
        Case "부산", "부산시", "부산광역시": Address_ProvinceShortName = "부산"
        Case "대구", "대구시", "대구광역시": Address_ProvinceShortName = "대구"
        Case "인천", "인천시", "인천광역시": Address_ProvinceShortName = "인천"
        Case "광주", "광주시", "광주광역시": Address_ProvinceShortName = "광주"
        Case "대전", "대전시", "대전광역시": Address_ProvinceShortName = "대전"
        Case "울산", "울산시", "울산광역시": Address_ProvinceShortName = "울산"
        Case "세종", "세종시", "세종특별자치시": Address_ProvinceShortName = "세종"
        Case "경기", "경기도": Address_ProvinceShortName = "경기"
        Case "강원", "강원도", "강원특별자치도": Address_ProvinceShortName = "강원"
        Case "충북", "충청북도": Address_ProvinceShortName = "충북"
        Case "충남", "충청남도": Address_ProvinceShortName = "충남"
        Case "전북", "전라북도", "전북특별자치도": Address_ProvinceShortName = "전북"
        Case "전남", "전라남도": Address_ProvinceShortName = "전남"
        Case "경북", "경상북도": Address_ProvinceShortName = "경북"
        Case "경남", "경상남도": Address_ProvinceShortName = "경남"
        Case "제주", "제주도", "제주특별자치도": Address_ProvinceShortName = "제주"
    End Select
End Function

Private Function Address_IsProvinceFormalToken(ByVal provinceName As String) As Boolean
    Select Case Trim$(provinceName)
        Case "서울특별시", "부산광역시", "대구광역시", "인천광역시", "광주광역시", _
             "대전광역시", "울산광역시", "세종특별자치시", "경기도", "강원도", "강원특별자치도", _
             "충청북도", "충청남도", "전라북도", "전북특별자치도", "전라남도", "경상북도", _
             "경상남도", "제주도", "제주특별자치도"
            Address_IsProvinceFormalToken = True
    End Select
End Function

Private Function Address_ProvinceFormalFromShort(ByVal shortName As String, ByVal formalStyle As String) As String
    Select Case shortName
        Case "서울": Address_ProvinceFormalFromShort = "서울특별시"
        Case "부산": Address_ProvinceFormalFromShort = "부산광역시"
        Case "대구": Address_ProvinceFormalFromShort = "대구광역시"
        Case "인천": Address_ProvinceFormalFromShort = "인천광역시"
        Case "광주": Address_ProvinceFormalFromShort = "광주광역시"
        Case "대전": Address_ProvinceFormalFromShort = "대전광역시"
        Case "울산": Address_ProvinceFormalFromShort = "울산광역시"
        Case "세종": Address_ProvinceFormalFromShort = "세종특별자치시"
        Case "경기": Address_ProvinceFormalFromShort = "경기도"
        Case "강원"
            If LCase$(formalStyle) = "legacy" Then Address_ProvinceFormalFromShort = "강원도" Else Address_ProvinceFormalFromShort = "강원특별자치도"
        Case "충북": Address_ProvinceFormalFromShort = "충청북도"
        Case "충남": Address_ProvinceFormalFromShort = "충청남도"
        Case "전북"
            If LCase$(formalStyle) = "legacy" Then Address_ProvinceFormalFromShort = "전라북도" Else Address_ProvinceFormalFromShort = "전북특별자치도"
        Case "전남": Address_ProvinceFormalFromShort = "전라남도"
        Case "경북": Address_ProvinceFormalFromShort = "경상북도"
        Case "경남": Address_ProvinceFormalFromShort = "경상남도"
        Case "제주"
            If LCase$(formalStyle) = "legacy" Then Address_ProvinceFormalFromShort = "제주도" Else Address_ProvinceFormalFromShort = "제주특별자치도"
    End Select
End Function

Private Sub Address_ClassifyRangeType(ByVal addressRange As Range)
    If addressRange.Areas.Count > 1 Then
        MsgBox "주소 유형 판별은 한 번에 하나의 연속된 범위만 처리합니다.", vbExclamation, "주소 유형 판별"
        Exit Sub
    End If

    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(addressRange)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 판별할 주소를 찾지 못했습니다.", vbInformation, "주소 유형 판별"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = workRange.Worksheet

    Dim firstRow As Long, lastRow As Long, dataStartRow As Long, headerRow As Long
    firstRow = workRange.Row
    lastRow = workRange.Row + workRange.Rows.Count - 1
    If Address_LooksLikeHeader(CStr(workRange.Cells(1, 1).Value)) Then
        headerRow = firstRow
        dataStartRow = firstRow + 1
    Else
        dataStartRow = firstRow
        If firstRow > 1 Then headerRow = firstRow - 1
    End If

    Dim outputCol As Long
    outputCol = workRange.Column + workRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outputCol, 2, "주소 유형 판별") Then Exit Sub

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStartRow + 1)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outputCol).Resize(, 2).Insert Shift:=xlToRight
    If headerRow > 0 Then
        ws.Cells(headerRow, outputCol).Value = "주소 유형"
        ws.Cells(headerRow, outputCol + 1).Value = "판별 기준"
        ws.Range(ws.Cells(headerRow, outputCol), ws.Cells(headerRow, outputCol + 1)).Font.Bold = True
    End If

    Dim r As Long
    Dim roadCount As Long
    Dim dongCount As Long
    Dim unknownCount As Long
    Dim sourceAddress As String
    Dim classified As Variant

    Toolkit_RowProgressStart "주소 유형 판별", totalRows, "주소 유형 판별 준비 중"
    For r = dataStartRow To lastRow
        sourceAddress = Trim$(CStr(ws.Cells(r, workRange.Column).Value))
        If Len(sourceAddress) > 0 Then
            classified = Address_ClassifyOneAddress(sourceAddress)
            ws.Cells(r, outputCol).Value = CStr(classified(0))
            ws.Cells(r, outputCol + 1).Value = CStr(classified(1))

            Select Case CStr(classified(0))
                Case "도로명주소": roadCount = roadCount + 1
                Case "법정/행정동 기준 주소": dongCount = dongCount + 1
                Case Else: unknownCount = unknownCount + 1
            End Select
        End If
        Toolkit_RowProgressUpdate r - dataStartRow + 1, totalRows, "주소 유형 판별 중: " & CStr(r - dataStartRow + 1) & "/" & CStr(totalRows) & "행"
    Next r

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(lastRow, outputCol + 1)).EntireColumn.AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True

    MsgBox "주소 유형 판별을 완료했습니다." & vbCrLf & _
           "도로명주소: " & roadCount & "건" & vbCrLf & _
           "법정/행정동 기준 주소: " & dongCount & "건" & vbCrLf & _
           "확인 필요: " & unknownCount & "건", vbInformation, "주소 유형 판별"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("주소 유형 판별") Then Exit Sub
    MsgBox "주소 유형 판별 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "주소 유형 판별"
End Sub

Public Sub Address_SplitDongAndRoad()
    Dim addressRange As Range
    If Not Toolkit_GetRangeFromUserForm("주소 체계 분리", _
        "동 기준 주소와 도로명주소로 나눌 주소 범위를 선택하세요. 결과는 원본 오른쪽에 2개 열로 생성됩니다.", _
        addressRange) Then Exit Sub

    Address_SplitRangeDongAndRoad addressRange
End Sub

Private Sub Address_SplitRangeDongAndRoad(ByVal addressRange As Range)
    If addressRange.Areas.Count > 1 Then
        MsgBox "주소 체계 분리는 한 번에 하나의 연속된 범위만 처리합니다.", vbExclamation, "주소 체계 분리"
        Exit Sub
    End If

    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(addressRange)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 처리할 주소를 찾지 못했습니다.", vbInformation, "주소 체계 분리"
        Exit Sub
    End If

    Dim records As Collection
    Set records = Address_LoadFlexibleMappingRecords(ActiveWorkbook)

    Dim ws As Worksheet
    Set ws = workRange.Worksheet

    Dim firstRow As Long, lastRow As Long, dataStartRow As Long, headerRow As Long
    firstRow = workRange.Row
    lastRow = workRange.Row + workRange.Rows.Count - 1
    If Address_LooksLikeHeader(CStr(workRange.Cells(1, 1).Value)) Then
        headerRow = firstRow
        dataStartRow = firstRow + 1
    Else
        dataStartRow = firstRow
        If firstRow > 1 Then headerRow = firstRow - 1
    End If

    Dim outputCol As Long
    outputCol = workRange.Column + workRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outputCol, 2, "주소 체계 분리") Then Exit Sub

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStartRow + 1)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outputCol).Resize(, 2).Insert Shift:=xlToRight
    If headerRow > 0 Then
        ws.Cells(headerRow, outputCol).Value = "동 기준 주소"
        ws.Cells(headerRow, outputCol + 1).Value = "도로명주소"
        ws.Range(ws.Cells(headerRow, outputCol), ws.Cells(headerRow, outputCol + 1)).Font.Bold = True
    End If
    ws.Columns(outputCol).Resize(, 2).NumberFormat = "@"

    Dim r As Long
    Dim matchedCount As Long
    Dim heuristicCount As Long
    Dim sourceAddress As String
    Dim splitResult As Variant

    Toolkit_RowProgressStart "주소 체계 분리", totalRows, "주소 체계 분리 준비 중"
    For r = dataStartRow To lastRow
        sourceAddress = Trim$(CStr(ws.Cells(r, workRange.Column).Value))
        If Len(sourceAddress) > 0 Then
            splitResult = Address_SplitOneAddress(sourceAddress, records)
            ws.Cells(r, outputCol).Value = CStr(splitResult(0))
            ws.Cells(r, outputCol + 1).Value = CStr(splitResult(1))
            If CBool(splitResult(2)) Then matchedCount = matchedCount + 1 Else heuristicCount = heuristicCount + 1
        End If
        Toolkit_RowProgressUpdate r - dataStartRow + 1, totalRows, "주소 체계 분리 중: " & CStr(r - dataStartRow + 1) & "/" & CStr(totalRows) & "행"
    Next r

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(lastRow, outputCol + 1)).EntireColumn.AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True

    MsgBox "주소 체계 분리를 완료했습니다." & vbCrLf & _
           "변환표 매칭: " & matchedCount & "건" & vbCrLf & _
           "주소 형태 기반 분리: " & heuristicCount & "건" & vbCrLf & _
           "변환표가 있으면 지번/도로명 양방향 매칭 정확도가 높아집니다.", vbInformation, "주소 체계 분리"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("주소 체계 분리") Then Exit Sub
    MsgBox "주소 체계 분리 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "주소 체계 분리"
End Sub

Public Sub Address_CombineDongAndRoad()
    Dim dongRange As Range
    If Not Toolkit_GetRangeFromUserForm("주소 통합 - 동 기준 주소", _
        "통합할 동 기준 주소 범위를 선택하세요. 주소 체계 분리 결과의 '동 기준 주소' 열을 선택하면 됩니다.", _
        dongRange) Then Exit Sub

    Dim roadRange As Range
    If Not Toolkit_GetRangeFromUserForm("주소 통합 - 도로명주소", _
        "통합할 도로명주소 범위를 선택하세요. 주소 체계 분리 결과의 '도로명주소' 열을 선택하면 됩니다.", _
        roadRange) Then Exit Sub

    Dim preference As String
    preference = Toolkit_ButtonMenu( _
        "주소 통합 - 우선순위", _
        "두 주소가 모두 있을 때 어떤 주소를 통합주소로 사용할지 선택하세요.", _
        "road" & vbTab & "도로명주소 우선" & vbLf & _
        "dong" & vbTab & "동 기준 주소 우선", _
        "road")
    If Len(Trim$(preference)) = 0 Then Exit Sub

    Address_CombineRangeDongAndRoad dongRange, roadRange, preference
End Sub

Private Sub Address_CombineRangeDongAndRoad(ByVal dongRange As Range, ByVal roadRange As Range, ByVal preference As String)
    If dongRange.Areas.Count > 1 Or roadRange.Areas.Count > 1 Then
        MsgBox "주소 통합은 각 주소 범위를 한 덩어리의 연속 범위로 선택해야 합니다.", vbExclamation, "주소 통합"
        Exit Sub
    End If

    If dongRange.Columns.Count > 1 Then Set dongRange = dongRange.Columns(1)
    If roadRange.Columns.Count > 1 Then Set roadRange = roadRange.Columns(1)

    If Not dongRange.Worksheet Is roadRange.Worksheet Then
        MsgBox "동 기준 주소와 도로명주소 범위는 같은 시트에 있어야 합니다.", vbExclamation, "주소 통합"
        Exit Sub
    End If

    If dongRange.Row <> roadRange.Row Or dongRange.Rows.Count <> roadRange.Rows.Count Then
        MsgBox "동 기준 주소와 도로명주소 범위의 시작 행 또는 행 수가 다릅니다.", vbExclamation, "주소 통합"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = dongRange.Worksheet

    Dim firstRow As Long, lastRow As Long, dataStartRow As Long, headerRow As Long
    firstRow = dongRange.Row
    lastRow = dongRange.Row + dongRange.Rows.Count - 1
    If Address_LooksLikeHeader(CStr(dongRange.Cells(1, 1).Value)) Or Address_LooksLikeHeader(CStr(roadRange.Cells(1, 1).Value)) Then
        headerRow = firstRow
        dataStartRow = firstRow + 1
    Else
        dataStartRow = firstRow
        If firstRow > 1 Then headerRow = firstRow - 1
    End If

    If dataStartRow > lastRow Then
        MsgBox "헤더를 제외하면 통합할 주소 데이터가 없습니다.", vbInformation, "주소 통합"
        Exit Sub
    End If

    Dim outputCol As Long
    outputCol = Application.Max(dongRange.Column + dongRange.Columns.Count, roadRange.Column + roadRange.Columns.Count)
    If Not Toolkit_CanInsertColumns(ws, outputCol, 1, "주소 통합") Then Exit Sub

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStartRow + 1)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outputCol).Insert Shift:=xlToRight
    If headerRow > 0 Then
        ws.Cells(headerRow, outputCol).Value = "통합주소"
        ws.Cells(headerRow, outputCol).Font.Bold = True
    End If
    ws.Columns(outputCol).NumberFormat = "@"

    Dim r As Long
    Dim roadSelectedCount As Long
    Dim dongSelectedCount As Long
    Dim blankCount As Long
    Toolkit_RowProgressStart "주소 통합", totalRows, "주소 통합 준비 중"
    For r = dataStartRow To lastRow
        Dim dongAddress As String
        Dim roadAddress As String
        Dim combined As String
        dongAddress = Trim$(CStr(ws.Cells(r, dongRange.Column).Value))
        roadAddress = Trim$(CStr(ws.Cells(r, roadRange.Column).Value))
        combined = Address_ChooseCombinedAddress(dongAddress, roadAddress, preference)
        ws.Cells(r, outputCol).Value = combined

        If Len(combined) = 0 Then
            blankCount = blankCount + 1
        ElseIf StrComp(combined, roadAddress, vbTextCompare) = 0 Then
            roadSelectedCount = roadSelectedCount + 1
        Else
            dongSelectedCount = dongSelectedCount + 1
        End If
        Toolkit_RowProgressUpdate r - dataStartRow + 1, totalRows, "주소 통합 중: " & CStr(r - dataStartRow + 1) & "/" & CStr(totalRows) & "행"
    Next r

    ws.Columns(outputCol).AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True

    MsgBox "분할 주소 통합을 완료했습니다." & vbCrLf & _
           "도로명주소 사용: " & roadSelectedCount & "건" & vbCrLf & _
           "동 기준 주소 사용: " & dongSelectedCount & "건" & vbCrLf & _
           "빈값: " & blankCount & "건", vbInformation, "주소 통합"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("주소 통합") Then Exit Sub
    MsgBox "주소 통합 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "주소 통합"
End Sub

Private Function Address_ChooseCombinedAddress(ByVal dongAddress As String, ByVal roadAddress As String, ByVal preference As String) As String
    If LCase$(Trim$(preference)) = "dong" Then
        If Len(Trim$(dongAddress)) > 0 Then
            Address_ChooseCombinedAddress = Trim$(dongAddress)
        Else
            Address_ChooseCombinedAddress = Trim$(roadAddress)
        End If
    Else
        If Len(Trim$(roadAddress)) > 0 Then
            Address_ChooseCombinedAddress = Trim$(roadAddress)
        Else
            Address_ChooseCombinedAddress = Trim$(dongAddress)
        End If
    End If
End Function

Private Sub Address_ConvertRangeWithMapping(ByVal addressRange As Range, ByVal mapRange As Range)
    If addressRange.Areas.Count > 1 Then
        MsgBox "주소 변환은 한 번에 하나의 연속된 범위만 처리합니다.", vbExclamation, "도로명주소 변환"
        Exit Sub
    End If

    Dim oldCol As Long, roadCol As Long, zipCol As Long
    If Not Address_DetectMappingColumns(mapRange, oldCol, roadCol, zipCol) Then
        MsgBox "변환표에서 '구주소', '도로명주소', '우편번호' 헤더를 찾지 못했습니다.", vbExclamation, "도로명주소 변환"
        Exit Sub
    End If

    Dim records As Collection
    Set records = Address_LoadMappingRecords(mapRange, oldCol, roadCol, zipCol)
    If records.Count = 0 Then
        MsgBox "변환표에 사용할 주소 데이터가 없습니다.", vbExclamation, "도로명주소 변환"
        Exit Sub
    End If

    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(addressRange)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 변환할 주소를 찾지 못했습니다.", vbInformation, "도로명주소 변환"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = workRange.Worksheet

    Dim firstRow As Long, lastRow As Long, dataStartRow As Long, headerRow As Long
    firstRow = workRange.Row
    lastRow = workRange.Row + workRange.Rows.Count - 1
    If Address_LooksLikeHeader(CStr(workRange.Cells(1, 1).Value)) Then
        headerRow = firstRow
        dataStartRow = firstRow + 1
    Else
        dataStartRow = firstRow
        If firstRow > 1 Then headerRow = firstRow - 1
    End If

    Dim outputCol As Long
    outputCol = workRange.Column + workRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outputCol, 3, "주소 변환") Then Exit Sub

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStartRow + 1)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outputCol).Resize(, 3).Insert Shift:=xlToRight
    If headerRow > 0 Then
        ws.Cells(headerRow, outputCol).Value = "우편번호"
        ws.Cells(headerRow, outputCol + 1).Value = "도로명주소"
        ws.Cells(headerRow, outputCol + 2).Value = "변환결과"
        ws.Range(ws.Cells(headerRow, outputCol), ws.Cells(headerRow, outputCol + 2)).Font.Bold = True
    End If

    ws.Columns(outputCol).Resize(, 3).NumberFormat = "@"

    Dim r As Long, successCount As Long, failCount As Long
    Toolkit_RowProgressStart "도로명주소 변환", totalRows, "주소 변환 준비 중"
    For r = dataStartRow To lastRow
        Dim sourceAddress As String
        sourceAddress = Trim$(CStr(ws.Cells(r, workRange.Column).Value))
        If Len(sourceAddress) > 0 Then
            Dim matched As Variant
            matched = Address_FindBestMatch(sourceAddress, records)
            If CBool(matched(0)) Then
                ws.Cells(r, outputCol).Value = CStr(matched(2))
                ws.Cells(r, outputCol + 1).Value = CStr(matched(1))
                ws.Cells(r, outputCol + 2).Value = "성공"
                successCount = successCount + 1
            Else
                ws.Cells(r, outputCol + 2).Value = "실패"
                failCount = failCount + 1
            End If
        End If
        Toolkit_RowProgressUpdate r - dataStartRow + 1, totalRows, "주소 변환 중: " & CStr(r - dataStartRow + 1) & "/" & CStr(totalRows) & "행"
    Next r

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(lastRow, outputCol + 2)).EntireColumn.AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True

    MsgBox "도로명주소 변환을 완료했습니다." & vbCrLf & _
           "성공: " & successCount & "건" & vbCrLf & _
           "실패: " & failCount & "건", vbInformation, "도로명주소 변환"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("도로명주소 변환") Then Exit Sub
    MsgBox "도로명주소 변환 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "도로명주소 변환"
End Sub

Private Function Address_FindMappingRange(ByVal wb As Workbook) As Range
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim usedRange As Range
        Set usedRange = ws.UsedRange
        Dim r As Long
        For r = usedRange.Row To WorksheetFunction.Min(usedRange.Row + usedRange.Rows.Count - 1, usedRange.Row + 20)
            Dim testRange As Range
            Set testRange = ws.Range(ws.Cells(r, usedRange.Column), ws.Cells(usedRange.Row + usedRange.Rows.Count - 1, usedRange.Column + usedRange.Columns.Count - 1))
            Dim oldCol As Long, roadCol As Long, zipCol As Long
            If Address_DetectMappingColumns(testRange, oldCol, roadCol, zipCol) Then
                Set Address_FindMappingRange = testRange
                Exit Function
            End If
        Next r
    Next ws
End Function

Private Function Address_DetectMappingColumns(ByVal mapRange As Range, ByRef oldCol As Long, ByRef roadCol As Long, ByRef zipCol As Long) As Boolean
    oldCol = 0: roadCol = 0: zipCol = 0

    Dim c As Long
    For c = 1 To mapRange.Columns.Count
        Dim headerText As String
        headerText = Replace(Trim$(CStr(mapRange.Cells(1, c).Value)), " ", "")
        If headerText = "구주소" Or headerText = "지번주소" Then oldCol = c
        If headerText = "도로명주소" Or headerText = "도로명" Then roadCol = c
        If headerText = "우편번호" Then zipCol = c
    Next c

    Address_DetectMappingColumns = (oldCol > 0 And roadCol > 0 And zipCol > 0)
End Function

Private Function Address_LoadMappingRecords(ByVal mapRange As Range, ByVal oldCol As Long, ByVal roadCol As Long, ByVal zipCol As Long) As Collection
    Dim records As Collection
    Set records = New Collection

    Dim r As Long
    For r = 2 To mapRange.Rows.Count
        Dim oldAddress As String, roadAddress As String, zipCode As String
        oldAddress = Trim$(CStr(mapRange.Cells(r, oldCol).Value))
        roadAddress = Trim$(CStr(mapRange.Cells(r, roadCol).Value))
        zipCode = Trim$(CStr(mapRange.Cells(r, zipCol).Value))
        If Len(oldAddress) > 0 And Len(roadAddress) > 0 Then
            records.Add Array(Address_NormalizeKey(oldAddress), oldAddress, roadAddress, zipCode)
        End If
    Next r

    Set Address_LoadMappingRecords = records
End Function

Private Function Address_LoadFlexibleMappingRecords(ByVal wb As Workbook) As Collection
    Dim records As New Collection
    If wb Is Nothing Then
        Set Address_LoadFlexibleMappingRecords = records
        Exit Function
    End If

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim usedRange As Range
        Set usedRange = ws.UsedRange
        If Not usedRange Is Nothing Then
            Dim r As Long
            For r = usedRange.Row To WorksheetFunction.Min(usedRange.Row + usedRange.Rows.Count - 1, usedRange.Row + 20)
                Dim testRange As Range
                Set testRange = ws.Range(ws.Cells(r, usedRange.Column), ws.Cells(usedRange.Row + usedRange.Rows.Count - 1, usedRange.Column + usedRange.Columns.Count - 1))

                Dim oldCol As Long, roadCol As Long, zipCol As Long
                If Address_DetectMappingColumnsFlexible(testRange, oldCol, roadCol, zipCol) Then
                    Dim loaded As Collection
                    Set loaded = Address_LoadMappingRecordsFlexible(testRange, oldCol, roadCol)
                    Dim item As Variant
                    For Each item In loaded
                        records.Add item
                    Next item
                End If
            Next r
        End If
    Next ws

    Set Address_LoadFlexibleMappingRecords = records
End Function

Private Function Address_DetectMappingColumnsFlexible(ByVal mapRange As Range, ByRef oldCol As Long, ByRef roadCol As Long, ByRef zipCol As Long) As Boolean
    oldCol = 0: roadCol = 0: zipCol = 0

    Dim c As Long
    For c = 1 To mapRange.Columns.Count
        Dim headerText As String
        headerText = Replace(Trim$(CStr(mapRange.Cells(1, c).Value)), " ", "")
        If headerText = "구주소" Or headerText = "지번주소" Or headerText = "동기준주소" Or headerText = "법정동주소" Or headerText = "행정동주소" Then oldCol = c
        If headerText = "도로명주소" Or headerText = "도로명" Then roadCol = c
        If headerText = "우편번호" Then zipCol = c
    Next c

    Address_DetectMappingColumnsFlexible = (oldCol > 0 And roadCol > 0)
End Function

Private Function Address_LoadMappingRecordsFlexible(ByVal mapRange As Range, ByVal oldCol As Long, ByVal roadCol As Long) As Collection
    Dim records As New Collection
    Dim r As Long
    For r = 2 To mapRange.Rows.Count
        Dim oldAddress As String, roadAddress As String
        oldAddress = Trim$(CStr(mapRange.Cells(r, oldCol).Value))
        roadAddress = Trim$(CStr(mapRange.Cells(r, roadCol).Value))
        If Len(oldAddress) > 0 Or Len(roadAddress) > 0 Then
            records.Add Array(Address_NormalizeKey(oldAddress), Address_NormalizeKey(roadAddress), oldAddress, roadAddress)
        End If
    Next r
    Set Address_LoadMappingRecordsFlexible = records
End Function

Private Function Address_SplitOneAddress(ByVal sourceAddress As String, ByVal records As Collection) As Variant
    Dim matched As Variant
    matched = Address_FindBestDongRoadMatch(sourceAddress, records)
    If CBool(matched(0)) Then
        Address_SplitOneAddress = Array(CStr(matched(1)), CStr(matched(2)), True)
        Exit Function
    End If

    If Address_IsRoadAddress(sourceAddress) Then
        Address_SplitOneAddress = Array(Address_ExtractDongBasisAddress(sourceAddress), sourceAddress, False)
    Else
        Address_SplitOneAddress = Array(sourceAddress, "", False)
    End If
End Function

Private Function Address_FindBestDongRoadMatch(ByVal sourceAddress As String, ByVal records As Collection) As Variant
    Dim sourceKey As String
    sourceKey = Address_NormalizeKey(sourceAddress)

    Dim bestScore As Long
    Dim bestDong As String
    Dim bestRoad As String

    If records Is Nothing Then
        Address_FindBestDongRoadMatch = Array(False, "", "")
        Exit Function
    End If

    Dim i As Long
    For i = 1 To records.Count
        Dim item As Variant
        item = records(i)

        Dim score As Long
        score = WorksheetFunction.Max( _
            Address_MatchScore(sourceKey, CStr(item(0))), _
            Address_MatchScore(sourceKey, CStr(item(1))))

        If score > bestScore Then
            bestScore = score
            bestDong = CStr(item(2))
            bestRoad = CStr(item(3))
        End If
    Next i

    Address_FindBestDongRoadMatch = Array(bestScore > 0, bestDong, bestRoad)
End Function

Private Function Address_MatchScore(ByVal sourceKey As String, ByVal candidateKey As String) As Long
    If Len(sourceKey) < 8 Or Len(candidateKey) < 8 Then Exit Function
    If sourceKey = candidateKey Then
        Address_MatchScore = 100000 + Len(candidateKey)
    ElseIf InStr(1, sourceKey, candidateKey, vbTextCompare) > 0 Then
        Address_MatchScore = 50000 + Len(candidateKey)
    ElseIf InStr(1, candidateKey, sourceKey, vbTextCompare) > 0 Then
        Address_MatchScore = 30000 + Len(sourceKey)
    End If
End Function

Private Function Address_FindBestMatch(ByVal sourceAddress As String, ByVal records As Collection) As Variant
    Dim sourceKey As String
    sourceKey = Address_NormalizeKey(sourceAddress)

    Dim bestScore As Long
    Dim bestRoad As String
    Dim bestZip As String

    Dim i As Long
    For i = 1 To records.Count
        Dim item As Variant
        item = records(i)
        Dim oldKey As String
        oldKey = CStr(item(0))
        If Len(oldKey) >= 8 And Len(sourceKey) >= 8 Then
            Dim score As Long
            If sourceKey = oldKey Then
                score = 100000 + Len(oldKey)
            ElseIf InStr(1, sourceKey, oldKey, vbTextCompare) > 0 Then
                score = 50000 + Len(oldKey)
            ElseIf InStr(1, oldKey, sourceKey, vbTextCompare) > 0 Then
                score = 30000 + Len(sourceKey)
            End If

            If score > bestScore Then
                bestScore = score
                bestRoad = CStr(item(2))
                bestZip = CStr(item(3))
            End If
        End If
    Next i

    Address_FindBestMatch = Array(bestScore > 0, bestRoad, bestZip)
End Function

Private Function Address_NormalizeKey(ByVal addressText As String) As String
    Dim textValue As String
    textValue = Trim$(CStr(addressText))
    textValue = Address_NormalizeRegionNames(textValue)
    textValue = Replace(textValue, ChrW(160), "")
    textValue = Replace(textValue, vbTab, "")
    textValue = Replace(textValue, " ", "")
    textValue = Replace(textValue, "번지", "")
    textValue = Replace(textValue, "<", "")
    textValue = Replace(textValue, ">", "")
    textValue = Replace(textValue, "(", "")
    textValue = Replace(textValue, ")", "")
    textValue = Replace(textValue, ",", "")
    Address_NormalizeKey = textValue
End Function

Private Function Address_NormalizeRegionNames(ByVal addressText As String) As String
    Address_NormalizeRegionNames = Address_ConvertProvinceFirstToken(addressText, "normalize", "latest")
End Function

Private Function Address_IsRoadAddress(ByVal addressText As String) As Boolean
    Dim textValue As String
    textValue = " " & Replace(Trim$(CStr(addressText)), "  ", " ") & " "
    Address_IsRoadAddress = (InStr(1, textValue, "로 ", vbTextCompare) > 0 _
                          Or InStr(1, textValue, "길 ", vbTextCompare) > 0 _
                          Or InStr(1, textValue, "대로 ", vbTextCompare) > 0 _
                          Or InStr(1, textValue, "번길 ", vbTextCompare) > 0)
End Function

Public Function Address_ClassifyOneAddress(ByVal addressText As String) As Variant
    Dim textValue As String
    textValue = Trim$(CStr(addressText))

    If Len(textValue) = 0 Then
        Address_ClassifyOneAddress = Array("확인 필요", "빈 주소")
    ElseIf Address_IsRoadAddress(textValue) Then
        Address_ClassifyOneAddress = Array("도로명주소", "도로명 키워드(로/길/대로/번길) 감지")
    ElseIf Address_IsDongBasisAddress(textValue) Then
        Address_ClassifyOneAddress = Array("법정/행정동 기준 주소", "읍/면/동/리/가 또는 지번 패턴 감지")
    Else
        Address_ClassifyOneAddress = Array("확인 필요", "도로명/동 기준 주소 패턴이 명확하지 않음")
    End If
End Function

Private Function Address_IsDongBasisAddress(ByVal addressText As String) As Boolean
    Dim textValue As String
    textValue = " " & Trim$(CStr(addressText)) & " "

    If InStr(1, textValue, "번지", vbTextCompare) > 0 Then
        Address_IsDongBasisAddress = True
        Exit Function
    End If
    If InStr(1, textValue, " 산 ", vbTextCompare) > 0 Then
        Address_IsDongBasisAddress = True
        Exit Function
    End If

    Dim parts() As String
    parts = Split(Trim$(textValue), " ")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Address_LooksLikeDongToken(parts(i)) Then
            Address_IsDongBasisAddress = True
            Exit Function
        End If
    Next i
End Function

Private Function Address_ExtractDongBasisAddress(ByVal roadAddress As String) As String
    Dim textValue As String
    textValue = Trim$(Address_NormalizeRegionNames(CStr(roadAddress)))

    Dim parenValue As String
    parenValue = Address_TextInsideLastParentheses(textValue)
    If Address_LooksLikeDongToken(parenValue) Then
        Dim regionPart As String
        regionPart = Address_RegionPrefixBeforeRoad(textValue)
        Address_ExtractDongBasisAddress = Trim$(regionPart & " " & parenValue)
        Exit Function
    End If

    Address_ExtractDongBasisAddress = Address_PrefixThroughDongToken(textValue)
End Function

Private Function Address_TextInsideLastParentheses(ByVal textValue As String) As String
    Dim openPos As Long
    Dim closePos As Long
    closePos = InStrRev(textValue, ")")
    openPos = InStrRev(textValue, "(")
    If openPos > 0 And closePos > openPos Then
        Address_TextInsideLastParentheses = Mid$(textValue, openPos + 1, closePos - openPos - 1)
    End If
End Function

Private Function Address_RegionPrefixBeforeRoad(ByVal textValue As String) As String
    Dim parts() As String
    parts = Split(Trim$(textValue), " ")

    Dim result As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Address_LooksLikeRoadToken(parts(i)) Then Exit For
        If Len(parts(i)) > 0 Then result = Trim$(result & " " & parts(i))
        If Address_LooksLikeDistrictToken(parts(i)) Then Exit For
    Next i

    Address_RegionPrefixBeforeRoad = result
End Function

Private Function Address_PrefixThroughDongToken(ByVal textValue As String) As String
    Dim cleaned As String
    cleaned = textValue
    If InStr(1, cleaned, "(", vbTextCompare) > 0 Then cleaned = Trim$(Left$(cleaned, InStr(1, cleaned, "(", vbTextCompare) - 1))

    Dim parts() As String
    parts = Split(Trim$(cleaned), " ")

    Dim result As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Address_LooksLikeRoadToken(parts(i)) Then Exit For
        If Len(parts(i)) > 0 Then result = Trim$(result & " " & parts(i))
        If Address_LooksLikeDongToken(parts(i)) Then Exit For
    Next i

    Address_PrefixThroughDongToken = result
End Function

Private Function Address_LooksLikeRoadToken(ByVal token As String) As Boolean
    token = Trim$(token)
    Address_LooksLikeRoadToken = (Right$(token, 1) = "로" _
                               Or Right$(token, 1) = "길" _
                               Or Right$(token, 2) = "대로" _
                               Or Right$(token, 2) = "번길")
End Function

Private Function Address_LooksLikeDistrictToken(ByVal token As String) As Boolean
    token = Trim$(token)
    Address_LooksLikeDistrictToken = (Right$(token, 1) = "구" _
                                   Or Right$(token, 1) = "군" _
                                   Or Right$(token, 1) = "시")
End Function

Private Function Address_LooksLikeDongToken(ByVal token As String) As Boolean
    token = Trim$(token)
    If Len(token) = 0 Then Exit Function
    Address_LooksLikeDongToken = (InStr(1, token, "동", vbTextCompare) > 0 _
                               Or InStr(1, token, "읍", vbTextCompare) > 0 _
                               Or InStr(1, token, "면", vbTextCompare) > 0 _
                               Or InStr(1, token, "리", vbTextCompare) > 0 _
                               Or Right$(token, 1) = "가")
End Function

Private Function Address_LooksLikeHeader(ByVal textValue As String) As Boolean
    textValue = Replace(Trim$(textValue), " ", "")
    Address_LooksLikeHeader = (InStr(1, textValue, "주소", vbTextCompare) > 0)
End Function

Public Sub Address_SplitAdministrativeParts()
    Dim addressRange As Range
    If Not Toolkit_GetRangeFromUserForm("행정구역 분해", _
        "주소 범위를 선택하세요. 오른쪽에 광역시도, 시군구, 읍면동, 상세주소를 생성합니다.", _
        addressRange) Then Exit Sub
    Set addressRange = addressRange.Areas(1)
    If addressRange.Columns.Count > 1 Then Set addressRange = addressRange.Columns(1)

    Dim ws As Worksheet: Set ws = addressRange.Worksheet
    Dim firstRow As Long: firstRow = addressRange.Row
    Dim lastRow As Long: lastRow = addressRange.Row + addressRange.Rows.Count - 1
    Dim dataStart As Long: dataStart = firstRow
    If Address_LooksLikeHeader(CStr(ws.Cells(firstRow, addressRange.Column).Value)) Then dataStart = firstRow + 1

    Dim outCol As Long: outCol = addressRange.Column + addressRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outCol, 4, "행정구역 분해") Then Exit Sub

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStart + 1)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outCol).Resize(, 4).Insert Shift:=xlToRight
    Roster_WriteArray ws, firstRow, outCol, Array("광역시도", "시군구", "읍면동", "상세주소")

    Dim r As Long, parts As Variant
    Toolkit_RowProgressStart "행정구역 분해", totalRows, "행정구역 분해 준비 중"
    For r = dataStart To lastRow
        parts = Address_ParseAdministrativeParts(CStr(ws.Cells(r, addressRange.Column).Value))
        Roster_WriteArray ws, r, outCol, parts
        Toolkit_RowProgressUpdate r - dataStart + 1, totalRows, "행정구역 분해 중: " & CStr(r - dataStart + 1) & "/" & CStr(totalRows) & "행"
    Next r
    ws.Columns(outCol).Resize(, 4).AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox "행정구역 분해를 완료했습니다.", vbInformation, "행정구역 분해"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("행정구역 분해") Then Exit Sub
    MsgBox "행정구역 분해 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "행정구역 분해"
End Sub

Private Function Address_ParseAdministrativeParts(ByVal addressText As String) As Variant
    Dim textValue As String: textValue = Trim$(Address_NormalizeRegionNames(CStr(addressText)))
    Dim tokens() As String: tokens = Split(textValue, " ")
    Dim sido As String, sigungu As String, dong As String, detail As String
    Dim i As Long, startDetail As Long
    If UBound(tokens) >= 0 Then sido = tokens(0)
    If UBound(tokens) >= 1 Then sigungu = tokens(1)
    For i = 2 To UBound(tokens)
        If Address_LooksLikeDongToken(tokens(i)) Then
            dong = tokens(i)
            startDetail = i + 1
            Exit For
        End If
    Next i
    If startDetail = 0 Then startDetail = 2
    For i = startDetail To UBound(tokens)
        detail = Trim$(detail & " " & tokens(i))
    Next i
    Address_ParseAdministrativeParts = Array(sido, sigungu, dong, detail)
End Function
