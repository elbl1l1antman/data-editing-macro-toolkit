Attribute VB_Name = "RosterTools"
Option Explicit


' ============================================================
' 명부 품질/중복/표 메타 도구
' ============================================================
Public Sub Roster_DetectDuplicates()
    Dim modeChoice As VbMsgBoxResult
    modeChoice = Roster_SelectDetectionMode("중복 업체 탐지")
    If modeChoice = vbCancel Then Exit Sub
    If modeChoice = vbNo Then
        Roster_DetectDuplicatesManual
        Exit Sub
    End If

    Dim dataRange As Range
    If Not Toolkit_GetRangeFromUserForm("중복 업체 탐지", _
        "헤더가 포함된 명부 범위를 선택하세요. 업체명, 주소, 전화번호, 사업자등록번호 계열 헤더를 자동 탐지합니다.", _
        dataRange) Then Exit Sub
    Set dataRange = dataRange.Areas(1)

    Dim nameCol As Long, addressCol As Long, phoneCol As Long, bizCol As Long
    nameCol = Roster_FindHeaderColumn(dataRange, Array("업체명", "기업체명", "사업체명", "기업명", "상호"))
    addressCol = Roster_FindHeaderColumn(dataRange, Array("주소", "도로명주소", "지번주소", "소재지"))
    phoneCol = Roster_FindHeaderColumn(dataRange, Array("전화", "전화번호", "대표번호", "연락처"))
    bizCol = Roster_FindHeaderColumn(dataRange, Array("사업자등록번호", "사업자번호", "등록번호"))

    If nameCol = 0 And addressCol = 0 And phoneCol = 0 And bizCol = 0 Then
        MsgBox "중복 판단에 사용할 헤더를 찾지 못했습니다.", vbExclamation, "중복 업체 탐지"
        Exit Sub
    End If

    If dataRange.Rows.Count > 1500 Then
        If MsgBox("선택한 명부가 " & (dataRange.Rows.Count - 1) & "행입니다." & vbCrLf & _
                  "중복 탐지는 행 조합을 비교하므로 시간이 오래 걸릴 수 있습니다. 계속할까요?", _
                  vbQuestion + vbYesNo, "중복 업체 탐지") <> vbYes Then Exit Sub
    End If

    Dim wb As Workbook: Set wb = dataRange.Worksheet.Parent
    Dim outWs As Worksheet
    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = SPSS_UniqueSheetName(wb, "중복탐지")
    Roster_WriteArray outWs, 1, 1, Array("행1", "행2", "중복등급", "점수", "판단근거", "업체명1", "업체명2", "주소1", "주소2", "전화1", "전화2", "사업자번호1", "사업자번호2")

    Dim lastRow As Long: lastRow = dataRange.Row + dataRange.Rows.Count - 1
    Dim r1 As Long, r2 As Long, outRow As Long: outRow = 2
    Dim score As Double, reason As String, grade As String
    Dim maxPairs As Long: maxPairs = 0
    Dim totalRows As Long: totalRows = Application.Max(0, dataRange.Rows.Count - 1)
    Dim pairPulseCount As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "중복 업체 탐지", totalRows, "중복 비교 준비 중"
    For r1 = dataRange.Row + 1 To lastRow - 1
        For r2 = r1 + 1 To lastRow
            score = Roster_DuplicateScore(dataRange.Worksheet, r1, r2, dataRange, nameCol, addressCol, phoneCol, bizCol, reason)
            pairPulseCount = pairPulseCount + 1
            If Toolkit_ShouldShowRowProgress(totalRows) And pairPulseCount >= TOOLKIT_PROGRESS_PAIR_BATCH Then
                pairPulseCount = 0
                DoEvents
                Toolkit_RaiseIfCancelled
            End If
            If score >= 55 Then
                grade = Roster_DuplicateGrade(score)
                Roster_WriteArray outWs, outRow, 1, Array(r1, r2, grade, Round(score, 1), reason, _
                    Roster_CellText(dataRange.Worksheet, r1, dataRange, nameCol), Roster_CellText(dataRange.Worksheet, r2, dataRange, nameCol), _
                    Roster_CellText(dataRange.Worksheet, r1, dataRange, addressCol), Roster_CellText(dataRange.Worksheet, r2, dataRange, addressCol), _
                    Roster_CellText(dataRange.Worksheet, r1, dataRange, phoneCol), Roster_CellText(dataRange.Worksheet, r2, dataRange, phoneCol), _
                    Roster_CellText(dataRange.Worksheet, r1, dataRange, bizCol), Roster_CellText(dataRange.Worksheet, r2, dataRange, bizCol))
                outRow = outRow + 1
                maxPairs = maxPairs + 1
            End If
        Next r2
        Toolkit_RowProgressUpdate r1 - dataRange.Row, totalRows, "중복 비교 중: 기준 행 " & CStr(r1 - dataRange.Row) & "/" & CStr(totalRows)
    Next r1
    outWs.Columns.AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox "중복 업체 탐지를 완료했습니다." & vbCrLf & "의심 쌍: " & maxPairs & "건", vbInformation, "중복 업체 탐지"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("중복 업체 탐지") Then Exit Sub
    MsgBox "중복 업체 탐지 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "중복 업체 탐지"
End Sub

Public Sub Roster_QualityReport()
    Dim modeChoice As VbMsgBoxResult
    modeChoice = Roster_SelectDetectionMode("명부 품질 진단")
    If modeChoice = vbCancel Then Exit Sub
    If modeChoice = vbNo Then
        Roster_QualityReportManual
        Exit Sub
    End If

    Dim dataRange As Range
    If Not Toolkit_GetRangeFromUserForm("명부 품질 진단", _
        "헤더가 포함된 명부 범위를 선택하세요. 빈값, 전화번호, 주소 유형, 사업자번호, 중복 의심을 요약합니다.", _
        dataRange) Then Exit Sub
    Set dataRange = dataRange.Areas(1)

    Dim ws As Worksheet: Set ws = dataRange.Worksheet
    Dim wb As Workbook: Set wb = ws.Parent
    Dim outWs As Worksheet
    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = SPSS_UniqueSheetName(wb, "명부품질진단")

    Dim nameCol As Long, addressCol As Long, phoneCol As Long, bizCol As Long
    nameCol = Roster_FindHeaderColumn(dataRange, Array("업체명", "기업체명", "사업체명", "기업명", "상호"))
    addressCol = Roster_FindHeaderColumn(dataRange, Array("주소", "도로명주소", "지번주소", "소재지"))
    phoneCol = Roster_FindHeaderColumn(dataRange, Array("전화", "전화번호", "대표번호", "연락처"))
    bizCol = Roster_FindHeaderColumn(dataRange, Array("사업자등록번호", "사업자번호", "등록번호"))

    Roster_WriteArray outWs, 1, 1, Array("진단항목", "건수", "설명")
    Dim outRow As Long: outRow = 2
    Dim lastRow As Long: lastRow = dataRange.Row + dataRange.Rows.Count - 1
    Dim totalRows As Long: totalRows = Application.Max(0, dataRange.Rows.Count - 1)
    Roster_WriteArray outWs, outRow, 1, Array("전체 데이터 행", totalRows, "헤더 제외 행 수"): outRow = outRow + 1

    Dim blankName As Long, blankAddress As Long, blankPhone As Long, blankBiz As Long
    Dim invalidPhone As Long, errorPhone As Long, invalidBiz As Long
    Dim roadAddress As Long, dongAddress As Long, unknownAddress As Long
    Dim r As Long, textValue As String, phoneResult As Variant, addressResult As Variant, digits As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "명부 품질 진단", totalRows, "명부 품질 진단 준비 중"
    For r = dataRange.Row + 1 To lastRow
        If nameCol > 0 Then
            If Len(Trim$(CStr(ws.Cells(r, dataRange.Column + nameCol - 1).Value))) = 0 Then blankName = blankName + 1
        End If
        If addressCol > 0 Then
            textValue = Trim$(CStr(ws.Cells(r, dataRange.Column + addressCol - 1).Value))
            If Len(textValue) = 0 Then
                blankAddress = blankAddress + 1
            Else
                addressResult = Address_ClassifyOneAddress(textValue)
                Select Case CStr(addressResult(0))
                    Case "도로명주소": roadAddress = roadAddress + 1
                    Case "법정/행정동 기준 주소": dongAddress = dongAddress + 1
                    Case Else: unknownAddress = unknownAddress + 1
                End Select
            End If
        End If
        If phoneCol > 0 Then
            textValue = Trim$(CStr(ws.Cells(r, dataRange.Column + phoneCol - 1).Value))
            If Len(textValue) = 0 Then
                blankPhone = blankPhone + 1
            Else
                phoneResult = Phone_ClassifyOne(textValue)
                Select Case CStr(phoneResult(1))
                    Case "오류"
                        errorPhone = errorPhone + 1
                    Case "정상"
                        ' 정상 번호는 집계하지 않는다.
                    Case Else
                        invalidPhone = invalidPhone + 1
                End Select
            End If
        End If
        If bizCol > 0 Then
            textValue = Trim$(CStr(ws.Cells(r, dataRange.Column + bizCol - 1).Value))
            If Len(textValue) = 0 Then
                blankBiz = blankBiz + 1
            Else
                digits = DigitsOnly(textValue)
                If Len(digits) > 0 And Len(digits) <> 10 Then invalidBiz = invalidBiz + 1
            End If
        End If
        Toolkit_RowProgressUpdate r - dataRange.Row, totalRows, "명부 품질 진단 중: " & CStr(r - dataRange.Row) & "/" & CStr(totalRows) & "행"
    Next r
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True

    If nameCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("업체명 빈값", blankName, "빈값 여부"): outRow = outRow + 1
    If addressCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("주소 빈값", blankAddress, "빈값 여부"): outRow = outRow + 1
    If phoneCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 빈값", blankPhone, "빈값 여부"): outRow = outRow + 1
    If bizCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("사업자번호 빈값", blankBiz, "빈값 여부"): outRow = outRow + 1
    If phoneCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 오류", errorPhone, "영값, 동일 숫자 6회 이상, 반복 블록, 순차 숫자 패턴"): outRow = outRow + 1
    If phoneCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 확인 필요", invalidPhone, "숫자 길이/복수 대표번호/동일 숫자 5회 반복 기준"): outRow = outRow + 1
    If bizCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("사업자번호 확인 필요", invalidBiz, "숫자 10자리 여부 기준"): outRow = outRow + 1
    If addressCol > 0 Then
        Roster_WriteArray outWs, outRow, 1, Array("도로명주소", roadAddress, "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("법정/행정동 기준 주소", dongAddress, "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("주소 확인 필요", unknownAddress, "주소 유형 판별 기준"): outRow = outRow + 1
    End If

    outWs.Columns.AutoFit
    MsgBox "명부 품질 진단 리포트를 생성했습니다.", vbInformation, "명부 품질 진단"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("명부 품질 진단") Then Exit Sub
    MsgBox "명부 품질 진단 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "명부 품질 진단"
End Sub

Public Sub Roster_CreateValidationWorkspace()
    Dim choice As String
    choice = Toolkit_ButtonMenu( _
        "명부 검증 시트", _
        "명부 검증용 표준 입력 시트를 만드는 방식을 선택하세요.", _
        "1" & vbTab & "빈 표준 입력 시트 만들기" & vbLf & _
        "2" & vbTab & "기존 명부 자동 인식해서 구성", _
        "1")
    If Len(choice) = 0 Then Exit Sub

    Select Case choice
        Case "1": Roster_CreateValidationTemplate
        Case "2": Roster_CreateValidationSheetFromSelection
    End Select
End Sub

Public Sub Roster_CreateValidationTemplate()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = Roster_CreateValidationSheetCore(wb, True)
    ws.Activate

    MsgBox "명부 검증용 표준 입력 시트를 생성했습니다." & vbCrLf & _
           "헤더 아래에 데이터를 붙여넣은 뒤 품질 진단/중복 탐지 기능을 실행하세요.", _
           vbInformation, "명부 검증 시트"
End Sub

Public Sub Roster_CreateValidationSheetFromSelection()
    Dim dataRange As Range
    If Not Toolkit_GetRangeFromUserForm("명부 검증 시트 자동 구성", _
        "헤더가 포함된 원본 명부 범위를 선택하세요." & vbCrLf & _
        "첫 행을 헤더로 보고 ID, 업체명, 연락처, 주소, 이메일, 직무/직책/부서 열을 자동 인식합니다.", _
        dataRange) Then Exit Sub
    If dataRange Is Nothing Then Exit Sub
    Set dataRange = dataRange.Areas(1)

    If dataRange.Rows.Count < 2 Then
        MsgBox "자동 구성에는 헤더 1행과 데이터 1행 이상이 필요합니다.", vbExclamation, "명부 검증 시트"
        Exit Sub
    End If

    Dim headerMap As Object
    Set headerMap = Roster_BuildValidationHeaderMap(dataRange)

    Dim manualIdAnswer As VbMsgBoxResult
    manualIdAnswer = MsgBox("ID/KEY로 사용할 열을 직접 지정할까요?" & vbCrLf & _
                            "예: 선택한 원본 범위 안에서 ID, 관리번호, 사업체 고유번호 등 기준 열을 직접 선택합니다." & vbCrLf & _
                            "아니오: 헤더명 자동 인식 결과를 사용하고, 없으면 자동 ID를 생성합니다.", _
                            vbQuestion + vbYesNoCancel, "명부 검증 시트 자동 구성")
    If manualIdAnswer = vbCancel Then Exit Sub
    If manualIdAnswer = vbYes Then
        If Not Roster_ApplyManualIdColumn(dataRange, headerMap) Then Exit Sub
    End If

    Dim summary As String
    summary = Roster_ValidationMapSummary(dataRange, headerMap)
    If MsgBox(summary & vbCrLf & vbCrLf & "이 구성으로 명부 검증용 표준 시트를 만들까요?", _
              vbQuestion + vbYesNo, "명부 검증 시트 자동 구성") <> vbYes Then Exit Sub

    Dim wb As Workbook
    Set wb = dataRange.Worksheet.Parent
    Dim outWs As Worksheet
    Set outWs = Roster_CreateValidationSheetCore(wb, False)

    Dim headers As Variant
    headers = Roster_ValidationHeaders()

    Dim dataRows As Long
    dataRows = dataRange.Rows.Count - 1

    Dim outData() As Variant
    ReDim outData(1 To dataRows, 1 To UBound(headers) - LBound(headers) + 1)

    Dim r As Long, targetIndex As Long, srcRelCol As Long, targetHeader As String
    For r = 1 To dataRows
        For targetIndex = LBound(headers) To UBound(headers)
            targetHeader = CStr(headers(targetIndex))
            srcRelCol = CLng(headerMap(targetHeader))
            If srcRelCol > 0 Then
                outData(r, targetIndex - LBound(headers) + 1) = dataRange.Cells(r + 1, srcRelCol).Value
            ElseIf targetHeader = "ID" Then
                outData(r, targetIndex - LBound(headers) + 1) = "AUTO_" & dataRange.Worksheet.Name & "_R" & CStr(dataRange.Row + r)
            End If
        Next targetIndex
    Next r

    outWs.Range(outWs.Cells(2, 1), outWs.Cells(dataRows + 1, UBound(headers) - LBound(headers) + 1)).Value = outData
    outWs.Columns.AutoFit
    outWs.Activate

    MsgBox "원본 명부를 표준 검증 시트로 구성했습니다." & vbCrLf & _
           "결과 시트: " & outWs.Name, vbInformation, "명부 검증 시트"
End Sub

Private Function Roster_SelectDetectionMode(ByVal title As String) As VbMsgBoxResult
    Roster_SelectDetectionMode = MsgBox( _
        "진행 방식을 선택하세요." & vbCrLf & vbCrLf & _
        "[예] 자동 탐지: 헤더가 포함된 명부 범위에서 업체명, 주소, 전화번호, 사업자번호 열을 자동으로 찾습니다." & vbCrLf & _
        "[아니오] 수동 선택: 검증 항목과 각 항목의 범위를 직접 선택합니다." & vbCrLf & _
        "[취소] 작업을 중단합니다.", _
        vbQuestion + vbYesNoCancel, title)
End Function

Private Function Roster_SelectManualItems(ByVal title As String, ByVal promptText As String, ByVal maxIndex As Long, ByVal defaultText As String) As Object
    Dim inputText As String
    Dim optionsText As String
    optionsText = Toolkit_OptionsFromNumberedPrompt(promptText)
    If Len(optionsText) > 0 Then
        inputText = Toolkit_CheckListMenu(title, Split(promptText, vbCrLf)(0), optionsText, defaultText)
    Else
        inputText = InputBox(promptText, title, defaultText)
    End If
    Set Roster_SelectManualItems = SPSS_ParseIndexSelection(inputText, maxIndex)
End Function

Private Function Roster_AskHeaderMode(ByVal title As String, ByRef hasHeader As Boolean) As Boolean
    Dim answer As VbMsgBoxResult
    answer = MsgBox("선택할 범위의 첫 행을 헤더로 사용할까요?" & vbCrLf & _
                    "예: 첫 행은 항목명이고 실제 데이터는 다음 행부터입니다." & vbCrLf & _
                    "아니오: 선택한 첫 행부터 데이터로 계산합니다.", _
                    vbQuestion + vbYesNoCancel, title)
    If answer = vbCancel Then Exit Function
    hasHeader = (answer = vbYes)
    Roster_AskHeaderMode = True
End Function

Private Function Roster_FirstDataRow(ByVal rng As Range, ByVal hasHeader As Boolean) As Long
    If hasHeader Then
        Roster_FirstDataRow = rng.Row + 1
    Else
        Roster_FirstDataRow = rng.Row
    End If
End Function

Private Function Roster_RequestManualRange(ByVal title As String, ByVal label As String, ByRef targetRange As Range) As Boolean
    If Not Toolkit_GetRangeFromUserForm(title & " - " & label, _
        label & " 범위를 선택하세요. 여러 열을 선택하면 첫 번째 열만 사용합니다." & vbCrLf & _
        "여러 항목을 선택하는 경우 각 범위는 같은 시작 행과 같은 행 수로 맞춰주세요.", _
        targetRange) Then Exit Function

    If targetRange.Areas.Count > 1 Then
        MsgBox "수동 선택 범위는 한 덩어리의 연속 범위만 지원합니다." & vbCrLf & _
               "떨어진 여러 범위가 아니라, 하나의 열 또는 하나의 연속 범위를 다시 선택하세요.", _
               vbExclamation, title
        Set targetRange = Nothing
        Exit Function
    End If

    Set targetRange = targetRange.Areas(1).Columns(1)
    Roster_RequestManualRange = True
End Function

Private Sub Roster_QualityReportManual()
    Dim selectedItems As Object
    Set selectedItems = Roster_SelectManualItems( _
        "명부 품질 진단 - 검증 항목 선택", _
        "검증할 항목 번호를 입력하세요." & vbCrLf & vbCrLf & _
        "1. 업체명 빈값" & vbCrLf & _
        "2. 주소 빈값" & vbCrLf & _
        "3. 전화번호 유효성" & vbCrLf & _
        "4. 사업자번호 유효성" & vbCrLf & _
        "5. 주소 유형 판별" & vbCrLf & vbCrLf & _
        "예: 1,2,3,5 / 비워두면 취소", 5, "1,2,3,4,5")
    If selectedItems.Count = 0 Then Exit Sub

    Dim hasHeader As Boolean
    If Not Roster_AskHeaderMode("명부 품질 진단 - 헤더 설정", hasHeader) Then Exit Sub

    Dim nameRange As Range, addressRange As Range, phoneRange As Range, bizRange As Range
    If selectedItems.Exists("1") Then
        If Not Roster_RequestManualRange("명부 품질 진단", "업체명", nameRange) Then Exit Sub
    End If
    If selectedItems.Exists("2") Or selectedItems.Exists("5") Then
        If Not Roster_RequestManualRange("명부 품질 진단", "주소", addressRange) Then Exit Sub
    End If
    If selectedItems.Exists("3") Then
        If Not Roster_RequestManualRange("명부 품질 진단", "전화번호", phoneRange) Then Exit Sub
    End If
    If selectedItems.Exists("4") Then
        If Not Roster_RequestManualRange("명부 품질 진단", "사업자번호", bizRange) Then Exit Sub
    End If

    Dim baseWs As Worksheet
    Set baseWs = Roster_FirstRangeWorksheet(nameRange, addressRange, phoneRange, bizRange)
    If baseWs Is Nothing Then Exit Sub
    If Not Roster_RangesOnSameWorksheet(baseWs, nameRange, addressRange, phoneRange, bizRange) Then
        MsgBox "수동 선택한 범위는 모두 같은 시트에 있어야 합니다.", vbExclamation, "명부 품질 진단"
        Exit Sub
    End If
    If Not Roster_ManualRangesAligned("명부 품질 진단", nameRange, addressRange, phoneRange, bizRange) Then Exit Sub

    Dim wb As Workbook: Set wb = baseWs.Parent
    Dim outWs As Worksheet
    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = SPSS_UniqueSheetName(wb, "명부품질진단")

    Roster_WriteArray outWs, 1, 1, Array("진단항목", "건수", "설명")
    Dim outRow As Long: outRow = 2
    Roster_WriteArray outWs, outRow, 1, Array("검증 방식", "수동 선택", IIf(hasHeader, "첫 행 헤더 제외", "첫 행부터 데이터")): outRow = outRow + 1

    Dim baseRange As Range
    Set baseRange = Roster_FirstAvailableRange(nameRange, addressRange, phoneRange, bizRange)
    Dim firstDataRow As Long: firstDataRow = Roster_FirstDataRow(baseRange, hasHeader)
    Dim lastRow As Long: lastRow = baseRange.Row + baseRange.Rows.Count - 1
    Dim totalRows As Long: totalRows = Application.Max(0, lastRow - firstDataRow + 1)

    Dim blankName As Long, blankAddress As Long
    Dim invalidPhone As Long, errorPhone As Long, invalidBiz As Long
    Dim roadAddress As Long, dongAddress As Long, unknownAddress As Long
    Dim r As Long, textValue As String, result As Variant, digits As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "명부 품질 진단", totalRows, "수동 선택 명부 품질 진단 준비 중"
    For r = firstDataRow To lastRow
        If selectedItems.Exists("1") Then
            If Len(Roster_RangeTextAtRow(nameRange, r)) = 0 Then blankName = blankName + 1
        End If
        If selectedItems.Exists("2") Or selectedItems.Exists("5") Then
            textValue = Roster_RangeTextAtRow(addressRange, r)
            If selectedItems.Exists("2") And Len(textValue) = 0 Then blankAddress = blankAddress + 1
            If selectedItems.Exists("5") And Len(textValue) > 0 Then
                result = Address_ClassifyOneAddress(textValue)
                Select Case CStr(result(0))
                    Case "도로명주소": roadAddress = roadAddress + 1
                    Case "법정/행정동 기준 주소": dongAddress = dongAddress + 1
                    Case Else: unknownAddress = unknownAddress + 1
                End Select
            End If
        End If
        If selectedItems.Exists("3") Then
            textValue = Roster_RangeTextAtRow(phoneRange, r)
            If Len(textValue) > 0 Then
                result = Phone_ClassifyOne(textValue)
                Select Case CStr(result(1))
                    Case "오류"
                        errorPhone = errorPhone + 1
                    Case "정상"
                        ' 정상 번호는 집계하지 않는다.
                    Case Else
                        invalidPhone = invalidPhone + 1
                End Select
            End If
        End If
        If selectedItems.Exists("4") Then
            digits = DigitsOnly(Roster_RangeTextAtRow(bizRange, r))
            If Len(digits) > 0 And Len(digits) <> 10 Then invalidBiz = invalidBiz + 1
        End If
        Toolkit_RowProgressUpdate r - firstDataRow + 1, totalRows, "수동 품질 진단 중: " & CStr(r - firstDataRow + 1) & "/" & CStr(totalRows) & "행"
    Next r
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True

    If selectedItems.Exists("1") Then Roster_WriteArray outWs, outRow, 1, Array("업체명 빈값", blankName, "수동 선택 범위 빈값 여부"): outRow = outRow + 1
    If selectedItems.Exists("2") Then Roster_WriteArray outWs, outRow, 1, Array("주소 빈값", blankAddress, "수동 선택 범위 빈값 여부"): outRow = outRow + 1
    If selectedItems.Exists("3") Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 오류", errorPhone, "영값, 동일 숫자 6회 이상, 반복 블록, 순차 숫자 패턴"): outRow = outRow + 1
    If selectedItems.Exists("3") Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 확인 필요", invalidPhone, "숫자 길이/복수 대표번호/동일 숫자 5회 반복 기준"): outRow = outRow + 1
    If selectedItems.Exists("4") Then Roster_WriteArray outWs, outRow, 1, Array("사업자번호 확인 필요", invalidBiz, "숫자 10자리 여부 기준"): outRow = outRow + 1
    If selectedItems.Exists("5") Then
        Roster_WriteArray outWs, outRow, 1, Array("도로명주소", roadAddress, "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("법정/행정동 기준 주소", dongAddress, "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("주소 확인 필요", unknownAddress, "주소 유형 판별 기준"): outRow = outRow + 1
    End If

    outWs.Columns.AutoFit
    MsgBox "수동 선택 기반 명부 품질 진단 리포트를 생성했습니다." & vbCrLf & outWs.Name, vbInformation, "명부 품질 진단"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("명부 품질 진단") Then Exit Sub
    MsgBox "수동 선택 기반 명부 품질 진단 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "명부 품질 진단"
End Sub

Private Sub Roster_DetectDuplicatesManual()
    Dim selectedItems As Object
    Set selectedItems = Roster_SelectManualItems( _
        "중복 업체 탐지 - 비교 항목 선택", _
        "중복 판단에 사용할 항목 번호를 입력하세요." & vbCrLf & vbCrLf & _
        "1. 업체명" & vbCrLf & _
        "2. 주소" & vbCrLf & _
        "3. 전화번호" & vbCrLf & _
        "4. 사업자번호" & vbCrLf & vbCrLf & _
        "예: 1,2,3,4 / 비워두면 취소", 4, "1,2,3,4")
    If selectedItems.Count = 0 Then Exit Sub

    Dim hasHeader As Boolean
    If Not Roster_AskHeaderMode("중복 업체 탐지 - 헤더 설정", hasHeader) Then Exit Sub

    Dim nameRange As Range, addressRange As Range, phoneRange As Range, bizRange As Range
    If selectedItems.Exists("1") Then
        If Not Roster_RequestManualRange("중복 업체 탐지", "업체명", nameRange) Then Exit Sub
    End If
    If selectedItems.Exists("2") Then
        If Not Roster_RequestManualRange("중복 업체 탐지", "주소", addressRange) Then Exit Sub
    End If
    If selectedItems.Exists("3") Then
        If Not Roster_RequestManualRange("중복 업체 탐지", "전화번호", phoneRange) Then Exit Sub
    End If
    If selectedItems.Exists("4") Then
        If Not Roster_RequestManualRange("중복 업체 탐지", "사업자번호", bizRange) Then Exit Sub
    End If

    Dim baseRange As Range
    Set baseRange = Roster_FirstAvailableRange(nameRange, addressRange, phoneRange, bizRange)
    If baseRange Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = baseRange.Worksheet
    If Not Roster_RangesOnSameWorksheet(ws, nameRange, addressRange, phoneRange, bizRange) Then
        MsgBox "수동 선택한 범위는 모두 같은 시트에 있어야 합니다.", vbExclamation, "중복 업체 탐지"
        Exit Sub
    End If
    If Not Roster_ManualRangesAligned("중복 업체 탐지", nameRange, addressRange, phoneRange, bizRange) Then Exit Sub

    Dim firstDataRow As Long: firstDataRow = Roster_FirstDataRow(baseRange, hasHeader)
    Dim lastRow As Long: lastRow = baseRange.Row + baseRange.Rows.Count - 1
    If firstDataRow >= lastRow Then
        MsgBox "중복 비교에는 데이터 행이 2개 이상 필요합니다.", vbExclamation, "중복 업체 탐지"
        Exit Sub
    End If

    If (lastRow - firstDataRow + 1) > 1500 Then
        If MsgBox("선택한 범위가 " & (lastRow - firstDataRow + 1) & "행입니다." & vbCrLf & _
                  "중복 탐지는 행 조합을 비교하므로 시간이 오래 걸릴 수 있습니다. 계속할까요?", _
                  vbQuestion + vbYesNo, "중복 업체 탐지") <> vbYes Then Exit Sub
    End If

    Dim wb As Workbook: Set wb = ws.Parent
    Dim outWs As Worksheet
    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = SPSS_UniqueSheetName(wb, "중복탐지")
    Roster_WriteArray outWs, 1, 1, Array("행1", "행2", "중복등급", "점수", "판단근거", "업체명1", "업체명2", "주소1", "주소2", "전화1", "전화2", "사업자번호1", "사업자번호2")

    Dim r1 As Long, r2 As Long, outRow As Long: outRow = 2
    Dim score As Double, reason As String, grade As String, maxPairs As Long
    Dim totalRows As Long: totalRows = Application.Max(0, lastRow - firstDataRow + 1)
    Dim pairPulseCount As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "중복 업체 탐지", totalRows, "수동 선택 중복 비교 준비 중"
    For r1 = firstDataRow To lastRow - 1
        For r2 = r1 + 1 To lastRow
            score = Roster_DuplicateScoreManual(ws, r1, r2, nameRange, addressRange, phoneRange, bizRange, reason)
            pairPulseCount = pairPulseCount + 1
            If Toolkit_ShouldShowRowProgress(totalRows) And pairPulseCount >= TOOLKIT_PROGRESS_PAIR_BATCH Then
                pairPulseCount = 0
                DoEvents
                Toolkit_RaiseIfCancelled
            End If
            If score >= 55 Then
                grade = Roster_DuplicateGrade(score)
                Roster_WriteArray outWs, outRow, 1, Array(r1, r2, grade, Round(score, 1), reason, _
                    Roster_RangeTextAtRow(nameRange, r1), Roster_RangeTextAtRow(nameRange, r2), _
                    Roster_RangeTextAtRow(addressRange, r1), Roster_RangeTextAtRow(addressRange, r2), _
                    Roster_RangeTextAtRow(phoneRange, r1), Roster_RangeTextAtRow(phoneRange, r2), _
                    Roster_RangeTextAtRow(bizRange, r1), Roster_RangeTextAtRow(bizRange, r2))
                outRow = outRow + 1
                maxPairs = maxPairs + 1
            End If
        Next r2
        Toolkit_RowProgressUpdate r1 - firstDataRow + 1, totalRows, "수동 중복 비교 중: 기준 행 " & CStr(r1 - firstDataRow + 1) & "/" & CStr(totalRows)
    Next r1
    outWs.Columns.AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox "수동 선택 기반 중복 업체 탐지를 완료했습니다." & vbCrLf & "의심 쌍: " & maxPairs & "건", vbInformation, "중복 업체 탐지"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("중복 업체 탐지") Then Exit Sub
    MsgBox "수동 선택 기반 중복 업체 탐지 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "중복 업체 탐지"
End Sub

Public Sub Roster_MatchCompanyNames()
    Dim nameRange As Range
    If Not Toolkit_GetRangeFromUserForm("업체명 유사도 매칭", _
        "업체명 범위를 선택하세요. 원본 오른쪽에 표준화명과 유사그룹 열을 생성합니다.", _
        nameRange) Then Exit Sub
    Set nameRange = nameRange.Areas(1)
    If nameRange.Columns.Count > 1 Then Set nameRange = nameRange.Columns(1)

    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(nameRange)
    If workRange Is Nothing Then Exit Sub

    Dim ws As Worksheet: Set ws = workRange.Worksheet
    Dim firstRow As Long: firstRow = workRange.Row
    Dim lastRow As Long: lastRow = workRange.Row + workRange.Rows.Count - 1
    Dim dataStart As Long: dataStart = firstRow
    If InStr(1, CStr(ws.Cells(firstRow, workRange.Column).Value), "명", vbTextCompare) > 0 Then dataStart = firstRow + 1

    Dim outCol As Long: outCol = workRange.Column + workRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outCol, 2, "업체명 유사도 매칭") Then Exit Sub

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStart + 1)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outCol).Resize(, 2).Insert Shift:=xlToRight
    ws.Cells(firstRow, outCol).Value = "업체명 표준화"
    ws.Cells(firstRow, outCol + 1).Value = "유사그룹"

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, key As String, groupNo As Long, groupLabel As String
    Toolkit_RowProgressStart "업체명 유사도 매칭", totalRows, "업체명 매칭 준비 중"
    For r = dataStart To lastRow
        key = Roster_NormalizeCompanyName(CStr(ws.Cells(r, workRange.Column).Value))
        ws.Cells(r, outCol).Value = key
        If Len(key) > 0 Then
            groupLabel = Roster_FindCompanyGroupLabel(dict, key)
            If Len(groupLabel) = 0 Then
                groupNo = groupNo + 1
                groupLabel = "G" & Format(groupNo, "0000")
            End If
            If Not dict.Exists(key) Then dict(key) = groupLabel
            ws.Cells(r, outCol + 1).Value = groupLabel
        End If
        Toolkit_RowProgressUpdate r - dataStart + 1, totalRows, "업체명 매칭 중: " & CStr(r - dataStart + 1) & "/" & CStr(totalRows) & "행"
    Next r
    ws.Columns(outCol).Resize(, 2).AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox "업체명 표준화와 유사그룹 생성을 완료했습니다.", vbInformation, "업체명 유사도 매칭"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("업체명 유사도 매칭") Then Exit Sub
    MsgBox "업체명 유사도 매칭 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "업체명 유사도 매칭"
End Sub

Private Function Roster_FindHeaderColumn(ByVal dataRange As Range, ByVal candidates As Variant) As Long
    Dim c As Long, i As Long, headerText As String
    For c = 1 To dataRange.Columns.Count
        headerText = Replace(Trim$(CStr(dataRange.Cells(1, c).Value)), " ", "")
        For i = LBound(candidates) To UBound(candidates)
            If InStr(1, headerText, Replace(CStr(candidates(i)), " ", ""), vbTextCompare) > 0 Then
                Roster_FindHeaderColumn = c
                Exit Function
            End If
        Next i
    Next c
End Function

Private Function Roster_ValidationHeaders() As Variant
    Roster_ValidationHeaders = Array("ID", "사업체명", "기업체명", "이름", "연락처1", "연락처2", "연락처3", _
                                     "주소1", "주소2", "주소3", "이메일", "직무", "직책", "소속부서", "기타정보")
End Function

Private Function Roster_ValidationExampleRow() As Variant
    Roster_ValidationExampleRow = Array("R001", "예시상사", "예시기업", "홍길동", "02-000-0000", "010-0000-0000", "", _
                                        "서울특별시 중구 세종대로 110", "", "", "sample@example.com", "조사 응답", "팀장", "경영지원팀", "비고")
End Function

Private Function Roster_CreateValidationSheetCore(ByVal wb As Workbook, ByVal includeExample As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = SPSS_UniqueSheetName(wb, "명부검증_입력")

    Dim headers As Variant
    headers = Roster_ValidationHeaders()
    Roster_WriteArray ws, 1, 1, headers
    If includeExample Then Roster_WriteArray ws, 2, 1, Roster_ValidationExampleRow()

    Dim lastCol As Long
    lastCol = UBound(headers) - LBound(headers) + 1
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(220, 235, 250)
        .Borders.LineStyle = xlContinuous
    End With
    If includeExample Then
        With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
            .Interior.Color = RGB(245, 245, 245)
            .Font.Color = RGB(120, 120, 120)
        End With
    End If

    ws.Columns.AutoFit
    Set Roster_CreateValidationSheetCore = ws
End Function

Private Function Roster_BuildValidationHeaderMap(ByVal dataRange As Range) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim usedColumns As Object
    Set usedColumns = CreateObject("Scripting.Dictionary")

    Dim headers As Variant
    headers = Roster_ValidationHeaders()

    Dim i As Long, targetHeader As String, relCol As Long
    For i = LBound(headers) To UBound(headers)
        targetHeader = CStr(headers(i))
        relCol = Roster_FindHeaderColumnUnused(dataRange, Roster_ValidationCandidates(targetHeader), usedColumns)
        result.Add targetHeader, relCol
        If relCol > 0 Then usedColumns(CStr(relCol)) = True
    Next i

    Set Roster_BuildValidationHeaderMap = result
End Function

Private Function Roster_ApplyManualIdColumn(ByVal dataRange As Range, ByVal headerMap As Object) As Boolean
    Dim keyRange As Range
    dataRange.Worksheet.Activate

    On Error Resume Next
    Set keyRange = Application.InputBox( _
        "ID/KEY로 사용할 열을 선택하세요." & vbCrLf & _
        "처음 선택한 원본 범위 안의 한 개 열만 사용할 수 있습니다.", _
        "명부 검증 시트 - ID 열 직접 지정", Type:=8)
    On Error GoTo 0

    If keyRange Is Nothing Then Exit Function
    Set keyRange = keyRange.Areas(1)

    If Not keyRange.Worksheet Is dataRange.Worksheet Then
        MsgBox "ID/KEY 열은 처음 선택한 원본 명부 범위와 같은 시트에서 선택해야 합니다.", vbExclamation, "명부 검증 시트"
        Exit Function
    End If
    If keyRange.Columns.Count <> 1 Then
        MsgBox "ID/KEY 열은 한 개 열만 선택하세요.", vbExclamation, "명부 검증 시트"
        Exit Function
    End If

    Dim relCol As Long
    relCol = Roster_RelativeColumnWithinRange(dataRange.Column, dataRange.Columns.Count, keyRange.Column)
    If relCol = 0 Then
        MsgBox "선택한 ID/KEY 열이 처음 선택한 원본 명부 범위 안에 없습니다.", vbExclamation, "명부 검증 시트"
        Exit Function
    End If

    headerMap("ID") = relCol
    Roster_ApplyManualIdColumn = True
End Function

Private Function Roster_RelativeColumnWithinRange(ByVal dataFirstCol As Long, ByVal dataColCount As Long, ByVal selectedCol As Long) As Long
    If dataFirstCol < 1 Or dataColCount < 1 Or selectedCol < dataFirstCol Then Exit Function
    If selectedCol > dataFirstCol + dataColCount - 1 Then Exit Function
    Roster_RelativeColumnWithinRange = selectedCol - dataFirstCol + 1
End Function

Public Function Roster_ManualKeyRelativeColumnForTest(ByVal dataFirstCol As Long, ByVal dataColCount As Long, ByVal selectedCol As Long) As Long
    Roster_ManualKeyRelativeColumnForTest = Roster_RelativeColumnWithinRange(dataFirstCol, dataColCount, selectedCol)
End Function

Private Function Roster_ValidationCandidates(ByVal targetHeader As String) As Variant
    Select Case targetHeader
        Case "ID"
            Roster_ValidationCandidates = Array("ID", "아이디", "관리번호", "고유번호", "순번", "KEY", "키")
        Case "사업체명"
            Roster_ValidationCandidates = Array("사업체명", "사업장명", "업체명", "상호", "상호명", "사업자명")
        Case "기업체명"
            Roster_ValidationCandidates = Array("기업체명", "기업명", "회사명", "법인명", "업체명")
        Case "이름"
            Roster_ValidationCandidates = Array("이름", "성명", "대표자", "대표자명", "담당자", "담당자명", "응답자")
        Case "연락처1"
            Roster_ValidationCandidates = Array("연락처1", "전화번호1", "대표번호", "전화", "전화번호", "연락처", "TEL", "Mobile", "휴대폰", "핸드폰")
        Case "연락처2"
            Roster_ValidationCandidates = Array("연락처2", "전화번호2", "보조연락처", "보조전화", "휴대폰2", "팩스", "FAX")
        Case "연락처3"
            Roster_ValidationCandidates = Array("연락처3", "전화번호3", "기타연락처", "기타전화")
        Case "주소1"
            Roster_ValidationCandidates = Array("주소1", "주소", "도로명주소", "소재지", "사업장주소", "본사주소")
        Case "주소2"
            Roster_ValidationCandidates = Array("주소2", "지번주소", "상세주소", "주소상세", "추가주소")
        Case "주소3"
            Roster_ValidationCandidates = Array("주소3", "기타주소")
        Case "이메일"
            Roster_ValidationCandidates = Array("이메일", "메일", "전자우편", "E-mail", "Email")
        Case "직무"
            Roster_ValidationCandidates = Array("직무", "업무", "담당업무", "역할")
        Case "직책"
            Roster_ValidationCandidates = Array("직책", "직위", "직급")
        Case "소속부서"
            Roster_ValidationCandidates = Array("소속부서", "부서", "부서명", "소속", "팀")
        Case Else
            Roster_ValidationCandidates = Array("기타정보", "기타", "비고", "메모", "참고")
    End Select
End Function

Private Function Roster_FindHeaderColumnUnused(ByVal dataRange As Range, ByVal candidates As Variant, ByVal usedColumns As Object) As Long
    Dim passNo As Long, c As Long, i As Long
    Dim headerKey As String, candidateKey As String

    For passNo = 1 To 2
        For c = 1 To dataRange.Columns.Count
            If Not usedColumns.Exists(CStr(c)) Then
                headerKey = Roster_NormalizeHeaderForMap(CStr(dataRange.Cells(1, c).Value))
                If Len(headerKey) > 0 Then
                    For i = LBound(candidates) To UBound(candidates)
                        candidateKey = Roster_NormalizeHeaderForMap(CStr(candidates(i)))
                        If Len(candidateKey) > 0 Then
                            If (passNo = 1 And headerKey = candidateKey) _
                               Or (passNo = 2 And InStr(1, headerKey, candidateKey, vbTextCompare) > 0) Then
                                Roster_FindHeaderColumnUnused = c
                                Exit Function
                            End If
                        End If
                    Next i
                End If
            End If
        Next c
    Next passNo
End Function

Private Function Roster_NormalizeHeaderForMap(ByVal headerText As String) As String
    Dim result As String
    result = LCase$(Trim$(CStr(headerText)))
    result = Replace(result, " ", "")
    result = Replace(result, "_", "")
    result = Replace(result, "-", "")
    result = Replace(result, ".", "")
    result = Replace(result, "/", "")
    result = Replace(result, "\", "")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, "[", "")
    result = Replace(result, "]", "")
    Roster_NormalizeHeaderForMap = result
End Function

Private Function Roster_ValidationMapSummary(ByVal dataRange As Range, ByVal headerMap As Object) As String
    Dim headers As Variant
    headers = Roster_ValidationHeaders()

    Dim result As String
    result = "자동 인식 결과" & vbCrLf & _
             "원본: " & dataRange.Worksheet.Name & "!" & dataRange.Address(False, False) & vbCrLf & _
             "데이터 행: " & CStr(dataRange.Rows.Count - 1) & vbCrLf & vbCrLf

    Dim i As Long, targetHeader As String, relCol As Long
    For i = LBound(headers) To UBound(headers)
        targetHeader = CStr(headers(i))
        relCol = CLng(headerMap(targetHeader))
        If relCol > 0 Then
            result = result & "- " & targetHeader & " <= " & CStr(dataRange.Cells(1, relCol).Value) & _
                     " (선택범위 " & CStr(relCol) & "열)" & vbCrLf
        ElseIf targetHeader = "ID" Then
            result = result & "- ID <= 자동 생성" & vbCrLf
        Else
            result = result & "- " & targetHeader & " <= 미인식/빈 열" & vbCrLf
        End If
    Next i

    Roster_ValidationMapSummary = result
End Function

Private Function Roster_CellText(ws As Worksheet, ByVal rowIndex As Long, ByVal dataRange As Range, ByVal relativeCol As Long) As String
    If relativeCol <= 0 Then Exit Function
    Roster_CellText = Trim$(CStr(ws.Cells(rowIndex, dataRange.Column + relativeCol - 1).Value))
End Function

Private Function Roster_DuplicateScore(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal dataRange As Range, ByVal nameCol As Long, ByVal addressCol As Long, ByVal phoneCol As Long, ByVal bizCol As Long, ByRef reason As String) As Double
    Dim score As Double
    reason = ""
    If bizCol > 0 Then
        Dim b1 As String, b2 As String
        b1 = DigitsOnly(Roster_CellText(ws, r1, dataRange, bizCol))
        b2 = DigitsOnly(Roster_CellText(ws, r2, dataRange, bizCol))
        If Len(b1) = 10 And b1 = b2 Then score = score + 100: reason = reason & "사업자번호 일치; "
    End If
    If phoneCol > 0 Then
        Dim p1 As String, p2 As String
        p1 = Phone_DigitsForKoreanRules(Roster_CellText(ws, r1, dataRange, phoneCol))
        p2 = Phone_DigitsForKoreanRules(Roster_CellText(ws, r2, dataRange, phoneCol))
        If Len(p1) >= 8 And p1 = p2 Then score = score + 35: reason = reason & "전화번호 일치; "
    End If
    If nameCol > 0 Then
        Dim ns As Double, nameReason As String
        Dim name1 As String, name2 As String
        name1 = Roster_CellText(ws, r1, dataRange, nameCol)
        name2 = Roster_CellText(ws, r2, dataRange, nameCol)
        If Roster_StrongCompanyNameMatch(name1, name2, nameReason) Then
            score = score + 60
            reason = reason & nameReason & "; "
        Else
            ns = NaverLocal_Similarity(name1, name2)
            If ns >= 85 Then score = score + 35: reason = reason & "업체명 유사; "
        End If
    End If
    If addressCol > 0 Then
        Dim ads As Double
        ads = NaverLocal_Similarity(Roster_CellText(ws, r1, dataRange, addressCol), Roster_CellText(ws, r2, dataRange, addressCol))
        If ads >= 85 Then score = score + 30: reason = reason & "주소 유사; "
    End If
    Roster_DuplicateScore = Application.Min(score, 100)
End Function

Private Function Roster_DuplicateGrade(ByVal score As Double) As String
    If score >= 90 Then
        Roster_DuplicateGrade = "정확 중복"
    ElseIf score >= 75 Then
        Roster_DuplicateGrade = "높음"
    ElseIf score >= 60 Then
        Roster_DuplicateGrade = "보통"
    Else
        Roster_DuplicateGrade = "낮음"
    End If
End Function

Public Sub Roster_WriteArray(ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        ws.Cells(rowIndex, colIndex + i).Value = values(i)
    Next i
End Sub

Private Function Roster_FirstAvailableRange(ByVal nameRange As Range, ByVal addressRange As Range, ByVal phoneRange As Range, ByVal bizRange As Range) As Range
    If Not nameRange Is Nothing Then
        Set Roster_FirstAvailableRange = nameRange
        Exit Function
    End If
    If Not addressRange Is Nothing Then
        Set Roster_FirstAvailableRange = addressRange
        Exit Function
    End If
    If Not phoneRange Is Nothing Then
        Set Roster_FirstAvailableRange = phoneRange
        Exit Function
    End If
    If Not bizRange Is Nothing Then
        Set Roster_FirstAvailableRange = bizRange
        Exit Function
    End If
End Function

Private Function Roster_FirstRangeWorksheet(ByVal nameRange As Range, ByVal addressRange As Range, ByVal phoneRange As Range, ByVal bizRange As Range) As Worksheet
    Dim firstRange As Range
    Set firstRange = Roster_FirstAvailableRange(nameRange, addressRange, phoneRange, bizRange)
    If Not firstRange Is Nothing Then Set Roster_FirstRangeWorksheet = firstRange.Worksheet
End Function

Private Function Roster_RangesOnSameWorksheet(ByVal ws As Worksheet, ByVal nameRange As Range, ByVal addressRange As Range, ByVal phoneRange As Range, ByVal bizRange As Range) As Boolean
    Roster_RangesOnSameWorksheet = False
    If Not nameRange Is Nothing Then If Not nameRange.Worksheet Is ws Then Exit Function
    If Not addressRange Is Nothing Then If Not addressRange.Worksheet Is ws Then Exit Function
    If Not phoneRange Is Nothing Then If Not phoneRange.Worksheet Is ws Then Exit Function
    If Not bizRange Is Nothing Then If Not bizRange.Worksheet Is ws Then Exit Function
    Roster_RangesOnSameWorksheet = True
End Function

Private Function Roster_ManualRangesAligned(ByVal title As String, ByVal nameRange As Range, ByVal addressRange As Range, ByVal phoneRange As Range, ByVal bizRange As Range) As Boolean
    Dim baseRange As Range
    Set baseRange = Roster_FirstAvailableRange(nameRange, addressRange, phoneRange, bizRange)
    If baseRange Is Nothing Then Exit Function

    Dim baseRow As Long: baseRow = baseRange.Row
    Dim baseRows As Long: baseRows = baseRange.Rows.Count
    Dim detail As String
    detail = "기준 범위: " & Roster_RangeSummary(baseRange) & vbCrLf

    If Not Roster_OneRangeAligned(nameRange, baseRow, baseRows, "업체명", detail) Then GoTo Misaligned
    If Not Roster_OneRangeAligned(addressRange, baseRow, baseRows, "주소", detail) Then GoTo Misaligned
    If Not Roster_OneRangeAligned(phoneRange, baseRow, baseRows, "전화번호", detail) Then GoTo Misaligned
    If Not Roster_OneRangeAligned(bizRange, baseRow, baseRows, "사업자번호", detail) Then GoTo Misaligned

    Roster_ManualRangesAligned = True
    Exit Function

Misaligned:
    MsgBox "수동 선택한 범위의 시작 행 또는 행 수가 서로 다릅니다." & vbCrLf & vbCrLf & _
           detail & vbCrLf & _
           "같은 행에 있는 업체 정보끼리 비교하려면 모든 선택 범위를 같은 시작 행과 같은 행 수로 맞춰주세요.", _
           vbExclamation, title
End Function

Private Function Roster_OneRangeAligned(ByVal targetRange As Range, ByVal baseRow As Long, ByVal baseRows As Long, ByVal label As String, ByRef detail As String) As Boolean
    If targetRange Is Nothing Then
        Roster_OneRangeAligned = True
        Exit Function
    End If

    detail = detail & label & ": " & Roster_RangeSummary(targetRange) & vbCrLf
    Roster_OneRangeAligned = (targetRange.Row = baseRow And targetRange.Rows.Count = baseRows)
End Function

Private Function Roster_RangeSummary(ByVal targetRange As Range) As String
    If targetRange Is Nothing Then
        Roster_RangeSummary = "(선택 안 함)"
    Else
        Roster_RangeSummary = targetRange.Worksheet.Name & "!" & targetRange.Address(False, False) & _
                              " / 시작 행 " & targetRange.Row & _
                              " / 행 수 " & targetRange.Rows.Count
    End If
End Function

Private Function Roster_CountBlankRange(ByVal targetRange As Range, ByVal hasHeader As Boolean) As Long
    Dim r As Long
    For r = Roster_FirstDataRow(targetRange, hasHeader) To targetRange.Row + targetRange.Rows.Count - 1
        If Len(Trim$(CStr(targetRange.Worksheet.Cells(r, targetRange.Column).Value))) = 0 Then Roster_CountBlankRange = Roster_CountBlankRange + 1
    Next r
End Function

Private Function Roster_CountInvalidPhonesRange(ByVal targetRange As Range, ByVal hasHeader As Boolean) As Long
    Dim r As Long
    For r = Roster_FirstDataRow(targetRange, hasHeader) To targetRange.Row + targetRange.Rows.Count - 1
        Dim result As Variant
        result = Phone_ClassifyOne(CStr(targetRange.Worksheet.Cells(r, targetRange.Column).Value))
        If CStr(result(1)) <> "정상" And Len(Trim$(CStr(targetRange.Worksheet.Cells(r, targetRange.Column).Value))) > 0 Then Roster_CountInvalidPhonesRange = Roster_CountInvalidPhonesRange + 1
    Next r
End Function

Private Function Roster_CountInvalidBizNoRange(ByVal targetRange As Range, ByVal hasHeader As Boolean) As Long
    Dim r As Long, digits As String
    For r = Roster_FirstDataRow(targetRange, hasHeader) To targetRange.Row + targetRange.Rows.Count - 1
        digits = DigitsOnly(CStr(targetRange.Worksheet.Cells(r, targetRange.Column).Value))
        If Len(digits) > 0 And Len(digits) <> 10 Then Roster_CountInvalidBizNoRange = Roster_CountInvalidBizNoRange + 1
    Next r
End Function

Private Function Roster_CountAddressTypeRange(ByVal targetRange As Range, ByVal hasHeader As Boolean, ByVal targetType As String) As Long
    Dim r As Long, classified As Variant
    For r = Roster_FirstDataRow(targetRange, hasHeader) To targetRange.Row + targetRange.Rows.Count - 1
        classified = Address_ClassifyOneAddress(CStr(targetRange.Worksheet.Cells(r, targetRange.Column).Value))
        If CStr(classified(0)) = targetType Then Roster_CountAddressTypeRange = Roster_CountAddressTypeRange + 1
    Next r
End Function

Private Function Roster_RangeTextAtRow(ByVal targetRange As Range, ByVal rowIndex As Long) As String
    If targetRange Is Nothing Then Exit Function
    If rowIndex < targetRange.Row Or rowIndex > targetRange.Row + targetRange.Rows.Count - 1 Then Exit Function
    Roster_RangeTextAtRow = Trim$(CStr(targetRange.Worksheet.Cells(rowIndex, targetRange.Column).Value))
End Function

Private Function Roster_DuplicateScoreManual(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                             ByVal nameRange As Range, ByVal addressRange As Range, _
                                             ByVal phoneRange As Range, ByVal bizRange As Range, _
                                             ByRef reason As String) As Double
    Dim score As Double, b1 As String, b2 As String, p1 As String, p2 As String
    Dim ns As Double, ads As Double
    reason = ""

    If Not bizRange Is Nothing Then
        b1 = DigitsOnly(Roster_RangeTextAtRow(bizRange, r1))
        b2 = DigitsOnly(Roster_RangeTextAtRow(bizRange, r2))
        If Len(b1) = 10 And b1 = b2 Then score = score + 70: reason = reason & "사업자번호 일치; "
    End If

    If Not phoneRange Is Nothing Then
        p1 = Phone_DigitsForKoreanRules(Roster_RangeTextAtRow(phoneRange, r1))
        p2 = Phone_DigitsForKoreanRules(Roster_RangeTextAtRow(phoneRange, r2))
        If Len(p1) >= 7 And p1 = p2 Then score = score + 45: reason = reason & "전화번호 일치; "
    End If

    If Not nameRange Is Nothing Then
        Dim nameReason As String
        Dim name1 As String, name2 As String
        name1 = Roster_RangeTextAtRow(nameRange, r1)
        name2 = Roster_RangeTextAtRow(nameRange, r2)
        If Roster_StrongCompanyNameMatch(name1, name2, nameReason) Then
            score = score + 60
            reason = reason & nameReason & "; "
        Else
            ns = NaverLocal_Similarity(name1, name2)
            score = score + ns * 0.35
            If ns >= 85 Then reason = reason & "업체명 유사; "
        End If
    End If

    If Not addressRange Is Nothing Then
        ads = NaverLocal_Similarity(Roster_RangeTextAtRow(addressRange, r1), Roster_RangeTextAtRow(addressRange, r2))
        score = score + ads * 0.25
        If ads >= 80 Then reason = reason & "주소 유사; "
    End If

    Roster_DuplicateScoreManual = Application.Min(score, 100)
End Function

Private Function Roster_WriteMissingSummary(outWs As Worksheet, ByVal outRow As Long, ws As Worksheet, dataRange As Range, ByVal relCol As Long, ByVal label As String) As Long
    If relCol > 0 Then
        Roster_WriteArray outWs, outRow, 1, Array(label, Roster_CountBlank(ws, dataRange, relCol), "빈값 여부")
        Roster_WriteMissingSummary = outRow + 1
    Else
        Roster_WriteMissingSummary = outRow
    End If
End Function

Private Function Roster_CountBlank(ws As Worksheet, dataRange As Range, ByVal relCol As Long) As Long
    Dim r As Long, lastRow As Long
    lastRow = dataRange.Row + dataRange.Rows.Count - 1
    For r = dataRange.Row + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, dataRange.Column + relCol - 1).Value))) = 0 Then Roster_CountBlank = Roster_CountBlank + 1
    Next r
End Function

Private Function Roster_CountInvalidPhones(ws As Worksheet, dataRange As Range, ByVal relCol As Long) As Long
    Dim r As Long, result As Variant, lastRow As Long
    lastRow = dataRange.Row + dataRange.Rows.Count - 1
    For r = dataRange.Row + 1 To lastRow
        result = Phone_ClassifyOne(CStr(ws.Cells(r, dataRange.Column + relCol - 1).Value))
        If CStr(result(1)) <> "정상" And Len(Trim$(CStr(ws.Cells(r, dataRange.Column + relCol - 1).Value))) > 0 Then Roster_CountInvalidPhones = Roster_CountInvalidPhones + 1
    Next r
End Function

Private Function Roster_CountInvalidBizNo(ws As Worksheet, dataRange As Range, ByVal relCol As Long) As Long
    Dim r As Long, digits As String, lastRow As Long
    lastRow = dataRange.Row + dataRange.Rows.Count - 1
    For r = dataRange.Row + 1 To lastRow
        digits = DigitsOnly(CStr(ws.Cells(r, dataRange.Column + relCol - 1).Value))
        If Len(digits) > 0 And Len(digits) <> 10 Then Roster_CountInvalidBizNo = Roster_CountInvalidBizNo + 1
    Next r
End Function

Private Function Roster_CountAddressType(ws As Worksheet, dataRange As Range, ByVal relCol As Long, ByVal targetType As String) As Long
    Dim r As Long, classified As Variant, lastRow As Long
    lastRow = dataRange.Row + dataRange.Rows.Count - 1
    For r = dataRange.Row + 1 To lastRow
        classified = Address_ClassifyOneAddress(CStr(ws.Cells(r, dataRange.Column + relCol - 1).Value))
        If CStr(classified(0)) = targetType Then Roster_CountAddressType = Roster_CountAddressType + 1
    Next r
End Function

Private Function Roster_NormalizeCompanyName(ByVal textValue As String) As String
    Dim options As Object
    Set options = CompanyName_DefaultOptions()
    Roster_NormalizeCompanyName = CompanyName_CleanOne(textValue, options)
End Function

Private Function Roster_FindCompanyGroupLabel(ByVal groups As Object, ByVal companyKey As String) As String
    Dim existingKey As Variant
    Dim unusedReason As String
    For Each existingKey In groups.Keys
        If Roster_StrongCompanyNameMatch(CStr(existingKey), companyKey, unusedReason) Then
            Roster_FindCompanyGroupLabel = CStr(groups(existingKey))
            Exit Function
        End If
    Next existingKey
End Function

Private Function Roster_StrongCompanyNameMatch(ByVal leftName As String, ByVal rightName As String, ByRef matchReason As String) As Boolean
    Dim leftKey As String, rightKey As String
    leftKey = Roster_CompanyMatchKey(leftName)
    rightKey = Roster_CompanyMatchKey(rightName)

    If Len(leftKey) < 3 Or Len(rightKey) < 3 Then Exit Function

    If leftKey = rightKey Then
        matchReason = "업체명 표준화 일치"
        Roster_StrongCompanyNameMatch = True
    ElseIf InStr(1, leftKey, rightKey, vbTextCompare) > 0 Or InStr(1, rightKey, leftKey, vbTextCompare) > 0 Then
        matchReason = "업체명 포함 관계"
        Roster_StrongCompanyNameMatch = True
    ElseIf Roster_CommonPrefixLength(leftKey, rightKey) >= 3 Then
        matchReason = "업체명 3자 이상 접두어 일치"
        Roster_StrongCompanyNameMatch = True
    End If
End Function

Private Function Roster_CompanyMatchKey(ByVal textValue As String) As String
    Dim key As String
    key = UCase$(Roster_NormalizeCompanyName(textValue))
    key = Replace(key, "유한회사", "")
    key = Replace(key, "합자회사", "")
    key = Replace(key, "합명회사", "")
    key = Replace(key, "사단법인", "")
    key = Replace(key, "재단법인", "")
    key = Replace(key, " ", "")
    key = Replace(key, ChrW(160), "")
    key = Replace(key, vbTab, "")
    key = Replace(key, "-", "")
    key = Replace(key, "_", "")
    key = Replace(key, ".", "")
    key = Replace(key, ",", "")
    key = Replace(key, "/", "")
    key = Replace(key, "\", "")
    key = Replace(key, "(", "")
    key = Replace(key, ")", "")
    key = Replace(key, "[", "")
    key = Replace(key, "]", "")
    key = Replace(key, "·", "")
    key = Replace(key, "ㆍ", "")
    key = Replace(key, "㈜", "")
    Roster_CompanyMatchKey = Trim$(key)
End Function

Private Function Roster_CommonPrefixLength(ByVal leftKey As String, ByVal rightKey As String) As Long
    Dim limitLen As Long
    limitLen = Application.Min(Len(leftKey), Len(rightKey))

    Dim i As Long
    For i = 1 To limitLen
        If Mid$(leftKey, i, 1) <> Mid$(rightKey, i, 1) Then Exit Function
        Roster_CommonPrefixLength = i
    Next i
End Function
