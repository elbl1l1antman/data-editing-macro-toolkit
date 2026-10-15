Attribute VB_Name = "DataCleanupTools"
Option Explicit


' ============================================================
' 프로시저 : ConvertTextNumbersInSelection
' 설  명  : 사용자가 선택한 셀 범위에서 텍스트 형식 숫자를
'           실제 숫자로 변환한다.
'           실행 전에 반드시 대상 셀을 선택해야 한다.
' ============================================================
Public Sub ConvertTextNumbersInSelection()
    ' 선택 영역이 Range 타입인지 확인 (차트, 도형 등 선택 시 오류 방지)
    If TypeName(Selection) <> "Range" Then
        MsgBox "변환할 셀을 먼저 선택하세요.", vbExclamation, "매크로 툴킷"
        Exit Sub
    End If
    ConvertTextNumbersInRange Selection
End Sub

' ============================================================
' 프로시저 : ConvertTextNumbersInUsedRange
' 설  명  : 현재 활성 시트의 사용 범위(UsedRange) 전체에서
'           텍스트 형식 숫자를 실제 숫자로 일괄 변환한다.
' ============================================================
Public Sub ConvertTextNumbersInUsedRange()
    ConvertTextNumbersInCurrentSheetDataRange
End Sub

' ============================================================
' 프로시저 : ConvertTextNumbersInCurrentSheetDataRange
' 설  명  : 현재 활성 시트에서 실제 데이터가 있는 마지막 행/열까지만 잡아
'           텍스트 형식 숫자를 실제 숫자로 일괄 변환한다.
' ============================================================
Public Sub ConvertTextNumbersInCurrentSheetDataRange()
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "현재 활성화된 항목이 워크시트가 아닙니다.", vbExclamation, "시트 전체 숫자 변환"
        Exit Sub
    End If

    Dim targetRange As Range
    Set targetRange = Toolkit_GetActualDataRange(ActiveSheet)
    If targetRange Is Nothing Then
        MsgBox "현재 시트에서 변환할 데이터가 있는 범위를 찾지 못했습니다.", vbInformation, "시트 전체 숫자 변환"
        Exit Sub
    End If

    If MsgBox("현재 시트의 실제 데이터 영역만 변환합니다." & vbCrLf & _
              "범위: " & targetRange.Address(False, False) & vbCrLf & vbCrLf & _
              "계속할까요?", vbQuestion + vbYesNo, "시트 전체 숫자 변환") = vbNo Then Exit Sub

    ConvertTextNumbersInRange targetRange
End Sub

' ============================================================
' 프로시저 : ConvertTextNumbersInColumns
' 설  명  : 사용자가 지정한 범위에서 텍스트 형식 숫자를
'           실제 숫자로 변환한다.
'           범위 지정 방법 (둘 다 지원):
'             - 마우스로 셀/열을 클릭하거나 드래그하여 선택
'             - 입력창에 범위 주소를 직접 입력 (예: A:A, B1:B100)
'           Ctrl 키를 누른 채 클릭하면 복수 영역 선택 가능
' ============================================================
Public Sub ConvertTextNumbersInColumns()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("숫자 변환", _
        "엑셀 시트에서 변환할 범위를 먼저 선택하세요. Ctrl+Shift+방향키, Shift+방향키, 마우스 드래그를 모두 사용할 수 있습니다.", _
        targetRange) Then Exit Sub
    ConvertTextNumbersInRange targetRange
End Sub

Public Sub ConvertTextNumbersWithForm()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("숫자 변환", _
        "텍스트로 저장된 숫자를 실제 숫자값으로 변환합니다. 변환할 셀/열/범위를 선택하세요.", _
        targetRange) Then Exit Sub
    ConvertTextNumbersInRange targetRange
End Sub

' ============================================================
' 프로시저 : NormalizePhonesInSelection
' 설  명  : 사용자가 선택한 셀 범위의 전화번호를
'           표준 형식(지역번호-국번-번호)으로 정규화한다.
'           실행 전에 반드시 대상 셀을 선택해야 한다.
' ============================================================
Public Sub NormalizePhonesInSelection()
    ' 선택 영역이 Range 타입인지 확인
    If TypeName(Selection) <> "Range" Then
        MsgBox "전화번호 셀을 먼저 선택하세요.", vbExclamation, "매크로 툴킷"
        Exit Sub
    End If

    Dim options As Object
    Set options = Phone_GetNormalizeOptionsFromUser()
    NormalizePhonesToCopiedColumns Selection, options
End Sub

' ============================================================
' 프로시저 : NormalizePhonesInColumns
' 설  명  : 사용자가 지정한 범위의 전화번호를 표준 형식으로 정규화한다.
'           범위 지정 방법 (둘 다 지원):
'             - 마우스로 셀/열을 클릭하거나 드래그하여 선택
'             - 입력창에 범위 주소를 직접 입력 (예: B:B, C1:C100)
'           Ctrl 키를 누른 채 클릭하면 복수 영역 선택 가능
' ============================================================
Public Sub NormalizePhonesInColumns()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("전화번호 정리", _
        "엑셀 시트에서 전화번호 범위를 먼저 선택하세요. Ctrl+Shift+방향키, Shift+방향키, 마우스 드래그를 모두 사용할 수 있습니다.", _
        targetRange) Then Exit Sub

    Dim options As Object
    Set options = Phone_GetNormalizeOptionsFromUser()
    NormalizePhonesToCopiedColumns targetRange, options
End Sub

Public Sub NormalizePhonesWithForm()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("전화번호 정리", _
        "한국 전화번호 형식을 하이픈 포함 표준 형식으로 정리합니다. 정리할 셀/열/범위를 선택하세요.", _
        targetRange) Then Exit Sub

    Dim options As Object
    Set options = Phone_GetNormalizeOptionsFromUser()
    NormalizePhonesToCopiedColumns targetRange, options
End Sub

Public Sub RemovePhoneDashesToCopiedColumns()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("전화번호 대시 제거", _
        "대시를 제거할 전화번호 범위를 선택하세요. 원본 오른쪽에 복사본 열을 만들고 숫자만 남깁니다.", _
        targetRange) Then Exit Sub
    StripDigitsToCopiedColumns targetRange, "전화번호 대시 제거"
End Sub

Public Sub RemoveBizNoDashesToCopiedColumns()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("사업자등록번호 대시 제거", _
        "대시를 제거할 사업자등록번호 범위를 선택하세요. 원본 오른쪽에 복사본 열을 만들고 숫자만 남깁니다.", _
        targetRange) Then Exit Sub
    StripDigitsToCopiedColumns targetRange, "사업자등록번호 대시 제거"
End Sub

' ============================================================
' 프로시저 : ConvertTextNumbersInRange (내부 전용)
' 설  명  : 주어진 Range 내 각 셀을 순회하며
'           텍스트 형식 숫자를 실제 Double 숫자로 변환한다.
'           수식 셀은 건너뛰고 완료 후 통계를 표시한다.
' 파라미터 : target - 변환 대상 Range
' ============================================================
Private Sub ConvertTextNumbersInRange(ByVal target As Range)
    Dim cell As Range
    Dim convertedCount As Long   ' 변환 성공 셀 수
    Dim skippedFormulaCount As Long  ' 수식 셀이라 건너뛴 수
    Dim newValue As Variant      ' 변환된 숫자 값을 임시 저장
    Dim processedCells As Long
    Dim totalCells As Long
    Dim snapshot As Collection
    totalCells = target.Cells.CountLarge

    On Error GoTo ErrHandler
    Set snapshot = Toolkit_CreateRangeSnapshot(target)
    ' 화면 갱신 중지: 셀 업데이트마다 화면을 다시 그리지 않아 속도 향상
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "텍스트 숫자 변환", totalCells, "변환 준비 중"
    For Each cell In target.Cells
        If cell.HasFormula Then
            ' 수식 셀은 값을 덮어쓰면 수식이 사라지므로 건너뜀
            skippedFormulaCount = skippedFormulaCount + 1
        ElseIf TryNumberFromText(cell.Value, newValue) Then
            ' 텍스트를 숫자로 파싱 성공 시 셀 값 교체
            cell.Value = newValue
            convertedCount = convertedCount + 1
        End If
        processedCells = processedCells + 1
        Toolkit_RowProgressUpdate processedCells, totalCells, "텍스트 숫자 변환 중: " & CStr(processedCells) & "/" & CStr(totalCells) & "셀"
    Next cell

    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    MsgBox convertedCount & "개 셀이 변환되었습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다.", _
           vbInformation, "텍스트 숫자 변환"
    Exit Sub

ErrHandler:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    Toolkit_RestoreRangeSnapshot snapshot
    On Error GoTo 0
    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    If Toolkit_IsCancelError(errorNumber) Then
        MsgBox "작업을 취소해 변경 전 값으로 복구했습니다.", vbInformation, "텍스트 숫자 변환"
        Exit Sub
    End If
    MsgBox "오류가 발생해 변경 전 값으로 복구했습니다." & vbCrLf & errorDescription, _
           vbExclamation, "텍스트 숫자 변환"
End Sub

Public Sub DataCleanup_ClearResidualBlankValues()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("수식 빈값 정리", _
        "XLOOKUP, VLOOKUP 등의 결과를 값으로 붙여넣은 범위를 선택하세요. 빈 문자열만 실제 빈셀로 바꿉니다.", _
        targetRange) Then Exit Sub

    If Not Toolkit_ValidateWorksheet(targetRange.Worksheet, "수식 빈값 정리", True) Then Exit Sub
    If MsgBox("선택 범위에서 값 붙여넣기 후 남은 빈 문자열만 실제 빈셀로 변경합니다." & vbCrLf & vbCrLf & _
              "수식 셀, 숫자 0, 일반 텍스트, 공백 문자는 유지됩니다." & vbCrLf & _
              "이 작업은 선택한 셀을 제자리에서 수정합니다. 계속할까요?", _
              vbQuestion + vbYesNo, "수식 빈값 정리") <> vbYes Then Exit Sub

    Dim workRange As Range
    Set workRange = Intersect(targetRange, targetRange.Worksheet.UsedRange)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에 정리할 값이 없습니다.", vbInformation, "수식 빈값 정리"
        Exit Sub
    End If

    Dim candidates As Range
    On Error Resume Next
    Set candidates = workRange.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If candidates Is Nothing Then
        MsgBox "선택한 범위에 값으로 붙여넣은 빈 문자열이 없습니다.", vbInformation, "수식 빈값 정리"
        Exit Sub
    End If

    Dim totalCells As Long, processedCells As Long, clearedCount As Long
    Dim snapshot As Collection
    totalCells = candidates.Cells.CountLarge

    Dim cell As Range
    On Error GoTo ErrHandler
    Set snapshot = Toolkit_CreateRangeSnapshot(candidates)
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "수식 빈값 정리", totalCells, "빈 문자열 확인 준비 중"
    For Each cell In candidates.Cells
        If DataCleanup_IsResidualBlankValue(cell.Value2, cell.HasFormula) Then
            cell.ClearContents
            clearedCount = clearedCount + 1
        End If
        processedCells = processedCells + 1
        Toolkit_RowProgressUpdate processedCells, totalCells, _
            "수식 빈값 정리 중: " & CStr(processedCells) & "/" & CStr(totalCells) & "셀"
    Next cell

    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    MsgBox CStr(clearedCount) & "개 셀을 실제 빈셀로 변경했습니다.", vbInformation, "수식 빈값 정리"
    Exit Sub

ErrHandler:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    Toolkit_RestoreRangeSnapshot snapshot
    On Error GoTo 0
    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    If Toolkit_IsCancelError(errorNumber) Then
        MsgBox "작업을 취소해 변경 전 값으로 복구했습니다.", vbInformation, "수식 빈값 정리"
        Exit Sub
    End If
    MsgBox "수식 빈값 정리 중 오류가 발생해 변경 전 값으로 복구했습니다." & vbCrLf & errorDescription, vbExclamation, "수식 빈값 정리"
End Sub

Private Function DataCleanup_IsResidualBlankValue(ByVal cellValue As Variant, ByVal hasFormula As Boolean) As Boolean
    If hasFormula Then Exit Function
    DataCleanup_IsResidualBlankValue = (VarType(cellValue) = vbString And Len(CStr(cellValue)) = 0)
End Function

Public Function DataCleanup_IsResidualBlankForTest(ByVal cellValue As Variant, ByVal hasFormula As Boolean) As Boolean
    DataCleanup_IsResidualBlankForTest = DataCleanup_IsResidualBlankValue(cellValue, hasFormula)
End Function

Public Sub DataCleanup_ClearErrorAndZeroValues()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("오류값/0 지우기", _
        "#N/A, #VALUE!, #REF!, #DIV/0! 등의 Excel 오류값과 숫자 0을 지울 범위를 선택하세요.", _
        targetRange) Then Exit Sub

    If Not Toolkit_ValidateWorksheet(targetRange.Worksheet, "오류값/0 지우기", True) Then Exit Sub
    If MsgBox("선택 범위에서 Excel 오류값과 숫자 0이 있는 셀의 내용을 지웁니다." & vbCrLf & vbCrLf & _
              "수식 결과가 오류 또는 0인 셀은 수식도 함께 지워집니다." & vbCrLf & _
              "문자열 '0', 빈셀, 일반 텍스트, 0이 아닌 숫자는 유지됩니다." & vbCrLf & _
              "행과 열은 삭제하지 않습니다. 계속할까요?", _
              vbQuestion + vbYesNo, "오류값/0 지우기") <> vbYes Then Exit Sub

    Dim workRange As Range
    Set workRange = Intersect(targetRange, targetRange.Worksheet.UsedRange)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에 확인할 값이 없습니다.", vbInformation, "오류값/0 지우기"
        Exit Sub
    End If

    Dim snapshot As Collection
    Dim values As Variant
    Dim rowIndex As Long, columnIndex As Long
    Dim totalRows As Long, errorCount As Long, zeroCount As Long
    Dim valueKind As Long

    totalRows = workRange.Rows.Count
    values = workRange.Value2

    On Error GoTo ErrHandler
    Set snapshot = Toolkit_CreateRangeSnapshot(workRange)
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "오류값/0 지우기", totalRows, "정리할 셀 확인 준비 중"

    For rowIndex = 1 To workRange.Rows.Count
        For columnIndex = 1 To workRange.Columns.Count
            If workRange.Cells.CountLarge = 1 Then
                valueKind = DataCleanup_ClearValueKind(values)
            Else
                valueKind = DataCleanup_ClearValueKind(values(rowIndex, columnIndex))
            End If

            If valueKind > 0 Then
                workRange.Cells(rowIndex, columnIndex).ClearContents
                If valueKind = 1 Then
                    errorCount = errorCount + 1
                Else
                    zeroCount = zeroCount + 1
                End If
            End If
        Next columnIndex
        Toolkit_RowProgressUpdate rowIndex, totalRows, _
            "오류값/0 지우는 중: " & CStr(rowIndex) & "/" & CStr(totalRows) & "행"
    Next rowIndex

    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox "Excel 오류값 " & CStr(errorCount) & "개와 숫자 0 " & CStr(zeroCount) & _
           "개 셀의 내용을 지웠습니다.", vbInformation, "오류값/0 지우기"
    Exit Sub

ErrHandler:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    Toolkit_RestoreRangeSnapshot snapshot
    On Error GoTo 0
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_IsCancelError(errorNumber) Then
        MsgBox "작업을 취소해 변경 전 값으로 복구했습니다.", vbInformation, "오류값/0 지우기"
        Exit Sub
    End If
    MsgBox "오류값/0 정리 중 문제가 발생해 변경 전 값으로 복구했습니다." & vbCrLf & _
           errorDescription, vbExclamation, "오류값/0 지우기"
End Sub

Private Function DataCleanup_ClearValueKind(ByVal cellValue As Variant) As Long
    If IsError(cellValue) Then
        DataCleanup_ClearValueKind = 1
        Exit Function
    End If

    Select Case VarType(cellValue)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            If CDbl(cellValue) = 0 Then DataCleanup_ClearValueKind = 2
    End Select
End Function

Public Function DataCleanup_ClearValueKindForTest(ByVal cellValue As Variant) As Long
    DataCleanup_ClearValueKindForTest = DataCleanup_ClearValueKind(cellValue)
End Function

' ============================================================
' 프로시저 : NormalizePhonesToCopiedColumns (내부 전용)
' 설  명  : 선택 범위를 오른쪽 새 열에 복사한 뒤 복사본만 전화번호
'           표준 형식으로 정규화한다.
' ============================================================
Private Function Phone_GetNormalizeOptionsFromUser() As Object
    Dim options As Object
    Set options = Phone_DefaultNormalizeOptions()

    Dim optionsText As String
    optionsText = "prepend_zero" & vbTab & "앞자리가 0이 아니면 0을 붙인 후 정리 (우선 실행)" & vbLf & _
                  "repeat_review" & vbTab & "같은 숫자 4개 이상 반복 시 번호 대신 '확인 필요' 입력" & vbLf & _
                  "mobile_010" & vbTab & "대표번호 접두어가 아닌 8자리 값을 휴대폰 010-XXXX-XXXX 형식으로 정리"

    Dim selectedKeys As String
    selectedKeys = Toolkit_CheckListMenu( _
        "전화번호 정리 옵션", _
        "필요한 옵션을 선택하세요. 선택하지 않으면 기존 방식으로 복사본 열을 정리합니다.", _
        optionsText, "")

    selectedKeys = "," & selectedKeys & ","
    If InStr(1, selectedKeys, ",prepend_zero,", vbTextCompare) > 0 Then options("prependLeadingZero") = True
    If InStr(1, selectedKeys, ",repeat_review,", vbTextCompare) > 0 Then options("repeatRunAsReview") = True
    If InStr(1, selectedKeys, ",mobile_010,", vbTextCompare) > 0 Then options("eightDigitAsMobile010") = True

    Set Phone_GetNormalizeOptionsFromUser = options
End Function

Private Sub NormalizePhonesToCopiedColumns(ByVal target As Range, Optional ByVal normalizeOptions As Variant)
    Dim workRange As Range
    Dim options As Object
    If IsObject(normalizeOptions) Then
        Set options = normalizeOptions
    Else
        Set options = Phone_DefaultNormalizeOptions()
    End If

    Set workRange = Toolkit_TrimRangeToActualData(target)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 처리할 데이터를 찾지 못했습니다.", vbInformation, "전화번호 정리"
        Exit Sub
    End If

    Dim outputRange As Range
    Set outputRange = CopyRangeToAdjacentColumns(workRange, "전화번호_정리")
    If outputRange Is Nothing Then Exit Sub

    Dim splitCount As Long
    splitCount = Phone_CreateRepresentativeSplitColumns(workRange, outputRange)
    NormalizePhonesInRange outputRange, "전화번호 정리 복사본", options, splitCount
End Sub

Private Function Phone_CreateRepresentativeSplitColumns(ByVal sourceRange As Range, ByVal normalizedRange As Range) As Long
    Dim cell As Range
    Dim firstNumber As String
    Dim secondNumber As String
    Dim splitCount As Long

    For Each cell In sourceRange.Cells
        If Not cell.HasFormula Then
            If Phone_TrySplitRepresentative(CStr(cell.Value), firstNumber, secondNumber) Then splitCount = splitCount + 1
        End If
    Next cell
    If splitCount = 0 Then Exit Function

    Dim ws As Worksheet
    Dim secondOutputCol As Long
    Set ws = sourceRange.Worksheet
    secondOutputCol = normalizedRange.Column + normalizedRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, secondOutputCol, sourceRange.Columns.Count, "복수 대표번호 분할") Then Exit Function

    ws.Columns(secondOutputCol).Resize(, sourceRange.Columns.Count).Insert Shift:=xlToRight

    Dim secondRange As Range
    Set secondRange = ws.Range(ws.Cells(sourceRange.Row, secondOutputCol), _
                               ws.Cells(sourceRange.Row + sourceRange.Rows.Count - 1, secondOutputCol + sourceRange.Columns.Count - 1))
    secondRange.NumberFormat = "@"

    Dim rowIndex As Long
    Dim columnIndex As Long
    For columnIndex = 1 To sourceRange.Columns.Count
        If Len(Phone_DigitsForKoreanRules(CStr(sourceRange.Cells(1, columnIndex).Value))) < 7 Then
            normalizedRange.Cells(1, columnIndex).Value = CStr(sourceRange.Cells(1, columnIndex).Value) & "_정리1"
            secondRange.Cells(1, columnIndex).Value = CStr(sourceRange.Cells(1, columnIndex).Value) & "_정리2"
        End If

        For rowIndex = 1 To sourceRange.Rows.Count
            If Not sourceRange.Cells(rowIndex, columnIndex).HasFormula Then
                If Phone_TrySplitRepresentative(CStr(sourceRange.Cells(rowIndex, columnIndex).Value), firstNumber, secondNumber) Then
                    normalizedRange.Cells(rowIndex, columnIndex).NumberFormat = "@"
                    normalizedRange.Cells(rowIndex, columnIndex).Value = firstNumber
                    secondRange.Cells(rowIndex, columnIndex).Value = secondNumber
                End If
            End If
        Next rowIndex
    Next columnIndex

    Phone_CreateRepresentativeSplitColumns = splitCount
End Function

' ============================================================
' 프로시저 : NormalizePhonesInRange (내부 전용)
' 설  명  : 주어진 Range 내 각 셀을 순회하며
'           한국 전화번호를 표준 형식으로 정규화한다.
'           셀 서식을 텍스트(@)로 변경해 앞자리 0이 유지되도록 한다.
'           수식 셀은 건너뛰고 완료 후 통계를 표시한다.
' 파라미터 : target - 정규화 대상 Range
' ============================================================
Private Sub NormalizePhonesInRange(ByVal target As Range, Optional ByVal resultTitle As String = "전화번호 정규화", Optional ByVal normalizeOptions As Variant, Optional ByVal splitCount As Long = 0)
    Dim cell As Range
    Dim changedCount As Long         ' 실제 변경된 셀 수
    Dim skippedFormulaCount As Long  ' 수식 셀이라 건너뛴 수
    Dim normalized As String         ' 정규화된 전화번호 문자열
    Dim processedCells As Long
    Dim totalCells As Long
    Dim options As Object
    totalCells = target.Cells.CountLarge
    If IsObject(normalizeOptions) Then
        Set options = normalizeOptions
    Else
        Set options = Phone_DefaultNormalizeOptions()
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart resultTitle, totalCells, "전화번호 정리 준비 중"
    For Each cell In target.Cells
        If cell.HasFormula Then
            skippedFormulaCount = skippedFormulaCount + 1
        ElseIf NormalizeKoreanPhoneWithOptions(CStr(cell.Value), normalized, options) Then
            ' 값이 실제로 바뀐 경우에만 셀을 업데이트
            If CStr(cell.Value) <> normalized Then
                ' 텍스트 서식으로 설정해야 "010-..."처럼 하이픈이 포함된
                ' 번호가 날짜/수식으로 오해되지 않고 앞자리 0도 보존됨
                cell.NumberFormat = "@"
                cell.Value = normalized
                changedCount = changedCount + 1
            End If
        End If
        processedCells = processedCells + 1
        Toolkit_RowProgressUpdate processedCells, totalCells, "전화번호 정리 중: " & CStr(processedCells) & "/" & CStr(totalCells) & "셀"
    Next cell

    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    Dim splitMessage As String
    If splitCount > 0 Then splitMessage = splitCount & "개 복수 대표번호를 두 개 열로 분할했습니다." & vbCrLf
    MsgBox changedCount & "개 전화번호가 정규화되었습니다." & vbCrLf & _
           splitMessage & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다." & vbCrLf & _
           "원본 범위는 변경하지 않았습니다.", _
           vbInformation, resultTitle
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage(resultTitle) Then Exit Sub
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, resultTitle
End Sub

Public Function CopyRangeToAdjacentColumns(ByVal sourceRange As Range, ByVal headerSuffix As String) As Range
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = sourceRange.Worksheet

    If sourceRange.Areas.Count > 1 Then
        MsgBox "복사본 열 생성 기능은 한 번에 하나의 연속된 범위만 처리합니다." & vbCrLf & _
               "처리할 열 또는 표 범위를 다시 선택하세요.", vbExclamation, "복사본 열 생성"
        Exit Function
    End If

    Dim outputCol As Long
    outputCol = sourceRange.Column + sourceRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outputCol, sourceRange.Columns.Count, "복사본 열 생성") Then Exit Function

    Toolkit_BeginOperation "복사본 열을 생성하는 중입니다..."
    ws.Columns(outputCol).Resize(, sourceRange.Columns.Count).Insert Shift:=xlToRight

    Dim outputRange As Range
    Set outputRange = ws.Range(ws.Cells(sourceRange.Row, outputCol), _
                               ws.Cells(sourceRange.Row + sourceRange.Rows.Count - 1, outputCol + sourceRange.Columns.Count - 1))
    sourceRange.Copy Destination:=outputRange
    Application.CutCopyMode = False

    Dim c As Long
    For c = 1 To outputRange.Columns.Count
        outputRange.Columns(c).NumberFormat = "@"
    Next c

    Toolkit_EndOperation
    Set CopyRangeToAdjacentColumns = outputRange
    Exit Function

ErrHandler:
    Toolkit_EndOperation
    Application.CutCopyMode = False
    MsgBox "복사본 열을 생성하는 중 오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, "복사본 열 생성"
End Function

Private Sub StripDigitsToCopiedColumns(ByVal target As Range, ByVal resultTitle As String)
    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(target)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 처리할 데이터를 찾지 못했습니다.", vbInformation, resultTitle
        Exit Sub
    End If

    Dim outputRange As Range
    Set outputRange = CopyRangeToAdjacentColumns(workRange, "대시제거")
    If outputRange Is Nothing Then Exit Sub

    Dim cell As Range
    Dim changedCount As Long
    Dim skippedFormulaCount As Long
    Dim digits As String
    Dim processedCells As Long
    Dim totalCells As Long
    totalCells = outputRange.Cells.CountLarge

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart resultTitle, totalCells, "구분기호 제거 준비 중"
    For Each cell In outputRange.Cells
        If cell.HasFormula Then
            skippedFormulaCount = skippedFormulaCount + 1
        Else
            If InStr(1, resultTitle, "전화번호", vbTextCompare) > 0 Then
                digits = Phone_DigitsForKoreanRules(CStr(cell.Value))
            Else
                digits = DigitsOnly(CStr(cell.Value))
            End If
            If Len(digits) > 0 And CStr(cell.Value) <> digits Then
                cell.NumberFormat = "@"
                cell.Value = digits
                changedCount = changedCount + 1
            End If
        End If
        processedCells = processedCells + 1
        Toolkit_RowProgressUpdate processedCells, totalCells, "구분기호 제거 중: " & CStr(processedCells) & "/" & CStr(totalCells) & "셀"
    Next cell

    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    MsgBox changedCount & "개 셀에서 대시/구분기호를 제거했습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다." & vbCrLf & _
           "원본 범위는 변경하지 않았습니다.", vbInformation, resultTitle
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage(resultTitle) Then Exit Sub
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, resultTitle
End Sub

' ============================================================
' 함  수 : TryNumberFromText (내부 전용)
' 설  명  : 셀 값이 텍스트 형식의 숫자이면 Double로 변환하여
'           valueOut에 저장하고 True를 반환한다.
'           처리하는 형식:
'             - 앞의 작은따옴표(') 제거  → '1234  =>  1234
'             - 비표시 공백(NBSP, U+00A0) 제거
'             - 천 단위 쉼표 제거          → 1,234  =>  1234
'             - 원화 기호(\, U+20A9) 제거  → \1000  =>  1000
'             - 괄호 음수 처리             → (500)  =>  -500
'             - 퍼센트 처리               → 75%    =>  0.75
' 파라미터 : valueIn  - 원본 셀 값 (Variant)
'            valueOut - 변환된 숫자 (ByRef, 호출자에게 반환)
' 반환값  : 변환 성공 시 True, 실패 시 False
' ============================================================
Private Function TryNumberFromText(ByVal valueIn As Variant, ByRef valueOut As Variant) As Boolean
    Dim textValue As String
    Dim negativeValue As Boolean  ' 괄호로 감싸진 음수 여부
    Dim percentValue As Boolean   ' 퍼센트(%) 값 여부

    ' 오류 값이거나 빈 셀이면 변환 불필요
    If IsError(valueIn) Or IsEmpty(valueIn) Then Exit Function
    ' 이미 숫자 타입(String이 아닌)이면 변환 불필요
    If IsNumeric(valueIn) And VarType(valueIn) <> vbString Then Exit Function

    textValue = Trim$(CStr(valueIn))
    If Len(textValue) = 0 Then Exit Function

    ' 엑셀이 숫자 앞에 자동으로 붙이는 작은따옴표(') 제거
    If Left$(textValue, 1) = "'" Then textValue = Mid$(textValue, 2)

    ' 비표시 공백(Non-Breaking Space, U+00A0) 및 일반 공백 제거
    textValue = Replace(textValue, ChrW(160), "")
    textValue = Replace(textValue, " ", "")
    ' 천 단위 구분 쉼표 제거
    textValue = Replace(textValue, ",", "")
    ' 원화 기호 두 종류 모두 제거 (역슬래시형 \, 유니코드 U+20A9)
    textValue = Replace(textValue, "\", "")
    textValue = Replace(textValue, ChrW$(8361), "")

    ' "(500)" 형태의 괄호 음수 처리
    If Left$(textValue, 1) = "(" And Right$(textValue, 1) = ")" Then
        negativeValue = True
        textValue = Mid$(textValue, 2, Len(textValue) - 2)  ' 괄호 제거
    End If

    ' "75%" 형태의 퍼센트 처리
    If Right$(textValue, 1) = "%" Then
        percentValue = True
        textValue = Left$(textValue, Len(textValue) - 1)  ' % 기호 제거
    End If

    ' 정리 후에도 숫자가 아니면 변환 실패
    If Len(textValue) = 0 Or Not IsNumeric(textValue) Then Exit Function

    valueOut = CDbl(textValue)
    If negativeValue Then valueOut = -CDbl(valueOut)      ' 음수 부호 적용
    If percentValue Then valueOut = CDbl(valueOut) / 100  ' 퍼센트 → 소수 변환
    TryNumberFromText = True
End Function

' ============================================================
' 함  수 : NormalizeKoreanPhone (내부 전용)
' 설  명  : 입력 문자열에서 숫자만 추출한 뒤
'           한국 전화번호 규칙에 따라 표준 형식(XXX-XXXX-XXXX 또는 XXXX-XXXX-XXXX)으로
'           정규화하여 valueOut에 저장하고 True를 반환한다.
'           지원 형식:
'             - 서울 지역번호 (02): 02-XXX-XXXX / 02-XXXX-XXXX
'             - 휴대전화 (01X): 01X-XXX-XXXX / 01X-XXXX-XXXX
'             - 기타 지역번호 (0XX): 0XX-XXX-XXXX / 0XX-XXXX-XXXX
'             - 안심번호 (050X): 050X-XXX-XXXX / 050X-XXXX-XXXX
'             - 국제번호 +82 또는 0082 → 0으로 치환 후 처리
' 파라미터 : valueIn  - 원본 전화번호 문자열
'            valueOut - 정규화된 전화번호 (ByRef, 호출자에게 반환)
' 반환값  : 정규화 성공 시 True, 인식 불가 형식이면 False
' ============================================================
Private Function NormalizeKoreanPhone(ByVal valueIn As String, ByRef valueOut As String) As Boolean
    Dim options As Object
    Set options = Phone_DefaultNormalizeOptions()
    NormalizeKoreanPhone = NormalizeKoreanPhoneWithOptions(valueIn, valueOut, options)
End Function

Private Function Phone_DefaultNormalizeOptions() As Object
    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    options("prependLeadingZero") = False
    options("repeatRunAsReview") = False
    options("eightDigitAsMobile010") = False
    Set Phone_DefaultNormalizeOptions = options
End Function

Private Function Phone_NormalizeOptionEnabled(ByVal options As Object, ByVal key As String) As Boolean
    On Error GoTo SafeExit
    If Not options Is Nothing Then
        If options.Exists(key) Then Phone_NormalizeOptionEnabled = CBool(options(key))
    End If
SafeExit:
End Function

Private Function Phone_ApplyLeadingZeroOption(ByVal digits As String, ByVal options As Object) As String
    Phone_ApplyLeadingZeroOption = digits
    If Not Phone_NormalizeOptionEnabled(options, "prependLeadingZero") Then Exit Function
    If Len(digits) = 0 Then Exit Function
    If Left$(digits, 1) = "0" Then Exit Function
    If Left$(digits, 4) = "0082" Or Left$(digits, 2) = "82" Then Exit Function
    If Phone_IsKnownRepresentativePrefix(digits) Then Exit Function
    Phone_ApplyLeadingZeroOption = "0" & digits
End Function

Private Function Phone_ShouldWriteReviewInstead(ByVal digits As String, ByVal options As Object) As Boolean
    If Not Phone_NormalizeOptionEnabled(options, "repeatRunAsReview") Then Exit Function
    Phone_ShouldWriteReviewInstead = Phone_HasRepeatedDigitRun(digits, 4)
End Function

Private Function Phone_IsKnownRepresentativePrefix(ByVal digits As String) As Boolean
    If Len(digits) < 4 Then Exit Function

    Select Case Left$(digits, 4)
        Case "1588", "1577", "1899", _
             "1544", "1644", "1661", _
             "1566", "1600", "1670", _
             "1688", "1666", "1599", _
             "1877", "1855", "1800"
            Phone_IsKnownRepresentativePrefix = True
    End Select
End Function

Private Function Phone_TrySplitRepresentative(ByVal valueIn As String, ByRef firstNumber As String, ByRef secondNumber As String) As Boolean
    firstNumber = ""
    secondNumber = ""

    Dim digits As String
    digits = Phone_DigitsForKoreanRules(valueIn)
    If Len(digits) <> 9 Then Exit Function
    If Not Phone_IsKnownRepresentativePrefix(digits) Then Exit Function
    If Mid$(digits, 8, 1) = Mid$(digits, 9, 1) Then Exit Function

    Dim sharedDigits As String
    sharedDigits = Mid$(digits, 5, 3)
    firstNumber = Left$(digits, 4) & "-" & sharedDigits & Mid$(digits, 8, 1)
    secondNumber = Left$(digits, 4) & "-" & sharedDigits & Mid$(digits, 9, 1)
    Phone_TrySplitRepresentative = True
End Function

Private Function Phone_SubscriberDigitsForSequenceCheck(ByVal digits As String) As String
    If Len(digits) = 0 Then Exit Function

    If Phone_IsSafeNumberDigits(digits) And (Len(digits) = 11 Or Len(digits) = 12) Then
        Phone_SubscriberDigitsForSequenceCheck = Mid$(digits, 5)
    ElseIf Left$(digits, 2) = "02" And (Len(digits) = 9 Or Len(digits) = 10) Then
        Phone_SubscriberDigitsForSequenceCheck = Mid$(digits, 3)
    ElseIf Left$(digits, 3) Like "01#" And (Len(digits) = 10 Or Len(digits) = 11) Then
        Phone_SubscriberDigitsForSequenceCheck = Mid$(digits, 4)
    ElseIf Left$(digits, 3) Like "0##" And (Len(digits) = 10 Or Len(digits) = 11) Then
        Phone_SubscriberDigitsForSequenceCheck = Mid$(digits, 4)
    Else
        Phone_SubscriberDigitsForSequenceCheck = digits
    End If
End Function

Private Function Phone_IsSequentialRun(ByVal digits As String) As Boolean
    If Len(digits) < 7 Then Exit Function

    Dim i As Long
    Dim prevDigit As Long
    Dim currentDigit As Long
    Dim ascending As Boolean
    Dim descending As Boolean

    ascending = True
    descending = True
    prevDigit = CLng(Left$(digits, 1))
    For i = 2 To Len(digits)
        currentDigit = CLng(Mid$(digits, i, 1))
        If currentDigit <> ((prevDigit + 1) Mod 10) Then ascending = False
        If currentDigit <> ((prevDigit + 9) Mod 10) Then descending = False
        prevDigit = currentDigit
    Next i

    Phone_IsSequentialRun = (ascending Or descending)
End Function

Private Function Phone_IsSequentialPlaceholderDigits(ByVal digits As String) As Boolean
    Phone_IsSequentialPlaceholderDigits = Phone_IsSequentialRun(Phone_SubscriberDigitsForSequenceCheck(digits))
End Function

Private Function Phone_BlockHasDistinctDigits(ByVal blockText As String) As Boolean
    If Len(blockText) < 2 Then Exit Function

    Dim i As Long
    For i = 2 To Len(blockText)
        If Mid$(blockText, i, 1) <> Left$(blockText, 1) Then
            Phone_BlockHasDistinctDigits = True
            Exit Function
        End If
    Next i
End Function

Private Function Phone_HasRepeatedBlockPattern(ByVal digits As String) As Boolean
    Dim subscriberDigits As String
    subscriberDigits = Phone_SubscriberDigitsForSequenceCheck(digits)
    If Len(subscriberDigits) < 4 Then Exit Function

    Dim blockLength As Long
    Dim startPos As Long
    Dim blockText As String
    For blockLength = 2 To Len(subscriberDigits) \ 2
        For startPos = 1 To Len(subscriberDigits) - (blockLength * 2) + 1
            blockText = Mid$(subscriberDigits, startPos, blockLength)
            If Phone_BlockHasDistinctDigits(blockText) Then
                If blockText = Mid$(subscriberDigits, startPos + blockLength, blockLength) Then
                    Phone_HasRepeatedBlockPattern = True
                    Exit Function
                End If
            End If
        Next startPos
    Next blockLength
End Function

Private Function NormalizeKoreanPhoneWithOptions(ByVal valueIn As String, ByRef valueOut As String, ByVal options As Object) As Boolean
    Dim source As String
    Dim digits As String  ' 숫자만 추출한 문자열
    Dim rawDigits As String

    source = Trim$(valueIn)
    If Len(source) = 0 Then Exit Function

    ' 괄호, 공백, 하이픈 등 비숫자 문자를 모두 제거하여 순수 숫자열 획득
    ' 예: 02)123-4567, (031)1234-5678, +82-10-1234-5678
    source = Replace(source, "(", "")
    source = Replace(source, ")", "")
    rawDigits = DigitsOnly(source)
    If Len(rawDigits) = 0 Then Exit Function
    rawDigits = Phone_ApplyLeadingZeroOption(rawDigits, options)
    digits = Phone_DigitsForKoreanRules(rawDigits)
    If Len(digits) = 0 Then Exit Function

    If Phone_IsSequentialPlaceholderDigits(digits) Then
        valueOut = ""
    ElseIf Phone_ShouldWriteReviewInstead(digits, options) Then
        valueOut = "확인 필요"
    ElseIf Phone_IsKnownRepresentativePrefix(digits) And Len(digits) = 8 Then
        valueOut = Left$(digits, 4) & "-" & Right$(digits, 4)
    ElseIf Phone_IsKnownRepresentativePrefix(digits) And Len(digits) = 9 Then
        Dim secondRepresentative As String
        If Not Phone_TrySplitRepresentative(valueIn, valueOut, secondRepresentative) Then Exit Function
    ElseIf Phone_IsSafeNumberDigits(digits) Then
        ' ── 안심번호 (0507 등 050X 계열) ─────────────────────
        ' 11자리: 050X-XXX-XXXX
        ' 12자리: 050X-XXXX-XXXX
        Select Case Len(digits)
            Case 11
                valueOut = Left$(digits, 4) & "-" & Mid$(digits, 5, 3) & "-" & Right$(digits, 4)
            Case 12
                valueOut = Left$(digits, 4) & "-" & Mid$(digits, 5, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function
        End Select
    ElseIf Len(digits) = 8 And Phone_NormalizeOptionEnabled(options, "eightDigitAsMobile010") Then
        ' ── 휴대폰 본문 8자리만 있는 경우 사용자가 선택하면 010을 붙여 처리
        valueOut = "010-" & Left$(digits, 4) & "-" & Right$(digits, 4)
    ElseIf Len(digits) = 8 Then
        ' ── 국번 없는 대표번호 (1588-1222 등) ────────────────
        valueOut = Left$(digits, 4) & "-" & Right$(digits, 4)
    ElseIf Left$(digits, 2) = "02" Then
        ' ── 서울 지역번호 (02) ──────────────────────────────
        ' 02 포함 9자리: 02-XXX-XXXX
        ' 02 포함 10자리: 02-XXXX-XXXX
        Select Case Len(digits)
            Case 9
                valueOut = "02-" & Mid$(digits, 3, 3) & "-" & Right$(digits, 4)
            Case 10
                valueOut = "02-" & Mid$(digits, 3, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function  ' 자릿수가 맞지 않으면 정규화 불가
        End Select
    ElseIf Left$(digits, 3) Like "01#" Then
        ' ── 휴대전화 (010, 011, 016, 017, 018, 019) ─────────
        ' 10자리: 01X-XXX-XXXX
        ' 11자리: 01X-XXXX-XXXX
        Select Case Len(digits)
            Case 10
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 3) & "-" & Right$(digits, 4)
            Case 11
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function
        End Select
    ElseIf Left$(digits, 2) = "10" Then
        ' 엑셀이 숫자로 읽으며 휴대전화의 선행 0을 제거한 경우 복원
        digits = "0" & digits
        Select Case Len(digits)
            Case 11
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function
        End Select
    ElseIf Left$(digits, 3) Like "0##" Then
        ' ── 기타 지역번호 (031, 032, 051, 052 … 등) ─────────
        ' 10자리: 0XX-XXX-XXXX
        ' 11자리: 0XX-XXXX-XXXX
        Select Case Len(digits)
            Case 10
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 3) & "-" & Right$(digits, 4)
            Case 11
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function
        End Select
    Else
        ' 위 패턴에 해당하지 않으면 인식 불가 → False 반환
        Exit Function
    End If

    NormalizeKoreanPhoneWithOptions = True
End Function

' ============================================================
' 함  수 : DigitsOnly (내부 전용)
' 설  명  : 입력 문자열에서 0~9 사이의 문자만 추출하여 반환한다.
'           전화번호의 하이픈, 괄호, 공백 등을 제거할 때 사용한다.
' 파라미터 : source - 원본 문자열
' 반환값  : 숫자 문자만으로 이루어진 문자열
' ============================================================
Public Function DigitsOnly(ByVal source As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(source)
        ch = Mid$(source, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

Public Function Phone_DigitsForKoreanRules(ByVal source As String) As String
    Dim digits As String
    digits = DigitsOnly(source)
    If Len(digits) = 0 Then Exit Function

    ' 국제전화 접두어(0082 또는 82)를 국내 형식(0)으로 변환
    If Left$(digits, 4) = "0082" Then digits = "0" & Mid$(digits, 5)
    If Left$(digits, 2) = "82" Then digits = "0" & Mid$(digits, 3)

    ' Excel이 0507 안심번호의 선행 0을 제거한 경우 복원
    If Left$(digits, 3) = "507" And (Len(digits) = 10 Or Len(digits) = 11) Then digits = "0" & digits

    Phone_DigitsForKoreanRules = digits
End Function

Private Function Phone_IsSafeNumberDigits(ByVal digits As String) As Boolean
    If Len(digits) < 4 Then Exit Function
    Phone_IsSafeNumberDigits = (Left$(digits, 4) Like "050#")
End Function

Private Function Phone_IsPlaceholderZeroPattern(ByVal digits As String) As Boolean
    If Len(digits) < 7 Then Exit Function

    Dim i As Long
    For i = 1 To Len(digits)
        If Mid$(digits, i, 1) <> "0" Then Exit Function
    Next i

    Phone_IsPlaceholderZeroPattern = True
End Function

Private Function Phone_HasRepeatedDigitRun(ByVal digits As String, ByVal minRunLength As Long) As Boolean
    If minRunLength < 2 Then minRunLength = 2

    Dim i As Long
    Dim previousDigit As String
    Dim runLength As Long
    For i = 1 To Len(digits)
        Dim currentDigit As String
        currentDigit = Mid$(digits, i, 1)
        If currentDigit = previousDigit Then
            runLength = runLength + 1
        Else
            previousDigit = currentDigit
            runLength = 1
        End If
        If runLength >= minRunLength Then
            Phone_HasRepeatedDigitRun = True
            Exit Function
        End If
    Next i
End Function

Public Sub Phone_ValidateSelection()
    Dim phoneRange As Range
    If Not Toolkit_GetRangeFromUserForm("전화번호 유효성 검사", _
        "전화번호 범위를 선택하세요. 오른쪽에 전화번호 유형과 유효성 결과를 생성합니다.", _
        phoneRange) Then Exit Sub
    Set phoneRange = phoneRange.Areas(1)
    If phoneRange.Columns.Count > 1 Then Set phoneRange = phoneRange.Columns(1)

    Dim ws As Worksheet: Set ws = phoneRange.Worksheet
    Dim firstRow As Long: firstRow = phoneRange.Row
    Dim lastRow As Long: lastRow = phoneRange.Row + phoneRange.Rows.Count - 1
    Dim dataStart As Long: dataStart = firstRow
    If Len(DigitsOnly(CStr(ws.Cells(firstRow, phoneRange.Column).Value))) = 0 Then dataStart = firstRow + 1

    Dim outCol As Long: outCol = phoneRange.Column + phoneRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outCol, 2, "전화번호 유효성 검사") Then Exit Sub

    Dim totalRows As Long
    totalRows = Application.Max(0, lastRow - dataStart + 1)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outCol).Resize(, 2).Insert Shift:=xlToRight
    Roster_WriteArray ws, firstRow, outCol, Array("전화번호 유형", "전화번호 유효성")

    Dim r As Long, result As Variant
    Toolkit_RowProgressStart "전화번호 유효성 검사", totalRows, "전화번호 검사 준비 중"
    For r = dataStart To lastRow
        result = Phone_ClassifyOne(CStr(ws.Cells(r, phoneRange.Column).Value))
        Roster_WriteArray ws, r, outCol, result
        Toolkit_RowProgressUpdate r - dataStart + 1, totalRows, "전화번호 검사 중: " & CStr(r - dataStart + 1) & "/" & CStr(totalRows) & "행"
    Next r
    ws.Columns(outCol).Resize(, 2).AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox "전화번호 유효성 검사를 완료했습니다.", vbInformation, "전화번호 유효성 검사"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("전화번호 유효성 검사") Then Exit Sub
    MsgBox "전화번호 유효성 검사 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "전화번호 유효성 검사"
End Sub

Public Function Phone_ClassifyOne(ByVal phoneText As String) As Variant
    Dim digits As String: digits = Phone_DigitsForKoreanRules(phoneText)
    Dim phoneType As String
    Dim statusText As String
    Dim firstNumber As String
    Dim secondNumber As String

    If Len(digits) = 0 Then
        phoneType = "빈값"
        statusText = "확인 필요"
    ElseIf Phone_IsPlaceholderZeroPattern(digits) Then
        phoneType = "오류 패턴"
        statusText = "오류"
    ElseIf Phone_IsSequentialPlaceholderDigits(digits) Then
        phoneType = "오류 패턴"
        statusText = "오류"
    ElseIf Phone_HasRepeatedDigitRun(digits, 6) Then
        phoneType = "오류 패턴"
        statusText = "오류"
    ElseIf Phone_HasRepeatedBlockPattern(digits) Then
        phoneType = "오류 패턴"
        statusText = "오류"
    ElseIf Phone_IsSafeNumberDigits(digits) Then
        phoneType = "안심번호"
        If Len(digits) = 11 Or Len(digits) = 12 Then statusText = "정상" Else statusText = "확인 필요"
    ElseIf Left$(digits, 3) = "010" And Len(digits) = 11 Then
        phoneType = "휴대폰"
        statusText = "정상"
    ElseIf Left$(digits, 2) = "02" And (Len(digits) = 9 Or Len(digits) = 10) Then
        phoneType = "서울 지역번호"
        statusText = "정상"
    ElseIf Phone_IsKnownRepresentativePrefix(digits) Then
        phoneType = "대표번호"
        If Len(digits) = 8 Then
            statusText = "정상"
        ElseIf Len(digits) = 9 And Phone_TrySplitRepresentative(phoneText, firstNumber, secondNumber) Then
            statusText = "대표번호 복수 확인 필요"
        Else
            statusText = "확인 필요"
        End If
    ElseIf Left$(digits, 2) = "15" Or Left$(digits, 2) = "16" Or Left$(digits, 2) = "18" Then
        phoneType = "대표번호"
        If Len(digits) = 8 Then statusText = "정상" Else statusText = "확인 필요"
    ElseIf Len(digits) = 8 Then
        phoneType = "대표번호"
        statusText = "정상"
    ElseIf Len(digits) = 10 Or Len(digits) = 11 Then
        phoneType = "지역/일반번호"
        statusText = "정상"
    Else
        phoneType = "알 수 없음"
        statusText = "확인 필요"
    End If

    If statusText = "정상" And Phone_HasRepeatedDigitRun(digits, 5) Then statusText = "확인 필요"
    Phone_ClassifyOne = Array(phoneType, statusText)
End Function

Public Function Phone_NormalizeForTest(ByVal phoneText As String) As String
    Dim normalized As String
    If NormalizeKoreanPhone(phoneText, normalized) Then Phone_NormalizeForTest = normalized
End Function

Public Function Phone_NormalizeResultForTest(ByVal phoneText As String) As String
    Dim normalized As String
    If NormalizeKoreanPhone(phoneText, normalized) Then
        Phone_NormalizeResultForTest = "TRUE|" & normalized
    Else
        Phone_NormalizeResultForTest = "FALSE|"
    End If
End Function

Public Function Phone_NormalizeForTestWithOptions(ByVal phoneText As String, _
                                                  ByVal prependLeadingZero As Boolean, _
                                                  ByVal repeatRunAsReview As Boolean, _
                                                  ByVal eightDigitAsMobile010 As Boolean) As String
    Dim options As Object
    Set options = Phone_DefaultNormalizeOptions()
    options("prependLeadingZero") = prependLeadingZero
    options("repeatRunAsReview") = repeatRunAsReview
    options("eightDigitAsMobile010") = eightDigitAsMobile010

    Dim normalized As String
    If NormalizeKoreanPhoneWithOptions(phoneText, normalized, options) Then Phone_NormalizeForTestWithOptions = normalized
End Function

Public Function Phone_ClassifyForTest(ByVal phoneText As String) As String
    Dim result As Variant
    result = Phone_ClassifyOne(phoneText)
    Phone_ClassifyForTest = CStr(result(0)) & "|" & CStr(result(1))
End Function

Public Function Phone_SplitRepresentativeForTest(ByVal phoneText As String) As String
    Dim firstNumber As String
    Dim secondNumber As String
    If Phone_TrySplitRepresentative(phoneText, firstNumber, secondNumber) Then
        Phone_SplitRepresentativeForTest = "TRUE|" & firstNumber & "|" & secondNumber
    Else
        Phone_SplitRepresentativeForTest = "FALSE||"
    End If
End Function
