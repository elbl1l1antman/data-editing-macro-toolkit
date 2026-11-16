Attribute VB_Name = "OutlierTools"
Option Explicit


' ============================================================
' ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'  이상치 탐지 및 대체 (OutlierDetector)
'  · OutlierDetect_Setup : 설정 시트 생성/갱신
'  · OutlierDetect_Run   : 이상치 탐지 및 결과 시트 출력
' ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ============================================================

Private Const OD_CFG_SHEET As String = "_이상치탐지_설정"
Private Const OD_WORK_SHEET As String = "이상치탐지_작업"
Private Const OD_MIN_N     As Long   = 30

Private Const OD_IQR15  As String = "IQR × 1.5"
Private Const OD_IQR20  As String = "IQR × 2.0"
Private Const OD_P95    As String = "상위 5% 제외"
Private Const OD_P99    As String = "상위 1% 제외"

Private Const OD_CAP    As String = "상한값 대체"
Private Const OD_MEDIAN As String = "중앙값 대체"
Private Const OD_MEAN   As String = "평균값 대체"
Private Const OD_REVIEW As String = "직접 검토(복사만)"

' 설정 시트 셀 주소
Private Const OD_C_SRC  As String = "B3"
Private Const OD_C_VAR  As String = "B4"
Private Const OD_C_CRIT As String = "B5"
Private Const OD_C_METH As String = "B6"
Private Const OD_C_MEMO As String = "B8"

' ============================================================
' 프로시저 : OutlierDetect_Setup
' 설    명  : 설정 시트를 생성하거나 갱신한다.
'             소스 시트를 변경했을 때 다시 실행하면
'             분석 변수 드롭다운이 최신 컬럼 목록으로 갱신된다.
' ============================================================
Public Sub OutlierDetect_Setup()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "이상치 탐지 작업 시트를 만들 통합문서를 먼저 여세요.", vbExclamation, "이상치 탐지"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = OD_CreateWorkspaceSheet(wb)
    ws.Activate

    MsgBox "이상치 탐지 작업 시트를 만들었습니다." & vbCrLf & vbCrLf & _
           "원본 데이터에서 검토 대상만 필터링해 이 시트에 붙여넣은 뒤 " & _
           "[데이터 툴킷] > [이상치 탐지] > [분석 열 선택]을 실행하세요.", _
           vbInformation, "이상치 탐지"
    Exit Sub

ErrHandler:
    MsgBox "이상치 탐지 작업 시트를 만드는 중 오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, "이상치 탐지"
End Sub

' ============================================================
' 프로시저 : OutlierDetect_Run
' 설    명  : 설정 시트 값을 읽어 이상치 탐지 및 대체를 수행하고
'             결과를 새 시트에 출력한다.
'             원본 시트는 수정하지 않는다.
' ============================================================
Public Sub OutlierDetect_Run()
    OutlierDetect_ApplyToSelection
End Sub

Private Sub OutlierDetect_RunLegacy()
    ' ── 실행 입력값 수집 ─────────────────────────────────
    Dim srcWs As Worksheet
    Dim dataArea As Range
    Dim analysisCol As Long
    Dim varName As String
    Dim criteria As String
    Dim method As String
    Dim memo As String
    Dim srcName As String

    If Not OD_CollectRunInputs(srcWs, dataArea, analysisCol, varName, criteria, method, memo) Then Exit Sub
    srcName = srcWs.Name

    ' ── 보이는 데이터 행 수집 ────────────────────────────
    Dim firstData As Long: firstData = dataArea.Row + 1
    Dim lastRow   As Long: lastRow   = dataArea.Row + dataArea.Rows.Count - 1
    Dim rowCount  As Long: rowCount  = 0
    Dim r As Long

    For r = firstData To lastRow
        If Not srcWs.Rows(r).Hidden Then rowCount = rowCount + 1
    Next r

    If rowCount = 0 Then
        MsgBox "필터 적용 후 보이는 데이터 행이 없습니다.", vbExclamation, "이상치 탐지"
        Exit Sub
    End If

    ' 실행 전 확인 대화상자
    If MsgBox("분석 대상 행: " & rowCount & "개" & vbCrLf & _
              "분석 변수: " & varName & vbCrLf & _
              "이상치 기준: " & criteria & vbCrLf & _
              "대체 방식: " & method & vbCrLf & vbCrLf & "계속하시겠습니까?", _
              vbQuestion + vbYesNo, "이상치 탐지 확인") = vbNo Then Exit Sub

    ' 행 인덱스 배열
    Dim dataRows() As Long: ReDim dataRows(1 To rowCount)
    Dim idx As Long: idx = 0
    For r = firstData To lastRow
        If Not srcWs.Rows(r).Hidden Then
            idx = idx + 1
            dataRows(idx) = r
        End If
    Next r

    ' ── 수치 값 추출 ──────────────────────────────────────
    Dim values()   As Double:  ReDim values(1 To rowCount)
    Dim isNumRow() As Boolean: ReDim isNumRow(1 To rowCount)
    Dim numCount   As Long: numCount = 0

    For idx = 1 To rowCount
        Dim cv As Variant: cv = srcWs.Cells(dataRows(idx), analysisCol).Value
        If Not IsError(cv) And Not IsEmpty(cv) And IsNumeric(cv) Then
            numCount = numCount + 1
            values(numCount) = CDbl(cv)
            isNumRow(idx) = True
        End If
    Next idx

    If numCount = 0 Then
        MsgBox "'" & varName & "' 컬럼에 유효한 숫자값이 없습니다.", _
               vbExclamation, "이상치 탐지"
        Exit Sub
    End If

    ReDim Preserve values(1 To numCount)

    ' 사례수 부족 경고
    If numCount < OD_MIN_N Then
        If MsgBox("유효 사례수(" & numCount & ")가 " & OD_MIN_N & "개 미만입니다." & vbCrLf & _
                  "통계량이 불안정할 수 있습니다. 계속하시겠습니까?", _
                  vbWarning + vbYesNo, "사례수 경고") = vbNo Then Exit Sub
    End If

    ' ── 통계량 계산 ──────────────────────────────────────
    Dim sortedValues() As Double
    OD_CopyAndSortValues values, numCount, sortedValues

    Dim q1     As Double: q1     = OD_PercentileIncSorted(sortedValues, numCount, 0.25)
    Dim q3     As Double: q3     = OD_PercentileIncSorted(sortedValues, numCount, 0.75)
    Dim iqrVal As Double: iqrVal = q3 - q1
    Dim medV   As Double: medV   = OD_PercentileIncSorted(sortedValues, numCount, 0.5)
    Dim meanV  As Double: meanV  = OD_AverageValues(values, numCount)
    Dim minV   As Double: minV   = sortedValues(1)
    Dim maxV   As Double: maxV   = sortedValues(numCount)

    Dim lo As Double, hi As Double
    Select Case criteria
        Case OD_IQR15: lo = q1 - 1.5 * iqrVal:                    hi = q3 + 1.5 * iqrVal
        Case OD_IQR20: lo = q1 - 2.0 * iqrVal:                    hi = q3 + 2.0 * iqrVal
        Case OD_P95:   lo = OD_PercentileIncSorted(sortedValues, numCount, 0.05): hi = OD_PercentileIncSorted(sortedValues, numCount, 0.95)
        Case OD_P99:   lo = OD_PercentileIncSorted(sortedValues, numCount, 0.01): hi = OD_PercentileIncSorted(sortedValues, numCount, 0.99)
    End Select

    ' 중앙값/평균값 대체 기준: 이상치 제외 후 계산
    Dim replBase As Double
    If method = OD_MEDIAN Or method = OD_MEAN Then
        Dim trimVals() As Double: ReDim trimVals(1 To numCount)
        Dim trimN As Long: trimN = 0
        Dim vi As Long
        For vi = 1 To numCount
            If values(vi) >= lo And values(vi) <= hi Then
                trimN = trimN + 1
                trimVals(trimN) = values(vi)
            End If
        Next vi
        If trimN > 0 Then
            ReDim Preserve trimVals(1 To trimN)
            If method = OD_MEDIAN Then replBase = OD_MedianValues(trimVals, trimN)
            If method = OD_MEAN Then replBase = OD_AverageValues(trimVals, trimN)
        Else
            ' 이상치 제외 후 데이터가 없는 극단적 경우 전체 통계량 사용
            If method = OD_MEDIAN Then replBase = medV Else replBase = meanV
        End If
    End If

    ' ── 이상치 분류 및 대체값 결정 ───────────────────────
    Dim outlierType() As Integer: ReDim outlierType(1 To rowCount)  ' 1=상한, -1=하한, 0=정상
    Dim replVal()     As Double:  ReDim replVal(1 To rowCount)
    Dim outlierCount  As Long: outlierCount = 0
    Dim numIdx        As Long: numIdx = 0  ' values() 배열 인덱스

    For idx = 1 To rowCount
        If isNumRow(idx) Then
            numIdx = numIdx + 1
            Dim dv As Double: dv = values(numIdx)
            If dv > hi Then
                outlierType(idx) = 1
                outlierCount = outlierCount + 1
                If method = OD_CAP Then replVal(idx) = hi Else replVal(idx) = replBase
            ElseIf dv < lo Then
                outlierType(idx) = -1
                outlierCount = outlierCount + 1
                If method = OD_CAP Then replVal(idx) = lo Else replVal(idx) = replBase
            Else
                outlierType(idx) = 0
                replVal(idx) = dv
            End If
        End If
    Next idx

    ' ── 결과 시트 이름 생성 ──────────────────────────────
    Dim memoSuf As String
    If Len(memo) > 0 Then memoSuf = "_" & OD_CleanName(memo)

    Dim baseName As String
    baseName = "이상치대체_" & OD_CleanName(varName) & memoSuf & "_" & Format(Now, "yyyymmdd_hhmm")

    Dim finalName As String: finalName = Left$(baseName, 31)
    Dim attempt   As Long:   attempt = 0
    Do While OD_SheetExists(finalName)
        attempt = attempt + 1
        finalName = Left$(baseName, 28) & "_" & attempt
    Loop

    ' ── 결과 시트 생성 ────────────────────────────────────
    Dim rWs As Worksheet
    Set rWs = ThisWorkbook.Worksheets.Add(After:=srcWs)
    rWs.Name = finalName

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wr As Long: wr = 1

    ' 섹션 1: 실행 조건 요약
    OD_WriteHdr rWs, wr, "실행 조건 요약"
    OD_WriteKV  rWs, wr, "실행 일시",    Format(Now, "yyyy-mm-dd hh:nn:ss")
    OD_WriteKV  rWs, wr, "소스 시트",    srcName
    OD_WriteKV  rWs, wr, "분석 변수",    varName
    OD_WriteKV  rWs, wr, "이상치 기준",  criteria
    OD_WriteKV  rWs, wr, "대체 방식",    method
    OD_WriteKV  rWs, wr, "쿼터 메모",    memo
    OD_WriteKV  rWs, wr, "분석 행 수",   rowCount
    OD_WriteKV  rWs, wr, "유효 수치 수", numCount
    OD_WriteKV  rWs, wr, "이상치 수",    outlierCount
    wr = wr + 1

    ' 섹션 2: 기본 통계량
    OD_WriteHdr rWs, wr, "기본 통계량"

    Dim sHdrs As Variant
    sHdrs = Array("N", "최솟값", "최댓값", "평균", "중앙값", "Q1", "Q3", "IQR", "하한", "상한")
    Dim sVals As Variant
    sVals = Array(numCount, minV, maxV, meanV, medV, q1, q3, iqrVal, lo, hi)

    Dim i As Long
    For i = 0 To UBound(sHdrs)
        With rWs.Cells(wr, i + 1)
            .Value = sHdrs(i)
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 245)
        End With
        rWs.Cells(wr + 1, i + 1).Value = sVals(i)
    Next i
    wr = wr + 3

    ' 섹션 3: 데이터 테이블
    OD_WriteHdr rWs, wr, "원본 데이터 및 이상치 대체 결과"

    Dim srcFC As Long: srcFC = dataArea.Column
    Dim colN  As Long: colN  = dataArea.Columns.Count

    ' 원본 헤더
    For i = 0 To colN - 1
        With rWs.Cells(wr, i + 1)
            .Value = srcWs.Cells(dataArea.Row, srcFC + i).Value
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 245)
        End With
    Next i

    ' 추가 컬럼 헤더
    Dim cOrig As Long: cOrig = colN + 1
    Dim cRepl As Long: cRepl = colN + 2
    Dim cFlag As Long: cFlag = colN + 3
    Dim cRsn  As Long: cRsn  = colN + 4
    Dim cDiff As Long: cDiff = colN + 5

    Dim addHdrs As Variant
    addHdrs = Array(varName & "_원본", varName & "_대체값", "이상치여부", "대체방식", "원본-대체값")
    For i = 0 To UBound(addHdrs)
        With rWs.Cells(wr, cOrig + i)
            .Value = addHdrs(i)
            .Font.Bold = True
            .Interior.Color = RGB(255, 235, 185)
        End With
    Next i
    wr = wr + 1

    ' 데이터 행 쓰기
    For idx = 1 To rowCount
        Dim srcR As Long: srcR = dataRows(idx)
        For i = 0 To colN - 1
            rWs.Cells(wr, i + 1).Value = srcWs.Cells(srcR, srcFC + i).Value
        Next i

        Dim origV As Variant: origV = srcWs.Cells(srcR, analysisCol).Value
        rWs.Cells(wr, cOrig).Value = origV

        If isNumRow(idx) Then
            Dim rv As Double: rv = replVal(idx)
            rWs.Cells(wr, cRepl).Value = rv
            rWs.Cells(wr, cDiff).Value = CDbl(origV) - rv

            If outlierType(idx) = 1 Then
                rWs.Cells(wr, cFlag).Value = "상한초과"
                rWs.Cells(wr, cRsn).Value  = criteria & " / " & method
                rWs.Cells(wr, cRepl).Interior.Color = RGB(255, 200, 200)  ' 연빨강
            ElseIf outlierType(idx) = -1 Then
                rWs.Cells(wr, cFlag).Value = "하한미만"
                rWs.Cells(wr, cRsn).Value  = criteria & " / " & method
                rWs.Cells(wr, cRepl).Interior.Color = RGB(200, 220, 255)  ' 연파랑
            Else
                rWs.Cells(wr, cFlag).Value = "-"
                rWs.Cells(wr, cRsn).Value  = "-"
            End If
        Else
            rWs.Cells(wr, cFlag).Value = "비해당"
            rWs.Cells(wr, cRsn).Value  = "-"
        End If

        wr = wr + 1
    Next idx

    rWs.UsedRange.Columns.AutoFit
    rWs.Activate
    Application.ScreenUpdating = True

    MsgBox "완료: " & finalName & vbCrLf & _
           "이상치 " & outlierCount & "개 / 유효값 " & numCount & "개", _
           vbInformation, "이상치 탐지"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    On Error Resume Next
    Application.DisplayAlerts = False
    rWs.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, "이상치 탐지"
End Sub

' ============================================================
' 프로시저 : OutlierDetect_ApplyToSelection
' 설    명  : 선택한 숫자 범위 바로 옆 열에 이상치 대체 결과를 만들고,
'             원본 이상치에는 조건부 서식과 강조 색을 적용한다.
' ============================================================
Public Sub OutlierDetect_ApplyToSelection(Optional ByVal criteriaOverride As String = "", _
                                          Optional ByVal methodOverride As String = "", _
                                          Optional ByVal skipConfirm As Boolean = False)
    On Error GoTo ErrHandler

    Dim suppressPrompts As Boolean
    suppressPrompts = skipConfirm _
                      Or Len(Trim$(criteriaOverride)) > 0 _
                      Or Len(Trim$(methodOverride)) > 0 _
                      Or Not Application.Visible

    If Not suppressPrompts Then
        If Not OD_PrepareInteractiveWorkspace() Then Exit Sub
    End If

    Dim dataArea As Range
    Dim targetRange As Range
    Dim analysisRelativeCol As Long

    Dim defaultAddress As String
    If TypeName(Selection) = "Range" Then
        defaultAddress = Selection.Areas(1).Address
    End If

    If suppressPrompts Then
        If TypeName(Selection) <> "Range" Then
            Err.Raise vbObjectError + 8101, "OutlierDetect_ApplyToSelection", "분석할 숫자 범위를 선택하세요."
        End If
        Set dataArea = Selection.Areas(1)
    Else
        If Not Toolkit_GetRangeFromUserForm("이상치 확인 - 범위 선택", _
            "엑셀 시트에서 이상치 확인 범위를 선택하세요. Ctrl+Shift+방향키, Shift+방향키, 마우스 드래그를 모두 사용할 수 있습니다.", _
            dataArea) Then Exit Sub
    End If

    If dataArea Is Nothing Then Exit Sub
    If dataArea.Areas.Count > 1 Then
        MsgBox "이상치 대체는 한 번에 하나의 연속된 범위만 처리합니다." & vbCrLf & _
               "분석할 숫자 열 하나 또는 표 범위 하나를 다시 선택하세요.", _
               vbExclamation, "이상치 대체"
        Exit Sub
    End If

    Set dataArea = OD_TrimToUsedRange(dataArea)
    If dataArea Is Nothing Then
        MsgBox "선택한 범위에서 사용 중인 데이터를 찾지 못했습니다.", vbExclamation, "이상치 대체"
        Exit Sub
    End If

    If dataArea.Columns.Count > 1 Then
        If Not OD_SelectNumericColumnFromRange(dataArea, suppressPrompts, analysisRelativeCol) Then Exit Sub
        Set targetRange = dataArea.Columns(analysisRelativeCol)
    Else
        Set targetRange = dataArea
    End If

    Dim ws As Worksheet: Set ws = targetRange.Worksheet
    Dim firstRow As Long: firstRow = targetRange.Row
    Dim lastRow As Long: lastRow = targetRange.Row + targetRange.Rows.Count - 1
    Dim headerRow As Long: headerRow = 0
    Dim firstDataRow As Long: firstDataRow = firstRow
    Dim varName As String

    If Len(DigitsOnly(CStr(ws.Cells(firstRow, targetRange.Column).Value))) = 0 Or Not IsNumeric(ws.Cells(firstRow, targetRange.Column).Value) Then
        headerRow = firstRow
        firstDataRow = firstRow + 1
        varName = Trim$(CStr(ws.Cells(firstRow, targetRange.Column).Value))
    End If
    If Len(varName) = 0 Then varName = ws.Cells(firstDataRow, targetRange.Column).Address(False, False)

    If Not suppressPrompts Then
        If Not OD_ConfirmHeaderAndRange(ws, targetRange, headerRow, firstDataRow, varName) Then Exit Sub
    End If

    If firstDataRow > lastRow Then
        MsgBox "분석할 숫자 데이터가 없습니다.", vbExclamation, "이상치 대체"
        Exit Sub
    End If

    Dim rowCount As Long: rowCount = lastRow - firstDataRow + 1
    Dim values() As Double: ReDim values(1 To rowCount)
    Dim rowMap() As Long: ReDim rowMap(1 To rowCount)
    Dim numCount As Long: numCount = 0
    Dim blankCount As Long: blankCount = 0
    Dim nonNumericCount As Long: nonNumericCount = 0
    Dim r As Long

    For r = firstDataRow To lastRow
        If Not ws.Rows(r).Hidden Then
            Dim cellValue As Variant
            cellValue = ws.Cells(r, targetRange.Column).Value
            Dim numericValue As Double
            If OD_TryNumericValue(cellValue, numericValue) Then
                numCount = numCount + 1
                values(numCount) = numericValue
                rowMap(numCount) = r
            ElseIf IsError(cellValue) Or IsEmpty(cellValue) Or Len(Trim$(CStr(cellValue))) = 0 Then
                blankCount = blankCount + 1
            Else
                nonNumericCount = nonNumericCount + 1
            End If
        End If
    Next r

    If numCount = 0 Then
        MsgBox "선택 범위에서 유효한 숫자값을 찾지 못했습니다.", vbExclamation, "이상치 대체"
        Exit Sub
    End If

    ReDim Preserve values(1 To numCount)
    ReDim Preserve rowMap(1 To numCount)

    If Not suppressPrompts Then
        If MsgBox("범위 확인" & vbCrLf & vbCrLf & _
                  "선택 범위: " & targetRange.Address(False, False) & vbCrLf & _
                  "헤더/변수명: " & varName & vbCrLf & _
                  "데이터 행 범위: " & firstDataRow & "행 ~ " & lastRow & "행" & vbCrLf & _
                  "숫자값: " & numCount & "개" & vbCrLf & _
                  "빈칸/오류: " & blankCount & "개" & vbCrLf & _
                  "숫자가 아닌 값: " & nonNumericCount & "개" & vbCrLf & vbCrLf & _
                  "이 범위로 기초통계량표와 Q1/Q3 검토열을 생성할까요?", _
                  vbQuestion + vbYesNo, "이상치 확인 - 범위 검증") = vbNo Then Exit Sub
    End If

    Dim sortedValues() As Double
    OD_CopyAndSortValues values, numCount, sortedValues

    Dim q1 As Double: q1 = OD_PercentileIncSorted(sortedValues, numCount, 0.25)
    Dim q3 As Double: q3 = OD_PercentileIncSorted(sortedValues, numCount, 0.75)
    Dim iqrVal As Double: iqrVal = q3 - q1
    Dim medV As Double: medV = OD_PercentileIncSorted(sortedValues, numCount, 0.5)
    Dim meanV As Double: meanV = OD_AverageValues(values, numCount)
    Dim minV As Double: minV = sortedValues(1)
    Dim maxV As Double: maxV = sortedValues(numCount)
    Dim lo As Double: lo = q1 - 1.5 * iqrVal
    Dim hi As Double: hi = q3 + 1.5 * iqrVal
    Dim reviewCount As Long: reviewCount = 0
    Dim thresholdCol As Long: thresholdCol = targetRange.Column + 1
    Dim copyCol As Long: copyCol = targetRange.Column + 2
    Dim statsCol As Long: statsCol = targetRange.Column + 8

    Application.ScreenUpdating = False
    OD_WriteQuartileReviewStats ws, headerRow, firstDataRow, statsCol, _
                                varName, targetRange.Address(False, False), numCount, _
                                blankCount, nonNumericCount, minV, maxV, meanV, medV, _
                                q1, q3, iqrVal, lo, hi

    If Not Toolkit_CanInsertColumns(ws, thresholdCol, 2, "이상치 대체") Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ws.Columns(thresholdCol).Resize(, 2).Insert Shift:=xlToRight
    If headerRow > 0 Then
        ws.Cells(headerRow, thresholdCol).Value = varName & "_Q1_Q3기준"
        ws.Cells(headerRow, copyCol).Value = varName & "_검토복사"
        ws.Range(ws.Cells(headerRow, thresholdCol), ws.Cells(headerRow, copyCol)).Font.Bold = True
        ws.Range(ws.Cells(headerRow, thresholdCol), ws.Cells(headerRow, copyCol)).Interior.Color = RGB(255, 235, 185)
    End If

    Dim sourceDataRange As Range
    Dim thresholdDataRange As Range
    Dim copiedDataRange As Range
    Set sourceDataRange = ws.Range(ws.Cells(firstDataRow, targetRange.Column), ws.Cells(lastRow, targetRange.Column))
    Set thresholdDataRange = ws.Range(ws.Cells(firstDataRow, thresholdCol), ws.Cells(lastRow, thresholdCol))
    Set copiedDataRange = ws.Range(ws.Cells(firstDataRow, copyCol), ws.Cells(lastRow, copyCol))
    copiedDataRange.NumberFormat = sourceDataRange.NumberFormat
    copiedDataRange.Value = sourceDataRange.Value

    Dim i As Long
    For i = 1 To numCount
        Dim v As Double: v = values(i)
        Dim outR As Long: outR = rowMap(i)
        If v < q1 Then
            reviewCount = reviewCount + 1
            ws.Cells(outR, thresholdCol).Value = q1
        ElseIf v > q3 Then
            reviewCount = reviewCount + 1
            ws.Cells(outR, thresholdCol).Value = q3
        End If
    Next i

    thresholdDataRange.NumberFormat = sourceDataRange.NumberFormat
    OD_ApplyQuartileReviewFormatting copiedDataRange, thresholdDataRange

    Dim displayTopRow As Long
    If headerRow > 0 Then displayTopRow = headerRow Else displayTopRow = firstDataRow
    ws.Range(ws.Cells(displayTopRow, thresholdCol), ws.Cells(lastRow, copyCol)).Columns.AutoFit
    ws.Range(ws.Cells(displayTopRow, statsCol + 2), ws.Cells(displayTopRow + 24, statsCol + 3)).Columns.AutoFit
    Application.ScreenUpdating = True

    If Not suppressPrompts Then
        MsgBox "Q1/Q3 검토열을 만들었습니다." & vbCrLf & _
               "유효값: " & numCount & "개" & vbCrLf & _
               "Q1 미만 또는 Q3 초과 표시: " & reviewCount & "개" & vbCrLf & _
               "기준값 열: " & ws.Cells(1, thresholdCol).EntireColumn.Address(False, False) & vbCrLf & _
               "검토복사 열: " & ws.Cells(1, copyCol).EntireColumn.Address(False, False), _
               vbInformation, "이상치 확인 완료"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    If suppressPrompts Then
        Err.Raise Err.Number, "OutlierDetect_ApplyToSelection", Err.Description
    Else
        MsgBox "이상치 대체 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "이상치 대체"
    End If
End Sub

' ──────────────────────────────────────────────────────────────
' OutlierDetector 내부 헬퍼
' ──────────────────────────────────────────────────────────────

Private Function OD_PrepareInteractiveWorkspace() As Boolean
    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "이상치 탐지 전 확인" & vbCrLf & vbCrLf & _
        "이 기능은 통계적으로 의심되는 값을 표시하는 보조 도구입니다." & vbCrLf & _
        "실행 전에 가능한 경우 다음과 같이 실제 확인 가능한 대체 절차를 먼저 검토하세요." & vbCrLf & _
        "- 2차 자료 대체" & vbCrLf & _
        "- 전년도 자료 대체" & vbCrLf & _
        "- 원자료 또는 응답자 재확인" & vbCrLf & vbCrLf & _
        "예: 안내용 작업 시트를 만들고 계속합니다." & vbCrLf & _
        "아니요: 이미 준비한 범위를 바로 선택합니다." & vbCrLf & _
        "취소: 실행을 중단합니다.", _
        vbQuestion + vbYesNoCancel, "이상치 탐지 - 사전 확인")

    Select Case answer
        Case vbCancel
            Exit Function
        Case vbNo
            OD_PrepareInteractiveWorkspace = True
            Exit Function
    End Select

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "이상치 탐지 작업 시트를 만들 통합문서를 먼저 여세요.", vbExclamation, "이상치 탐지"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = OD_CreateWorkspaceSheet(wb)
    ws.Activate

    answer = MsgBox( _
        "이상치 탐지 작업 시트를 만들었습니다." & vbCrLf & vbCrLf & _
        "시트 안내에 따라 ID, 업종, 매출액 규모, 종사자 규모, 대체 대상 값을 붙여넣으세요." & vbCrLf & _
        "지금 범위 선택을 계속하려면 확인을 누르고, 안내 시트만 준비하려면 취소를 누르세요.", _
        vbInformation + vbOKCancel, "이상치 탐지 - 작업 시트")
    OD_PrepareInteractiveWorkspace = (answer = vbOK)
End Function

Private Function OD_CreateWorkspaceSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = SPSS_UniqueSheetName(wb, OD_WORK_SHEET)

    OD_WriteWorkspaceGuide ws
    Set OD_CreateWorkspaceSheet = ws
End Function

Private Sub OD_WriteWorkspaceGuide(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Tab.Color = RGB(255, 192, 0)

    With ws.Range("A1:H1")
        .Merge
        .Value = "이상치 탐지 작업 시트"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = RGB(255, 235, 185)
        .HorizontalAlignment = xlLeft
    End With

    ws.Range("A3").Value = "사용 전 확인"
    ws.Range("A4").Value = "이상치 탐지는 최종 대체 결정을 자동으로 보장하지 않습니다. 가능한 경우 2차 자료, 전년도 자료, 원자료 또는 응답자 확인처럼 실제로 확인 가능한 대체 절차를 먼저 검토하세요."

    ws.Range("A6").Value = "권장 작업 흐름"
    ws.Range("A7").Value = "1. 원본 데이터에서 대체가 필요한 후보만 필터로 좁힙니다."
    ws.Range("A8").Value = "2. ID, 업종, 규모 구분, 실제 대체가 필요한 숫자값을 아래 예시 열에 맞춰 붙여넣습니다."
    ws.Range("A9").Value = "3. 대체 대상 값의 헤더와 데이터 범위를 선택한 뒤 [데이터 툴킷] > [이상치 탐지] > [분석 열 선택]을 실행합니다."
    ws.Range("A10").Value = "4. 생성된 Q1/Q3 기준열, 검토복사 열, 기초통계량표를 확인하고 필요한 값만 직접 수정합니다."

    ws.Range("A12").Value = "붙여넣기 예시"
    Dim headers As Variant
    headers = Array("ID", "업종", "매출액 규모", "종사자 규모", "대체대상_원값", "2차자료_대체값", "전년도자료_대체값", "검토메모")

    Dim samples As Variant
    samples = Array("A001", "제조", "10억 미만", "10~49명", 123456789, "공시자료 값", "전년도 값", "확인 필요")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(14, i + 1).Value = headers(i)
        ws.Cells(15, i + 1).Value = samples(i)
    Next i

    With ws.Range("A14:H14")
        .Font.Bold = True
        .Interior.Color = RGB(220, 235, 250)
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A15:H15")
        .Interior.Color = RGB(245, 245, 245)
        .Font.Color = RGB(120, 120, 120)
        .Borders.LineStyle = xlContinuous
    End With

    ws.Range("A17").Value = "선택 범위 기준"
    ws.Range("A18").Value = "대체 대상 숫자 열은 헤더 1행과 실제 숫자 데이터가 포함되도록 선택하세요. 예: E14:E500"
    ws.Range("A19").Value = "표 전체를 선택하면 숫자형 열을 자동 탐지한 뒤 분석할 열을 선택할 수 있습니다."
    ws.Range("A20").Value = "숫자가 텍스트로 저장되어 있다면 먼저 [값 정리] > [숫자 변환]을 실행하세요."

    With ws.Range("A3,A6,A12,A17")
        .Font.Bold = True
        .Font.Color = RGB(50, 80, 120)
    End With
    ws.Range("A4:H10").WrapText = True
    ws.Range("A18:H20").WrapText = True
    ws.Columns("A:H").AutoFit
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("E:H").ColumnWidth = 18
    ws.Rows("1:20").VerticalAlignment = xlTop
    ws.Range("E14:E15").NumberFormat = "#,##0"
    ws.Activate
    ws.Range("A1").Select
End Sub

Private Function OD_TrimToUsedRange(ByVal selectedRange As Range) As Range
    Dim usedPart As Range
    On Error Resume Next
    Set usedPart = Application.Intersect(selectedRange, selectedRange.Worksheet.UsedRange)
    On Error GoTo 0

    If usedPart Is Nothing Then
        Set OD_TrimToUsedRange = selectedRange
    Else
        Set OD_TrimToUsedRange = usedPart
    End If
End Function

Private Sub OD_CopyAndSortValues(ByRef sourceValues() As Double, _
                                 ByVal countN As Long, _
                                 ByRef sortedValues() As Double)
    If countN <= 0 Then Err.Raise vbObjectError + 8201, "OD_CopyAndSortValues", "정렬할 숫자값이 없습니다."

    ReDim sortedValues(1 To countN)

    Dim i As Long
    For i = 1 To countN
        sortedValues(i) = sourceValues(i)
    Next i

    If countN > 1 Then OD_QuickSortDoubles sortedValues, 1, countN
End Sub

Private Sub OD_QuickSortDoubles(ByRef values() As Double, ByVal firstIdx As Long, ByVal lastIdx As Long)
    Dim lowIdx As Long: lowIdx = firstIdx
    Dim highIdx As Long: highIdx = lastIdx
    Dim pivot As Double: pivot = values((firstIdx + lastIdx) \ 2)
    Dim tempValue As Double

    Do While lowIdx <= highIdx
        Do While values(lowIdx) < pivot
            lowIdx = lowIdx + 1
        Loop
        Do While values(highIdx) > pivot
            highIdx = highIdx - 1
        Loop

        If lowIdx <= highIdx Then
            tempValue = values(lowIdx)
            values(lowIdx) = values(highIdx)
            values(highIdx) = tempValue
            lowIdx = lowIdx + 1
            highIdx = highIdx - 1
        End If
    Loop

    If firstIdx < highIdx Then OD_QuickSortDoubles values, firstIdx, highIdx
    If lowIdx < lastIdx Then OD_QuickSortDoubles values, lowIdx, lastIdx
End Sub

Private Function OD_PercentileIncSorted(ByRef sortedValues() As Double, _
                                        ByVal countN As Long, _
                                        ByVal percentile As Double) As Double
    If countN <= 0 Then Err.Raise vbObjectError + 8202, "OD_PercentileIncSorted", "백분위수를 계산할 숫자값이 없습니다."
    If percentile < 0 Or percentile > 1 Then Err.Raise vbObjectError + 8203, "OD_PercentileIncSorted", "백분위수는 0과 1 사이여야 합니다."
    If countN = 1 Then
        OD_PercentileIncSorted = sortedValues(1)
        Exit Function
    End If

    Dim position As Double
    position = 1 + (countN - 1) * percentile

    Dim lowerIdx As Long
    lowerIdx = Fix(position)

    Dim upperIdx As Long
    upperIdx = lowerIdx + 1
    If upperIdx > countN Then upperIdx = countN

    Dim fraction As Double
    fraction = position - lowerIdx

    OD_PercentileIncSorted = sortedValues(lowerIdx) + fraction * (sortedValues(upperIdx) - sortedValues(lowerIdx))
End Function

Private Function OD_MedianValues(ByRef values() As Double, ByVal countN As Long) As Double
    Dim sortedValues() As Double
    OD_CopyAndSortValues values, countN, sortedValues
    OD_MedianValues = OD_PercentileIncSorted(sortedValues, countN, 0.5)
End Function

Private Function OD_AverageValues(ByRef values() As Double, ByVal countN As Long) As Double
    If countN <= 0 Then Err.Raise vbObjectError + 8204, "OD_AverageValues", "평균을 계산할 숫자값이 없습니다."

    Dim total As Double
    Dim i As Long
    For i = 1 To countN
        total = total + values(i)
    Next i

    OD_AverageValues = total / countN
End Function

Private Function OD_SelectEasyCriteria(ByRef criteria As String) As Boolean
    Dim criteriaInput As String
    criteriaInput = Toolkit_ButtonMenu( _
        "이상치 대체 - 기준 선택", _
        "이상치 기준을 선택하세요.", _
        "1" & vbTab & OD_IQR15 & " : 일반적인 윈저라이징 기준" & vbLf & _
        "2" & vbTab & OD_IQR20 & " : 조금 더 완화된 IQR 기준" & vbLf & _
        "3" & vbTab & OD_P95 & " : 하위 5% / 상위 5% 기준" & vbLf & _
        "4" & vbTab & OD_P99 & " : 하위 1% / 상위 1% 기준", _
        "1")
    If Len(Trim$(criteriaInput)) = 0 Then Exit Function

    If Not OD_MapEasyCriteria(criteriaInput, criteria) Then
        MsgBox "이상치 기준은 1~4 중 하나를 입력하세요.", vbExclamation, "이상치 대체"
        Exit Function
    End If

    OD_SelectEasyCriteria = True
End Function

Private Function OD_SelectNumericColumnFromRange(ByVal dataArea As Range, _
                                                 ByVal skipConfirm As Boolean, _
                                                 ByRef relativeCol As Long) As Boolean
    Dim numericCols As Collection
    Set numericCols = New Collection

    Dim c As Long
    For c = 1 To dataArea.Columns.Count
        Dim firstDataRow As Long
        Dim headerText As String
        Dim n As Long
        Dim outN As Long

        firstDataRow = OD_FirstDataRowForColumn(dataArea, c, headerText)
        n = OD_CountNumericValues(dataArea.Worksheet, dataArea.Column + c - 1, firstDataRow, dataArea.Row + dataArea.Rows.Count - 1)
        If n > 0 Then
            outN = OD_CountIqrOutliers(dataArea.Worksheet, dataArea.Column + c - 1, firstDataRow, dataArea.Row + dataArea.Rows.Count - 1)
            If Len(headerText) = 0 Then headerText = "열" & c
            numericCols.Add Array(c, headerText, n, outN)
        End If
    Next c

    If numericCols.Count = 0 Then
        MsgBox "선택 범위에서 숫자형 열을 찾지 못했습니다." & vbCrLf & _
               "숫자가 텍스트로 저장된 경우 먼저 [숫자 변환]을 실행하세요.", vbExclamation, "이상치 대체"
        Exit Function
    End If

    If numericCols.Count = 1 Then
        Dim onlyItem As Variant
        onlyItem = numericCols(1)
        relativeCol = CLng(onlyItem(0))
        OD_SelectNumericColumnFromRange = True
        Exit Function
    End If

    If skipConfirm Then
        Dim bestIdx As Long: bestIdx = 1
        Dim bestItem As Variant
        Dim curItem As Variant
        For c = 2 To numericCols.Count
            curItem = numericCols(c)
            bestItem = numericCols(bestIdx)
            If CLng(curItem(3)) > CLng(bestItem(3)) Then bestIdx = c
        Next c
        bestItem = numericCols(bestIdx)
        relativeCol = CLng(bestItem(0))
        OD_SelectNumericColumnFromRange = True
        Exit Function
    End If

    Dim optionsText As String

    For c = 1 To numericCols.Count
        Dim listItem As Variant
        listItem = numericCols(c)
        If Len(optionsText) > 0 Then optionsText = optionsText & vbLf
        optionsText = optionsText & c & vbTab & CStr(listItem(1)) & _
                      " (유효값 " & CLng(listItem(2)) & _
                      "개, IQR 이상치 " & CLng(listItem(3)) & "개)"
    Next c

    Dim userInput As String
    userInput = Toolkit_ButtonMenu("이상치 대체 - 숫자 열 선택", _
                                   "선택 범위에서 숫자형 열을 찾았습니다. 이상치를 탐지할 열을 선택하세요.", _
                                   optionsText, "1")
    If Len(Trim$(userInput)) = 0 Then Exit Function
    If Not IsNumeric(userInput) Then
        MsgBox "목록의 번호를 입력하세요.", vbExclamation, "이상치 대체"
        Exit Function
    End If

    Dim selectedIdx As Long
    selectedIdx = CLng(userInput)
    If selectedIdx < 1 Or selectedIdx > numericCols.Count Then
        MsgBox "선택 번호가 범위를 벗어났습니다.", vbExclamation, "이상치 대체"
        Exit Function
    End If

    Dim selectedItem As Variant
    selectedItem = numericCols(selectedIdx)
    relativeCol = CLng(selectedItem(0))
    OD_SelectNumericColumnFromRange = True
End Function

Private Function OD_FirstDataRowForColumn(ByVal dataArea As Range, _
                                          ByVal relativeCol As Long, _
                                          ByRef headerText As String) As Long
    Dim firstCell As Range
    Set firstCell = dataArea.Cells(1, relativeCol)
    headerText = Trim$(CStr(firstCell.Value))

    If Len(headerText) = 0 Or Not IsNumeric(firstCell.Value) Then
        OD_FirstDataRowForColumn = dataArea.Row + 1
    Else
        headerText = ""
        OD_FirstDataRowForColumn = dataArea.Row
    End If
End Function

Private Function OD_ConfirmHeaderAndRange(ByVal ws As Worksheet, _
                                          ByVal targetRange As Range, _
                                          ByRef headerRow As Long, _
                                          ByRef firstDataRow As Long, _
                                          ByRef varName As String) As Boolean
    Dim lastRow As Long
    lastRow = targetRange.Row + targetRange.Rows.Count - 1

    Do
        Dim headerInfo As String
        If headerRow > 0 Then
            headerInfo = "헤더 행: " & headerRow & "행" & vbCrLf & _
                         "헤더 값: " & Trim$(CStr(ws.Cells(headerRow, targetRange.Column).Value))
        Else
            headerInfo = "헤더 행: 없음" & vbCrLf & _
                         "변수명: " & varName
        End If

        Dim answer As VbMsgBoxResult
        answer = MsgBox("이상치 확인 범위를 검토하세요." & vbCrLf & vbCrLf & _
                        "분석 열: " & targetRange.Address(False, False) & vbCrLf & _
                        headerInfo & vbCrLf & _
                        "데이터 시작 행: " & firstDataRow & "행" & vbCrLf & _
                        "데이터 끝 행: " & lastRow & "행" & vbCrLf & vbCrLf & _
                        "[예] 이 설정으로 진행" & vbCrLf & _
                        "[아니오] 헤더 행을 직접 지정" & vbCrLf & _
                        "[취소] 작업 중단", _
                        vbQuestion + vbYesNoCancel, "이상치 확인 - 헤더/범위 검토")

        If answer = vbYes Then
            OD_ConfirmHeaderAndRange = True
            Exit Function
        ElseIf answer = vbCancel Then
            Exit Function
        End If

        Dim headerInput As String
        headerInput = InputBox("헤더 행 번호를 입력하세요." & vbCrLf & vbCrLf & _
                               "0: 헤더 없음, 선택 범위 첫 행부터 데이터로 사용" & vbCrLf & _
                               CStr(targetRange.Row) & "~" & CStr(lastRow) & ": 해당 행을 헤더로 사용", _
                               "이상치 확인 - 헤더 행 지정", IIf(headerRow > 0, CStr(headerRow), "0"))
        If Len(Trim$(headerInput)) = 0 Then Exit Function
        If Not IsNumeric(headerInput) Then
            MsgBox "헤더 행은 숫자로 입력하세요.", vbExclamation, "이상치 확인"
        Else
            Dim newHeaderRow As Long
            newHeaderRow = CLng(headerInput)
            If newHeaderRow = 0 Then
                headerRow = 0
                firstDataRow = targetRange.Row
                varName = targetRange.Cells(1, 1).Address(False, False)
            ElseIf newHeaderRow < targetRange.Row Or newHeaderRow >= lastRow Then
                MsgBox "헤더 행은 선택 범위 안에 있어야 하며, 헤더 아래에 데이터가 1행 이상 필요합니다.", _
                       vbExclamation, "이상치 확인"
            Else
                headerRow = newHeaderRow
                firstDataRow = newHeaderRow + 1
                varName = Trim$(CStr(ws.Cells(headerRow, targetRange.Column).Value))
                If Len(varName) = 0 Then varName = ws.Cells(firstDataRow, targetRange.Column).Address(False, False)
            End If
        End If
    Loop
End Function

Private Function OD_TryNumericValue(ByVal valueIn As Variant, ByRef valueOut As Double) As Boolean
    If IsError(valueIn) Or IsEmpty(valueIn) Then Exit Function

    If IsNumeric(valueIn) Then
        valueOut = CDbl(valueIn)
        OD_TryNumericValue = True
        Exit Function
    End If

    Dim textValue As String
    textValue = Trim$(CStr(valueIn))
    If Len(textValue) = 0 Then Exit Function
    textValue = Replace(textValue, ",", "")

    If IsNumeric(textValue) Then
        valueOut = CDbl(textValue)
        OD_TryNumericValue = True
    End If
End Function

Private Function OD_CountNumericValues(ByVal ws As Worksheet, _
                                       ByVal absoluteCol As Long, _
                                       ByVal firstDataRow As Long, _
                                       ByVal lastRow As Long) As Long
    Dim r As Long
    For r = firstDataRow To lastRow
        If Not ws.Rows(r).Hidden Then
            Dim v As Variant
            v = ws.Cells(r, absoluteCol).Value
            Dim numericValue As Double
            If OD_TryNumericValue(v, numericValue) Then
                OD_CountNumericValues = OD_CountNumericValues + 1
            End If
        End If
    Next r
End Function

Private Function OD_CountIqrOutliers(ByVal ws As Worksheet, _
                                     ByVal absoluteCol As Long, _
                                     ByVal firstDataRow As Long, _
                                     ByVal lastRow As Long) As Long
    Dim countN As Long
    countN = OD_CountNumericValues(ws, absoluteCol, firstDataRow, lastRow)
    If countN = 0 Then Exit Function

    Dim values() As Double
    ReDim values(1 To countN)

    Dim r As Long
    Dim idx As Long
    For r = firstDataRow To lastRow
        If Not ws.Rows(r).Hidden Then
            Dim v As Variant
            v = ws.Cells(r, absoluteCol).Value
            Dim numericValue As Double
            If OD_TryNumericValue(v, numericValue) Then
                idx = idx + 1
                values(idx) = numericValue
            End If
        End If
    Next r

    Dim sortedValues() As Double
    OD_CopyAndSortValues values, countN, sortedValues

    Dim q1 As Double: q1 = OD_PercentileIncSorted(sortedValues, countN, 0.25)
    Dim q3 As Double: q3 = OD_PercentileIncSorted(sortedValues, countN, 0.75)
    Dim lo As Double: lo = q1 - 1.5 * (q3 - q1)
    Dim hi As Double: hi = q3 + 1.5 * (q3 - q1)

    For idx = 1 To countN
        If values(idx) < lo Or values(idx) > hi Then OD_CountIqrOutliers = OD_CountIqrOutliers + 1
    Next idx
End Function

Private Function OD_SelectEasyMethod(ByRef method As String) As Boolean
    Dim methodInput As String
    methodInput = Toolkit_ButtonMenu( _
        "이상치 대체 - 방식 선택", _
        "이상치 대체 방식을 선택하세요.", _
        "1" & vbTab & OD_CAP & " : 하한보다 작으면 하한, 상한보다 크면 상한으로 대체" & vbLf & _
        "2" & vbTab & OD_MEDIAN & " : 정상 범위 데이터의 중앙값으로 대체" & vbLf & _
        "3" & vbTab & OD_MEAN & " : 정상 범위 데이터의 평균으로 대체" & vbLf & _
        "4" & vbTab & OD_REVIEW & " : 옆 열에 원본을 복사하고 이상치만 표시", _
        "1")
    If Len(Trim$(methodInput)) = 0 Then Exit Function

    If Not OD_MapEasyMethod(methodInput, method) Then
        MsgBox "대체 방식은 1~4 중 하나를 입력하세요.", vbExclamation, "이상치 대체"
        Exit Function
    End If

    OD_SelectEasyMethod = True
End Function

Private Function OD_MapEasyCriteria(ByVal userInput As String, ByRef criteria As String) As Boolean
    Select Case Trim$(userInput)
        Case "1", OD_IQR15
            criteria = OD_IQR15
        Case "2", OD_IQR20
            criteria = OD_IQR20
        Case "3", OD_P95
            criteria = OD_P95
        Case "4", OD_P99
            criteria = OD_P99
        Case Else
            Exit Function
    End Select
    OD_MapEasyCriteria = True
End Function

Private Function OD_MapEasyMethod(ByVal userInput As String, ByRef method As String) As Boolean
    Select Case Trim$(userInput)
        Case "1", OD_CAP
            method = OD_CAP
        Case "2", OD_MEDIAN
            method = OD_MEDIAN
        Case "3", OD_MEAN
            method = OD_MEAN
        Case "4", OD_REVIEW
            method = OD_REVIEW
        Case Else
            Exit Function
    End Select
    OD_MapEasyMethod = True
End Function

Private Function OD_ReplacementBase(ByRef values() As Double, _
                                    ByVal numCount As Long, _
                                    ByVal lo As Double, _
                                    ByVal hi As Double, _
                                    ByVal method As String, _
                                    ByVal medV As Double, _
                                    ByVal meanV As Double) As Double
    Dim trimVals() As Double
    Dim trimN As Long
    Dim i As Long

    ReDim trimVals(1 To numCount)
    For i = 1 To numCount
        If values(i) >= lo And values(i) <= hi Then
            trimN = trimN + 1
            trimVals(trimN) = values(i)
        End If
    Next i

    If trimN = 0 Then
        If method = OD_MEDIAN Then OD_ReplacementBase = medV Else OD_ReplacementBase = meanV
        Exit Function
    End If

    ReDim Preserve trimVals(1 To trimN)
    If method = OD_MEDIAN Then
        OD_ReplacementBase = OD_MedianValues(trimVals, trimN)
    Else
        OD_ReplacementBase = OD_AverageValues(trimVals, trimN)
    End If
End Function

Private Sub OD_ApplyOutlierConditionalFormatting(ByVal target As Range, ByVal lo As Double, ByVal hi As Double)
    On Error Resume Next
    target.FormatConditions.Delete
    On Error GoTo 0

    Dim lowerRule As FormatCondition
    Dim upperRule As FormatCondition
    Set lowerRule = target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=CStr(lo))
    lowerRule.Interior.Color = RGB(210, 225, 255)
    lowerRule.Font.Color = RGB(0, 60, 150)

    Set upperRule = target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:=CStr(hi))
    upperRule.Interior.Color = RGB(255, 210, 210)
    upperRule.Font.Color = RGB(160, 0, 0)
End Sub

Private Sub OD_ApplyQuartileReviewFormatting(ByVal copiedDataRange As Range, ByVal thresholdDataRange As Range)
    On Error Resume Next
    copiedDataRange.FormatConditions.Delete
    thresholdDataRange.FormatConditions.Delete
    On Error GoTo 0

    Dim formulaText As String
    formulaText = "=" & thresholdDataRange.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=True) & "<>" & Chr$(34) & Chr$(34)

    Dim copyRule As FormatCondition
    Set copyRule = copiedDataRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaText)
    copyRule.Interior.Color = RGB(255, 245, 190)
    copyRule.Font.Bold = True

    Dim thresholdRule As FormatCondition
    Set thresholdRule = thresholdDataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=" & Chr$(34) & Chr$(34))
    thresholdRule.Interior.Color = RGB(255, 235, 185)
    thresholdRule.Font.Color = RGB(160, 90, 0)
    thresholdRule.Font.Bold = True
End Sub

Private Sub OD_WriteQuartileReviewStats(ws As Worksheet, _
                                        ByVal headerRow As Long, _
                                        ByVal firstDataRow As Long, _
                                        ByVal startCol As Long, _
                                        ByVal varName As String, _
                                        ByVal selectedAddress As String, _
                                        ByVal numCount As Long, _
                                        ByVal blankCount As Long, _
                                        ByVal nonNumericCount As Long, _
                                        ByVal minV As Double, _
                                        ByVal maxV As Double, _
                                        ByVal meanV As Double, _
                                        ByVal medV As Double, _
                                        ByVal q1 As Double, _
                                        ByVal q3 As Double, _
                                        ByVal iqrVal As Double, _
                                        ByVal lo As Double, _
                                        ByVal hi As Double)
    Dim topRow As Long
    If headerRow > 0 Then topRow = headerRow Else topRow = firstDataRow
    If topRow < 1 Then topRow = 1

    Dim leftCol As Long: leftCol = startCol
    Dim valueCol As Long: valueCol = startCol + 1

    ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow + 18, valueCol)).Clear
    ws.Cells(topRow, leftCol).Value = "기초통계량"
    ws.Cells(topRow, leftCol).Font.Bold = True
    ws.Cells(topRow, leftCol).Font.Size = 12
    ws.Cells(topRow, leftCol).Interior.Color = RGB(40, 75, 115)
    ws.Cells(topRow, leftCol).Font.Color = RGB(255, 255, 255)
    ws.Cells(topRow, valueCol).Interior.Color = RGB(40, 75, 115)

    Dim rowNo As Long: rowNo = topRow + 1
    OD_WriteStatRow ws, rowNo, leftCol, "변수명", varName
    OD_WriteStatRow ws, rowNo, leftCol, "선택 범위", selectedAddress
    OD_WriteStatRow ws, rowNo, leftCol, "숫자값", numCount
    OD_WriteStatRow ws, rowNo, leftCol, "빈칸/오류", blankCount
    OD_WriteStatRow ws, rowNo, leftCol, "숫자가 아닌 값", nonNumericCount
    OD_WriteStatRow ws, rowNo, leftCol, "최솟값", minV
    OD_WriteStatRow ws, rowNo, leftCol, "Q1", q1
    OD_WriteStatRow ws, rowNo, leftCol, "중앙값", medV
    OD_WriteStatRow ws, rowNo, leftCol, "평균", meanV
    OD_WriteStatRow ws, rowNo, leftCol, "Q3", q3
    OD_WriteStatRow ws, rowNo, leftCol, "최댓값", maxV
    OD_WriteStatRow ws, rowNo, leftCol, "IQR", iqrVal
    OD_WriteStatRow ws, rowNo, leftCol, "IQR 하한", lo
    OD_WriteStatRow ws, rowNo, leftCol, "IQR 상한", hi

    With ws.Range(ws.Cells(topRow + 1, leftCol), ws.Cells(rowNo - 1, valueCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(215, 220, 230)
    End With
    ws.Range(ws.Cells(topRow + 1, leftCol), ws.Cells(rowNo - 1, leftCol)).Interior.Color = RGB(235, 240, 248)
    ws.Range(ws.Cells(topRow + 1, valueCol), ws.Cells(rowNo - 1, valueCol)).NumberFormat = "0.####"
End Sub

Private Sub OD_WriteStatRow(ws As Worksheet, ByRef rowNo As Long, ByVal startCol As Long, ByVal labelText As String, ByVal valueData As Variant)
    ws.Cells(rowNo, startCol).Value = labelText
    ws.Cells(rowNo, startCol + 1).Value = valueData
    rowNo = rowNo + 1
End Sub

Private Sub OD_WriteEasyStats(ws As Worksheet, _
                              ByVal headerRow As Long, _
                              ByVal firstDataRow As Long, _
                              ByVal startCol As Long, _
                              ByVal varName As String, _
                              ByVal criteria As String, _
                              ByVal method As String, _
                              ByVal numCount As Long, _
                              ByVal outlierCount As Long, _
                              ByVal minV As Double, _
                              ByVal maxV As Double, _
                              ByVal meanV As Double, _
                              ByVal medV As Double, _
                              ByVal q1 As Double, _
                              ByVal q3 As Double, _
                              ByVal iqrVal As Double, _
                              ByVal lo As Double, _
                              ByVal hi As Double, _
                              ByVal replBase As Double)
    Dim topRow As Long
    If headerRow > 0 Then topRow = headerRow Else topRow = firstDataRow

    Dim labels As Variant
    Dim vals As Variant
    labels = Array("변수", "기준", "대체 방식", "유효값 수", "이상치 수", "최솟값", "최댓값", "평균", "중앙값", "Q1", "Q3", "IQR", "하한", "상한", "대체 기준값")
    vals = Array(varName, criteria, method, numCount, outlierCount, minV, maxV, meanV, medV, q1, q3, iqrVal, lo, hi, IIf(method = OD_MEDIAN Or method = OD_MEAN, replBase, "-"))

    With ws.Cells(topRow, startCol)
        .Value = "이상치 대체 통계"
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 245)
    End With
    ws.Cells(topRow, startCol + 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        ws.Cells(topRow + i + 1, startCol).Value = labels(i)
        ws.Cells(topRow + i + 1, startCol).Font.Bold = True
        ws.Cells(topRow + i + 1, startCol + 1).Value = vals(i)
    Next i

    ws.Cells(topRow + UBound(labels) + 3, startCol).Value = "색상 안내"
    ws.Cells(topRow + UBound(labels) + 3, startCol).Font.Bold = True
    ws.Cells(topRow + UBound(labels) + 4, startCol).Value = "빨강"
    ws.Cells(topRow + UBound(labels) + 4, startCol + 1).Value = "상한 초과"
    ws.Cells(topRow + UBound(labels) + 4, startCol).Interior.Color = RGB(255, 210, 210)
    ws.Cells(topRow + UBound(labels) + 5, startCol).Value = "파랑"
    ws.Cells(topRow + UBound(labels) + 5, startCol + 1).Value = "하한 미만"
    ws.Cells(topRow + UBound(labels) + 5, startCol).Interior.Color = RGB(210, 225, 255)
End Sub


Private Function OD_CollectRunInputs(ByRef srcWs As Worksheet, _
                                     ByRef dataArea As Range, _
                                     ByRef analysisCol As Long, _
                                     ByRef varName As String, _
                                     ByRef criteria As String, _
                                     ByRef method As String, _
                                     ByRef memo As String) As Boolean
    Dim selectedRange As Range
    If Not Toolkit_GetRangeFromUserForm("이상치 탐지 - 데이터 범위", _
        "헤더 행을 포함한 분석 범위를 선택하세요. 시트에서 Ctrl+Shift+방향키, Shift+방향키, 마우스 드래그를 모두 사용할 수 있습니다.", _
        selectedRange) Then Exit Function
    If selectedRange Is Nothing Then Exit Function

    If selectedRange.Rows.Count < 2 Then
        MsgBox "헤더 1행과 데이터 1행 이상을 포함해서 선택하세요.", vbExclamation, "이상치 탐지"
        Exit Function
    End If

    Set dataArea = selectedRange
    Set srcWs = dataArea.Worksheet

    Dim headerList As String
    Dim headerOptions As String
    Dim defaultIndex As Long: defaultIndex = 0
    Dim c As Long
    For c = 1 To dataArea.Columns.Count
        Dim headerText As String
        headerText = Trim$(CStr(dataArea.Cells(1, c).Value))
        If Len(headerText) = 0 Then headerText = "열" & c
        headerList = headerList & c & ". " & headerText & vbCrLf
        If Len(headerOptions) > 0 Then headerOptions = headerOptions & vbLf
        headerOptions = headerOptions & c & vbTab & headerText
        If defaultIndex = 0 And OD_ColumnHasNumericValue(dataArea, c) Then defaultIndex = c
    Next c
    If defaultIndex = 0 Then defaultIndex = 1

    Dim varInput As String
    varInput = Toolkit_ButtonMenu("이상치 탐지 - 분석 변수", _
                                  "이상치를 탐지할 분석 변수를 선택하세요.", _
                                  headerOptions, CStr(defaultIndex))
    If Len(Trim$(varInput)) = 0 Then Exit Function

    Dim varIndex As Long
    If IsNumeric(varInput) Then
        varIndex = CLng(varInput)
        If varIndex < 1 Or varIndex > dataArea.Columns.Count Then
            MsgBox "분석 변수 번호가 범위를 벗어났습니다.", vbExclamation, "이상치 탐지"
            Exit Function
        End If
    Else
        For c = 1 To dataArea.Columns.Count
            If Trim$(CStr(dataArea.Cells(1, c).Value)) = Trim$(varInput) Then
                varIndex = c
                Exit For
            End If
        Next c
        If varIndex = 0 Then
            MsgBox "입력한 변수명을 선택 범위의 첫 행에서 찾을 수 없습니다.", vbExclamation, "이상치 탐지"
            Exit Function
        End If
    End If

    analysisCol = dataArea.Columns(varIndex).Column
    varName = Trim$(CStr(dataArea.Cells(1, varIndex).Value))
    If Len(varName) = 0 Then varName = "열" & varIndex

    Dim criteriaInput As String
    criteriaInput = Toolkit_ButtonMenu("이상치 탐지 - 기준", _
                                       "이상치 기준을 선택하세요.", _
                                       "1" & vbTab & OD_IQR15 & vbLf & _
                                       "2" & vbTab & OD_IQR20 & vbLf & _
                                       "3" & vbTab & OD_P95 & vbLf & _
                                       "4" & vbTab & OD_P99, _
                                       "1")
    If Len(Trim$(criteriaInput)) = 0 Then Exit Function

    Select Case Trim$(criteriaInput)
        Case "1", OD_IQR15: criteria = OD_IQR15
        Case "2", OD_IQR20: criteria = OD_IQR20
        Case "3", OD_P95: criteria = OD_P95
        Case "4", OD_P99: criteria = OD_P99
        Case Else
            MsgBox "이상치 기준은 1~4 중 하나를 입력하세요.", vbExclamation, "이상치 탐지"
            Exit Function
    End Select

    Dim methodInput As String
    methodInput = Toolkit_ButtonMenu("이상치 탐지 - 대체 방식", _
                                     "이상치 대체 방식을 선택하세요.", _
                                     "1" & vbTab & OD_CAP & vbLf & _
                                     "2" & vbTab & OD_MEDIAN & vbLf & _
                                     "3" & vbTab & OD_MEAN, _
                                     "1")
    If Len(Trim$(methodInput)) = 0 Then Exit Function

    Select Case Trim$(methodInput)
        Case "1", OD_CAP: method = OD_CAP
        Case "2", OD_MEDIAN: method = OD_MEDIAN
        Case "3", OD_MEAN: method = OD_MEAN
        Case Else
            MsgBox "대체 방식은 1~3 중 하나를 입력하세요.", vbExclamation, "이상치 탐지"
            Exit Function
    End Select

    memo = InputBox( _
        "⑤ 결과 시트명에 붙일 메모를 입력하세요." & vbCrLf & vbCrLf & _
        "비워두어도 됩니다. 예: 1차검토, 매출, 쿼터A", _
        "이상치 탐지 - 메모", "")

    OD_CollectRunInputs = True
End Function

Private Function OD_ColumnHasNumericValue(ByVal dataArea As Range, ByVal relativeCol As Long) As Boolean
    Dim r As Long
    For r = 2 To dataArea.Rows.Count
        If Not dataArea.Worksheet.Rows(dataArea.Row + r - 1).Hidden Then
            Dim v As Variant
            v = dataArea.Cells(r, relativeCol).Value
            Dim numericValue As Double
            If OD_TryNumericValue(v, numericValue) Then
                OD_ColumnHasNumericValue = True
                Exit Function
            End If
        End If
    Next r
End Function
Private Function OD_EnsureConfigSheet() As Worksheet
    On Error Resume Next
    Set OD_EnsureConfigSheet = ThisWorkbook.Worksheets(OD_CFG_SHEET)
    On Error GoTo 0
    If OD_EnsureConfigSheet Is Nothing Then
        Set OD_EnsureConfigSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        OD_EnsureConfigSheet.Name = OD_CFG_SHEET
        OD_InitCfgLayout OD_EnsureConfigSheet
    End If
End Function

Private Sub OD_InitCfgLayout(ws As Worksheet)
    With ws.Range("A1")
        .Value = "이상치 탐지 및 대체 설정"
        .Font.Bold = True
        .Font.Size = 13
    End With

    Dim labels As Variant
    labels = Array("데이터 소스 시트", "분석 대상 변수", "이상치 기준", "대체 방식")
    Dim r As Long
    For r = 0 To 3
        With ws.Cells(3 + r, 1)
            .Value = labels(r)
            .Font.Bold = True
        End With
    Next r

    With ws.Cells(8, 1)
        .Value = "쿼터 메모 (결과 시트명에 포함)"
        .Font.Bold = True
    End With

    ws.Range("A10").Value = "※ 소스 시트를 변경한 경우 OutlierDetect_Setup 을 다시 실행하면 분석 변수 목록이 갱신됩니다."
    ws.Range("A11").Value = "※ 설정 완료 후 OutlierDetect_Run 을 실행하세요."
    ws.Range("A10:A11").Font.Color = RGB(130, 130, 130)

    ws.Columns("A").ColumnWidth = 36
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("E").Hidden = True   ' 분석 변수 드롭다운 소스 목록 (자동 관리)
End Sub

Private Sub OD_RefreshVarDropdown(cfg As Worksheet)
    Dim srcName As String: srcName = Trim$(cfg.Range(OD_C_SRC).Value)
    If Len(srcName) = 0 Then Exit Sub

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(srcName)
    On Error GoTo 0
    If srcWs Is Nothing Then Exit Sub

    ' E열에 소스 시트의 컬럼 헤더 목록 기록 (Data Validation 255자 제한 우회)
    cfg.Columns("E").ClearContents
    Dim colRow As Long: colRow = 0
    Dim hc As Range
    For Each hc In srcWs.UsedRange.Rows(1).Cells
        Dim hv As String: hv = Trim$(CStr(hc.Value))
        If Len(hv) > 0 Then
            colRow = colRow + 1
            cfg.Cells(colRow, 5).Value = hv
        End If
    Next hc

    If colRow = 0 Then Exit Sub

    Dim listAddr As String
    listAddr = "=" & cfg.Range(cfg.Cells(1, 5), cfg.Cells(colRow, 5)).Address(True, True, xlA1, False)
    OD_SetListVal cfg.Range(OD_C_VAR), listAddr
End Sub

Private Sub OD_SetListVal(cel As Range, formula As String)
    With cel.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formula
        .ShowError = False
    End With
End Sub

Private Sub OD_WriteHdr(ws As Worksheet, ByRef row As Long, title As String)
    With ws.Cells(row, 1)
        .Value = "[ " & title & " ]"
        .Font.Bold = True
    End With
    row = row + 1
End Sub

Private Sub OD_WriteKV(ws As Worksheet, ByRef row As Long, key As String, val As Variant)
    ws.Cells(row, 1).Value = key
    ws.Cells(row, 2).Value = val
    row = row + 1
End Sub

Private Function OD_IsInternal(name As String) As Boolean
    OD_IsInternal = (Left$(name, 1) = "_") Or (Left$(name, 3) = "이상치")
End Function

Private Function OD_SheetExists(name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    OD_SheetExists = Not (ws Is Nothing)
End Function

Private Function OD_CleanName(s As String) As String
    Dim result As String: result = s
    Dim bad As Variant
    bad = Array("\", "/", "?", "*", "[", "]", ":")
    Dim ch As Variant
    For Each ch In bad
        result = Replace(result, CStr(ch), "_")
    Next ch
    OD_CleanName = result
End Function
