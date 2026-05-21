Attribute VB_Name = "MacroToolkit"
' ============================================================
' 모듈명  : MacroToolkit
' 설  명  : 엑셀 데이터 편집을 위한 매크로 도구 모음
'           - 텍스트 형식 숫자를 실제 숫자로 일괄 변환
'           - 혼용된 한국 전화번호 형식을 표준 형식으로 정규화
' ============================================================
Option Explicit

' 현재 툴킷 버전 (외부 모듈에서도 참조 가능하도록 Public 선언)
Public Const TOOLKIT_VERSION As String = "0.8.1"

' 버전 정보를 기록할 숨김 시트 이름 (언더스코어 접두어로 내부 시트임을 표시)
Private Const VERSION_SHEET As String = "_MacroToolkit"
Private Const BIZ_API_BASE_URL As String = "https://api.odcloud.kr/api/nts-businessman/v1/status"
Private Const BIZ_APP_NAME As String = "DataEditingMacroToolkit"
Private Const BIZ_APP_SECTION As String = "BusinessApi"
Private Const BIZ_APP_KEY As String = "ServiceKey"

' ============================================================
' 프로시저 : Toolkit_About
' 설  명  : 툴킷의 버전과 주요 기능을 팝업 메시지로 안내
' ============================================================
Public Sub Toolkit_About()
    MsgBox "데이터 편집 매크로 툴킷" & vbCrLf & _
           "버전: " & TOOLKIT_VERSION & vbCrLf & vbCrLf & _
           "주요 기능:" & vbCrLf & _
           "- 텍스트 형식 숫자를 실제 숫자로 변환" & vbCrLf & _
           "- 혼용된 한국 전화번호 형식을 표준 형식으로 정규화" & vbCrLf & _
           "- SPSS 내보내기 표의 선택 배너 블록 삭제" & vbCrLf & _
           "- 사업자등록번호 병합 및 국세청 상태조회 API 연동" & vbCrLf & _
           "- 선택 범위 기반 이상치 탐지 및 대체", _
           vbInformation, "매크로 툴킷"
End Sub

' ============================================================
' 프로시저 : Toolkit_SetupWorkbook
' 설  명  : 현재 통합 문서에 버전 정보 시트를 생성하고
'           사용 가능한 매크로 목록을 기록한다.
'           처음 워크북에 툴킷을 배포할 때 한 번 실행한다.
' ============================================================
Public Sub Toolkit_SetupWorkbook()
    Dim ws As Worksheet
    ' 버전 시트가 없으면 새로 만들고, 있으면 기존 시트를 반환
    Set ws = EnsureVersionSheet()

    ' 헤더 및 버전 정보 기록
    ws.Range("A1").Value = "데이터 편집 매크로 툴킷"
    ws.Range("A2").Value = "버전"
    ws.Range("B2").Value = TOOLKIT_VERSION
    ws.Range("A3").Value = "갱신 일시"
    ws.Range("B3").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' 사용 가능한 매크로 목록 기록
    ws.Range("A5").Value = "사용 가능한 매크로"
    ws.Range("A6").Value  = "Toolkit_About"
    ws.Range("A7").Value  = "ConvertTextNumbersInSelection"
    ws.Range("A8").Value  = "ConvertTextNumbersInUsedRange"
    ws.Range("A9").Value  = "ConvertTextNumbersInColumns"
    ws.Range("A10").Value = "NormalizePhonesInSelection"
    ws.Range("A11").Value = "NormalizePhonesInColumns"
    ws.Range("A12").Value = "OutlierDetect_Setup"
    ws.Range("A13").Value = "OutlierDetect_Run"
    ws.Range("A14").Value = "SPSS_DeleteSelectedBanner"
    ws.Range("A15").Value = "BizNo_SetApiKey"
    ws.Range("A16").Value = "BizNo_MergeThreeColumns"
    ws.Range("A17").Value = "BizNo_QueryStatusForSelection"
    ws.Range("A18").Value = "OutlierDetect_ApplyToSelection"

    ' 내용에 맞게 열 너비 자동 조절
    ws.Columns("A:B").AutoFit
End Sub

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
    ConvertTextNumbersInRange ActiveSheet.UsedRange
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
    ' Application.InputBox(Type:=8): 엑셀 내장 범위 선택 대화상자
    ' 마우스 드래그 선택과 직접 주소 입력을 모두 지원한다.
    ' 취소 시 False(Boolean)를 반환하므로 오류 처리로 Nothing 여부를 판별한다.
    On Error Resume Next
    Set targetRange = Application.InputBox( _
        "변환할 범위를 마우스로 선택하거나 직접 입력하세요." & vbCrLf & _
        "예시: A:A  /  B1:B100  /  A:A,C:C", _
        "텍스트 숫자 변환", Type:=8)
    On Error GoTo 0
    If targetRange Is Nothing Then Exit Sub
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
    NormalizePhonesInRange Selection
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
    On Error Resume Next
    Set targetRange = Application.InputBox( _
        "전화번호 범위를 마우스로 선택하거나 직접 입력하세요." & vbCrLf & _
        "예시: B:B  /  C1:C100  /  B:B,D:D", _
        "전화번호 정규화", Type:=8)
    On Error GoTo 0
    If targetRange Is Nothing Then Exit Sub
    NormalizePhonesInRange targetRange
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

    On Error GoTo ErrHandler
    ' 화면 갱신 중지: 셀 업데이트마다 화면을 다시 그리지 않아 속도 향상
    Application.ScreenUpdating = False
    For Each cell In target.Cells
        If cell.HasFormula Then
            ' 수식 셀은 값을 덮어쓰면 수식이 사라지므로 건너뜀
            skippedFormulaCount = skippedFormulaCount + 1
        ElseIf TryNumberFromText(cell.Value, newValue) Then
            ' 텍스트를 숫자로 파싱 성공 시 셀 값 교체
            cell.Value = newValue
            convertedCount = convertedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox convertedCount & "개 셀이 변환되었습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다.", _
           vbInformation, "텍스트 숫자 변환"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, "텍스트 숫자 변환"
End Sub

' ============================================================
' 프로시저 : NormalizePhonesInRange (내부 전용)
' 설  명  : 주어진 Range 내 각 셀을 순회하며
'           한국 전화번호를 표준 형식으로 정규화한다.
'           셀 서식을 텍스트(@)로 변경해 앞자리 0이 유지되도록 한다.
'           수식 셀은 건너뛰고 완료 후 통계를 표시한다.
' 파라미터 : target - 정규화 대상 Range
' ============================================================
Private Sub NormalizePhonesInRange(ByVal target As Range)
    Dim cell As Range
    Dim changedCount As Long         ' 실제 변경된 셀 수
    Dim skippedFormulaCount As Long  ' 수식 셀이라 건너뛴 수
    Dim normalized As String         ' 정규화된 전화번호 문자열

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    For Each cell In target.Cells
        If cell.HasFormula Then
            skippedFormulaCount = skippedFormulaCount + 1
        ElseIf NormalizeKoreanPhone(CStr(cell.Value), normalized) Then
            ' 값이 실제로 바뀐 경우에만 셀을 업데이트
            If CStr(cell.Value) <> normalized Then
                ' 텍스트 서식으로 설정해야 "010-..."처럼 하이픈이 포함된
                ' 번호가 날짜/수식으로 오해되지 않고 앞자리 0도 보존됨
                cell.NumberFormat = "@"
                cell.Value = normalized
                changedCount = changedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox changedCount & "개 전화번호가 정규화되었습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다.", _
           vbInformation, "전화번호 정규화"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, "전화번호 정규화"
End Sub

' ============================================================
' 함  수 : TryNumberFromText (내부 전용)
' 설  명  : 셀 값이 텍스트 형식의 숫자이면 Double로 변환하여
'           valueOut에 저장하고 True를 반환한다.
'           처리하는 형식:
'             - 앞의 작은따옴표(') 제거  → '1234  =>  1234
'             - 비표시 공백(NBSP, U+00A0) 제거
'             - 천 단위 쉼표 제거          → 1,234  =>  1234
'             - 원화 기호(\, ₩) 제거       → \1000  =>  1000
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
    ' 원화 기호 두 종류 모두 제거 (역슬래시형 \, 유니코드 ₩)
    textValue = Replace(textValue, "\", "")
    textValue = Replace(textValue, "₩", "")

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
'           한국 전화번호 규칙에 따라 표준 형식(XXX-XXXX-XXXX)으로
'           정규화하여 valueOut에 저장하고 True를 반환한다.
'           지원 형식:
'             - 서울 지역번호 (02): 02-XXX-XXXX / 02-XXXX-XXXX
'             - 휴대전화 (01X): 01X-XXX-XXXX / 01X-XXXX-XXXX
'             - 기타 지역번호 (0XX): 0XX-XXX-XXXX / 0XX-XXXX-XXXX
'             - 국제번호 +82 또는 0082 → 0으로 치환 후 처리
' 파라미터 : valueIn  - 원본 전화번호 문자열
'            valueOut - 정규화된 전화번호 (ByRef, 호출자에게 반환)
' 반환값  : 정규화 성공 시 True, 인식 불가 형식이면 False
' ============================================================
Private Function NormalizeKoreanPhone(ByVal valueIn As String, ByRef valueOut As String) As Boolean
    Dim source As String
    Dim digits As String  ' 숫자만 추출한 문자열

    source = Trim$(valueIn)
    If Len(source) = 0 Then Exit Function

    ' 괄호, 공백, 하이픈 등 비숫자 문자를 모두 제거하여 순수 숫자열 획득
    ' 예: 02)123-4567, (031)1234-5678, +82-10-1234-5678
    source = Replace(source, "(", "")
    source = Replace(source, ")", "")
    digits = DigitsOnly(source)
    If Len(digits) = 0 Then Exit Function

    ' 국제전화 접두어(0082 또는 82)를 국내 형식(0)으로 변환
    If Left$(digits, 4) = "0082" Then digits = "0" & Mid$(digits, 5)
    If Left$(digits, 2) = "82" Then digits = "0" & Mid$(digits, 3)

    If Left$(digits, 2) = "02" Then
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

    NormalizeKoreanPhone = True
End Function

' ============================================================
' 함  수 : DigitsOnly (내부 전용)
' 설  명  : 입력 문자열에서 0~9 사이의 문자만 추출하여 반환한다.
'           전화번호의 하이픈, 괄호, 공백 등을 제거할 때 사용한다.
' 파라미터 : source - 원본 문자열
' 반환값  : 숫자 문자만으로 이루어진 문자열
' ============================================================
Private Function DigitsOnly(ByVal source As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(source)
        ch = Mid$(source, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

' ============================================================
' 사업자등록번호 정리 및 국세청 상태조회 API
' ============================================================
Public Sub BizNo_SetApiKey()
    Dim currentKey As String
    Dim newKey As Variant

    currentKey = Biz_GetApiKey()
    newKey = Application.InputBox( _
        "공공데이터포털 일반 인증키를 입력하세요." & vbCrLf & _
        "Encoding 키와 Decoding 키 모두 사용할 수 있습니다.", _
        "사업자등록 API 인증키 설정", currentKey, Type:=2)

    If VarType(newKey) = vbBoolean Then Exit Sub
    If Len(Trim$(CStr(newKey))) = 0 Then
        MsgBox "인증키가 비어 있어 저장하지 않았습니다.", vbExclamation, "사업자등록 API"
        Exit Sub
    End If

    SaveSetting BIZ_APP_NAME, BIZ_APP_SECTION, BIZ_APP_KEY, Trim$(CStr(newKey))
    MsgBox "인증키를 저장했습니다.", vbInformation, "사업자등록 API"
End Sub

Public Sub BizNo_MergeThreeColumns()
    Dim targetRange As Range
    Dim ws As Worksheet
    Dim outputCol As Long
    Dim r As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim dataStartRow As Long
    Dim formatted As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "사업자등록번호가 나뉜 3개 열 범위를 먼저 선택하세요.", vbExclamation, "사업자등록번호 병합"
        Exit Sub
    End If

    Set targetRange = Selection
    If targetRange.Columns.Count < 3 Then
        MsgBox "3개 열로 나뉜 범위를 선택해야 합니다." & vbCrLf & _
               "예: 123 / 45 / 67890 형태의 세 열", vbExclamation, "사업자등록번호 병합"
        Exit Sub
    End If

    Set ws = targetRange.Worksheet
    firstRow = targetRange.Row
    lastRow = targetRange.Row + targetRange.Rows.Count - 1
    outputCol = targetRange.Column + targetRange.Columns.Count
    dataStartRow = firstRow
    If Len(DigitsOnly(CStr(ws.Cells(firstRow, targetRange.Column).Value))) <> 3 Or _
       Len(DigitsOnly(CStr(ws.Cells(firstRow, targetRange.Column + 1).Value))) <> 2 Or _
       Len(DigitsOnly(CStr(ws.Cells(firstRow, targetRange.Column + 2).Value))) <> 5 Then
        dataStartRow = firstRow + 1
    End If

    ws.Columns(outputCol).Insert Shift:=xlToRight
    If dataStartRow > firstRow Then
        ws.Cells(firstRow, outputCol).Value = "사업자등록번호"
        ws.Cells(firstRow, outputCol).NumberFormat = "@"
    End If

    For r = dataStartRow To lastRow
        formatted = Biz_FormatNoFromParts( _
            CStr(ws.Cells(r, targetRange.Column).Value), _
            CStr(ws.Cells(r, targetRange.Column + 1).Value), _
            CStr(ws.Cells(r, targetRange.Column + 2).Value))
        ws.Cells(r, outputCol).NumberFormat = "@"
        ws.Cells(r, outputCol).Value = formatted
    Next r

    ws.Columns(outputCol).AutoFit
    MsgBox "사업자등록번호 병합을 완료했습니다." & vbCrLf & _
           "형식: 000-00-00000", vbInformation, "사업자등록번호 병합"
End Sub

Public Sub BizNo_QueryStatusForSelection()
    Dim targetRange As Range
    Dim apiKey As String
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim dataStartRow As Long
    Dim headerRow As Long
    Dim outputCol As Long
    Dim rowIndex As Long
    Dim requestCount As Long
    Dim successCount As Long
    Dim failCount As Long
    Dim batchNos(1 To 100) As String
    Dim batchRows(1 To 100) As Long
    Dim batchCount As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "조회할 사업자등록번호 범위를 먼저 선택하세요.", vbExclamation, "사업자등록 상태조회"
        Exit Sub
    End If

    Set targetRange = Selection.Areas(1)
    If targetRange.Columns.Count > 1 Then
        Set targetRange = targetRange.Columns(1)
    End If

    apiKey = Biz_GetApiKey()
    If Len(apiKey) = 0 Then
        MsgBox "저장된 인증키가 없습니다. 먼저 [인증키 설정]을 실행하세요.", vbExclamation, "사업자등록 상태조회"
        Exit Sub
    End If

    Set ws = targetRange.Worksheet
    firstRow = targetRange.Row
    lastRow = targetRange.Row + targetRange.Rows.Count - 1
    dataStartRow = firstRow
    headerRow = firstRow
    If Len(DigitsOnly(CStr(ws.Cells(firstRow, targetRange.Column).Value))) = 10 Then
        If firstRow > 1 Then
            headerRow = firstRow - 1
        Else
            headerRow = 0
        End If
    Else
        dataStartRow = firstRow + 1
    End If

    If dataStartRow > lastRow Then
        MsgBox "조회할 사업자등록번호가 없습니다.", vbExclamation, "사업자등록 상태조회"
        Exit Sub
    End If

    outputCol = targetRange.Column + targetRange.Columns.Count
    ws.Columns(outputCol).Resize(, Biz_OutputFieldCount()).Insert Shift:=xlToRight
    If headerRow > 0 Then Biz_WriteStatusHeaders ws, headerRow, outputCol

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    For rowIndex = dataStartRow To lastRow
        Dim bizNo As String
        bizNo = DigitsOnly(CStr(ws.Cells(rowIndex, targetRange.Column).Value))
        If Len(bizNo) = 10 Then
            batchCount = batchCount + 1
            batchNos(batchCount) = bizNo
            batchRows(batchCount) = rowIndex
            requestCount = requestCount + 1

            If batchCount = 100 Then
                Biz_QueryStatusBatch ws, outputCol, batchNos, batchRows, batchCount, apiKey, successCount, failCount
                batchCount = 0
            End If
        ElseIf Len(Trim$(CStr(ws.Cells(rowIndex, targetRange.Column).Value))) > 0 Then
            failCount = failCount + 1
            Biz_WriteStatusError ws, rowIndex, outputCol, CStr(ws.Cells(rowIndex, targetRange.Column).Value), "사업자등록번호는 숫자 10자리여야 합니다."
        End If
    Next rowIndex

    If batchCount > 0 Then
        Biz_QueryStatusBatch ws, outputCol, batchNos, batchRows, batchCount, apiKey, successCount, failCount
    End If

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(firstRow, outputCol + Biz_OutputFieldCount() - 1)).EntireColumn.AutoFit
    Application.ScreenUpdating = True

    MsgBox "사업자등록 상태조회를 완료했습니다." & vbCrLf & _
           "요청: " & requestCount & "건" & vbCrLf & _
           "성공: " & successCount & "건" & vbCrLf & _
           "확인 필요: " & failCount & "건", vbInformation, "사업자등록 상태조회"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "조회 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical, "사업자등록 상태조회"
End Sub

Private Function Biz_GetApiKey() As String
    Biz_GetApiKey = GetSetting(BIZ_APP_NAME, BIZ_APP_SECTION, BIZ_APP_KEY, "")
End Function

Private Function Biz_FormatNoFromParts(ByVal part1 As String, ByVal part2 As String, ByVal part3 As String) As String
    Dim p1 As String, p2 As String, p3 As String
    p1 = Right$("000" & DigitsOnly(part1), 3)
    p2 = Right$("00" & DigitsOnly(part2), 2)
    p3 = Right$("00000" & DigitsOnly(part3), 5)
    Biz_FormatNoFromParts = p1 & "-" & p2 & "-" & p3
End Function

Private Function Biz_FormatNo(ByVal valueIn As String) As String
    Dim digits As String
    digits = DigitsOnly(valueIn)
    If Len(digits) = 10 Then
        Biz_FormatNo = Left$(digits, 3) & "-" & Mid$(digits, 4, 2) & "-" & Right$(digits, 5)
    Else
        Biz_FormatNo = valueIn
    End If
End Function

Private Function Biz_OutputFieldCount() As Long
    Biz_OutputFieldCount = 13
End Function

Private Sub Biz_WriteStatusHeaders(ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long)
    Dim headers As Variant
    headers = Array( _
        "status_code(응답상태코드)", _
        "requested_b_no(요청 사업자등록번호)", _
        "b_no(사업자등록번호)", _
        "b_stt(납세자상태)", _
        "b_stt_cd(납세자상태코드)", _
        "tax_type(과세유형)", _
        "tax_type_cd(과세유형코드)", _
        "end_dt(폐업일자)", _
        "utcc_yn(단위과세전환폐업여부)", _
        "tax_type_change_dt(과세유형전환일자)", _
        "invoice_apply_dt(세금계산서적용일자)", _
        "rbf_tax_type(직전과세유형)", _
        "rbf_tax_type_cd(직전과세유형코드)")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowIndex, colIndex + i).Value = headers(i)
    Next i
End Sub

Private Sub Biz_QueryStatusBatch(ws As Worksheet, ByVal outputCol As Long, ByRef batchNos() As String, ByRef batchRows() As Long, ByVal batchCount As Long, ByVal apiKey As String, ByRef successCount As Long, ByRef failCount As Long)
    Dim responseText As String
    Dim statusCode As String
    Dim dataObjects As Collection
    Dim i As Long

    responseText = Biz_PostStatus(batchNos, batchCount, apiKey)
    statusCode = Biz_JsonString(responseText, "status_code")
    If Len(statusCode) = 0 Then statusCode = "UNKNOWN"

    Set dataObjects = Biz_ExtractDataObjects(responseText)
    For i = 1 To batchCount
        If i <= dataObjects.Count Then
            Biz_WriteStatusRow ws, batchRows(i), outputCol, statusCode, batchNos(i), CStr(dataObjects(i))
            successCount = successCount + 1
        Else
            failCount = failCount + 1
            Biz_WriteStatusError ws, batchRows(i), outputCol, batchNos(i), "응답 데이터가 없습니다. status_code=" & statusCode
        End If
    Next i
End Sub

Private Function Biz_PostStatus(ByRef batchNos() As String, ByVal batchCount As Long, ByVal apiKey As String) As String
    Dim http As Object
    Dim url As String
    Dim body As String
    Dim i As Long

    body = "{""b_no"":["
    For i = 1 To batchCount
        If i > 1 Then body = body & ","
        body = body & """" & batchNos(i) & """"
    Next i
    body = body & "]}"

    url = BIZ_API_BASE_URL & "?serviceKey=" & Biz_ServiceKeyForUrl(apiKey)
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Accept", "application/json"
    http.send body

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 701, "Biz_PostStatus", "HTTP " & http.Status & ": " & http.responseText
    End If
    Biz_PostStatus = CStr(http.responseText)
End Function

Private Function Biz_ServiceKeyForUrl(ByVal apiKey As String) As String
    If InStr(1, apiKey, "%", vbBinaryCompare) > 0 Then
        Biz_ServiceKeyForUrl = apiKey
    Else
        Biz_ServiceKeyForUrl = Replace(Replace(Replace(apiKey, "+", "%2B"), "/", "%2F"), "=", "%3D")
    End If
End Function

Private Sub Biz_WriteStatusRow(ws As Worksheet, ByVal rowIndex As Long, ByVal outputCol As Long, ByVal statusCode As String, ByVal requestedNo As String, ByVal dataObject As String)
    ws.Cells(rowIndex, outputCol).Value = statusCode
    ws.Cells(rowIndex, outputCol + 1).NumberFormat = "@"
    ws.Cells(rowIndex, outputCol + 1).Value = Biz_FormatNo(requestedNo)
    ws.Cells(rowIndex, outputCol + 2).NumberFormat = "@"
    ws.Cells(rowIndex, outputCol + 2).Value = Biz_FormatNo(Biz_JsonString(dataObject, "b_no"))
    ws.Cells(rowIndex, outputCol + 3).Value = Biz_JsonString(dataObject, "b_stt")
    ws.Cells(rowIndex, outputCol + 4).Value = Biz_JsonString(dataObject, "b_stt_cd")
    ws.Cells(rowIndex, outputCol + 5).Value = Biz_JsonString(dataObject, "tax_type")
    ws.Cells(rowIndex, outputCol + 6).Value = Biz_JsonString(dataObject, "tax_type_cd")
    ws.Cells(rowIndex, outputCol + 7).Value = Biz_JsonString(dataObject, "end_dt")
    ws.Cells(rowIndex, outputCol + 8).Value = Biz_JsonString(dataObject, "utcc_yn")
    ws.Cells(rowIndex, outputCol + 9).Value = Biz_JsonString(dataObject, "tax_type_change_dt")
    ws.Cells(rowIndex, outputCol + 10).Value = Biz_JsonString(dataObject, "invoice_apply_dt")
    ws.Cells(rowIndex, outputCol + 11).Value = Biz_JsonString(dataObject, "rbf_tax_type")
    ws.Cells(rowIndex, outputCol + 12).Value = Biz_JsonString(dataObject, "rbf_tax_type_cd")
End Sub

Private Sub Biz_WriteStatusError(ws As Worksheet, ByVal rowIndex As Long, ByVal outputCol As Long, ByVal requestedNo As String, ByVal message As String)
    ws.Cells(rowIndex, outputCol).Value = "ERROR"
    ws.Cells(rowIndex, outputCol + 1).NumberFormat = "@"
    ws.Cells(rowIndex, outputCol + 1).Value = Biz_FormatNo(requestedNo)
    ws.Cells(rowIndex, outputCol + 3).Value = message
End Sub

Private Function Biz_ExtractDataObjects(ByVal jsonText As String) As Collection
    Dim result As New Collection
    Dim dataPos As Long
    Dim arrStart As Long
    Dim i As Long
    Dim depth As Long
    Dim objStart As Long
    Dim inString As Boolean
    Dim ch As String
    Dim prevCh As String

    dataPos = InStr(1, jsonText, """data""", vbTextCompare)
    If dataPos = 0 Then
        Set Biz_ExtractDataObjects = result
        Exit Function
    End If

    arrStart = InStr(dataPos, jsonText, "[")
    If arrStart = 0 Then
        Set Biz_ExtractDataObjects = result
        Exit Function
    End If

    For i = arrStart + 1 To Len(jsonText)
        ch = Mid$(jsonText, i, 1)
        If ch = """" And prevCh <> "\" Then inString = Not inString
        If Not inString Then
            If ch = "{" Then
                If depth = 0 Then objStart = i
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1
                If depth = 0 And objStart > 0 Then result.Add Mid$(jsonText, objStart, i - objStart + 1)
            ElseIf ch = "]" And depth = 0 Then
                Exit For
            End If
        End If
        prevCh = ch
    Next i

    Set Biz_ExtractDataObjects = result
End Function

Private Function Biz_JsonString(ByVal jsonText As String, ByVal key As String) As String
    Dim re As Object
    Dim matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = """" & key & """" & "\s*:\s*""(([^""\\]|\\.)*)"""
    Set matches = re.Execute(jsonText)
    If matches.Count > 0 Then Biz_JsonString = Biz_JsonUnescape(CStr(matches(0).SubMatches(0)))
End Function

Private Function Biz_JsonUnescape(ByVal valueIn As String) As String
    Biz_JsonUnescape = Replace(Replace(Replace(Replace(valueIn, "\""", """"), "\\", "\"), "\/", "/"), "\n", vbLf)
End Function

' ============================================================
' 함  수 : EnsureVersionSheet (내부 전용)
' 설  명  : VERSION_SHEET 이름의 시트가 없으면 새로 만들고
'           맨 앞에 배치한 뒤 반환한다. 이미 있으면 기존 시트를 반환한다.
' 반환값  : 버전 정보 기록용 Worksheet 객체
' ============================================================
Private Function EnsureVersionSheet() As Worksheet
    ' 오류 무시 후 시트 이름으로 참조 시도 (없으면 Nothing 상태 유지)
    On Error Resume Next
    Set EnsureVersionSheet = ThisWorkbook.Worksheets(VERSION_SHEET)
    On Error GoTo 0

    If EnsureVersionSheet Is Nothing Then
        ' 시트가 없으면 첫 번째 시트 앞에 새로 삽입
        Set EnsureVersionSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        EnsureVersionSheet.Name = VERSION_SHEET
    End If
End Function

' ============================================================
' ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'  이상치 탐지 및 대체 (OutlierDetector)
'  · OutlierDetect_Setup : 설정 시트 생성/갱신
'  · OutlierDetect_Run   : 이상치 탐지 및 결과 시트 출력
' ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' ============================================================

Private Const OD_CFG_SHEET As String = "_이상치탐지_설정"
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
    Dim cfg As Worksheet
    Set cfg = OD_EnsureConfigSheet()

    ' ── 소스 시트 목록 ──────────────────────────────────
    Dim ws As Worksheet
    Dim sheetList As String
    For Each ws In ThisWorkbook.Worksheets
        If Not OD_IsInternal(ws.Name) Then
            If Len(sheetList) > 0 Then sheetList = sheetList & ","
            sheetList = sheetList & ws.Name
        End If
    Next ws

    If Len(sheetList) = 0 Then
        MsgBox "분석 대상 시트가 없습니다.", vbExclamation, "이상치 탐지"
        Exit Sub
    End If

    OD_SetListVal cfg.Range(OD_C_SRC), sheetList

    ' 소스 시트 기본값 (비어있거나 이미 삭제된 경우)
    Dim curSrc As String: curSrc = Trim$(cfg.Range(OD_C_SRC).Value)
    If Len(curSrc) = 0 Or Not OD_SheetExists(curSrc) Then
        If Not OD_IsInternal(ActiveSheet.Name) Then
            cfg.Range(OD_C_SRC).Value = ActiveSheet.Name
        Else
            cfg.Range(OD_C_SRC).Value = Split(sheetList, ",")(0)
        End If
    End If

    ' ── 고정 드롭다운 ────────────────────────────────────
    OD_SetListVal cfg.Range(OD_C_CRIT), OD_IQR15 & "," & OD_IQR20 & "," & OD_P95 & "," & OD_P99
    OD_SetListVal cfg.Range(OD_C_METH), OD_CAP & "," & OD_MEDIAN & "," & OD_MEAN

    If Len(cfg.Range(OD_C_CRIT).Value) = 0 Then cfg.Range(OD_C_CRIT).Value = OD_IQR15
    If Len(cfg.Range(OD_C_METH).Value) = 0 Then cfg.Range(OD_C_METH).Value = OD_CAP

    ' ── 분석 변수 드롭다운 갱신 ──────────────────────────
    OD_RefreshVarDropdown cfg

    cfg.Activate
    cfg.Range(OD_C_SRC).Select

    MsgBox "설정 시트가 준비되었습니다." & vbCrLf & vbCrLf & _
           "① B3~B6 항목을 설정하세요." & vbCrLf & _
           "② 소스 시트를 변경한 경우 이 매크로를 다시 실행하세요." & vbCrLf & _
           "③ 설정 완료 후 OutlierDetect_Run 을 실행하세요.", _
           vbInformation, "이상치 탐지 설정"
End Sub

' ============================================================
' 프로시저 : OutlierDetect_Run
' 설    명  : 설정 시트 값을 읽어 이상치 탐지 및 대체를 수행하고
'             결과를 새 시트에 출력한다.
'             원본 시트는 수정하지 않는다.
' ============================================================
Public Sub OutlierDetect_Run()
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
    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction

    Dim q1     As Double: q1     = wf.Percentile(values, 0.25)
    Dim q3     As Double: q3     = wf.Percentile(values, 0.75)
    Dim iqrVal As Double: iqrVal = q3 - q1
    Dim medV   As Double: medV   = wf.Median(values)
    Dim meanV  As Double: meanV  = wf.Average(values)
    Dim minV   As Double: minV   = wf.Min(values)
    Dim maxV   As Double: maxV   = wf.Max(values)

    Dim lo As Double, hi As Double
    Select Case criteria
        Case OD_IQR15: lo = q1 - 1.5 * iqrVal:                    hi = q3 + 1.5 * iqrVal
        Case OD_IQR20: lo = q1 - 2.0 * iqrVal:                    hi = q3 + 2.0 * iqrVal
        Case OD_P95:   lo = wf.Percentile(values, 0.05):           hi = wf.Percentile(values, 0.95)
        Case OD_P99:   lo = wf.Percentile(values, 0.01):           hi = wf.Percentile(values, 0.99)
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
            If method = OD_MEDIAN Then replBase = wf.Median(trimVals)
            If method = OD_MEAN   Then replBase = wf.Average(trimVals)
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

    Dim dataArea As Range
    Dim targetRange As Range
    Dim analysisRelativeCol As Long

    If TypeName(Selection) = "Range" Then
        Set dataArea = Selection.Areas(1)
    Else
        On Error Resume Next
        Set dataArea = Application.InputBox( _
            "이상치를 탐지할 데이터 범위를 선택하세요." & vbCrLf & _
            "표 전체를 선택해도 숫자 열을 자동으로 탐지합니다. 예: A1:W12001", _
            "이상치 대체 - 범위 선택", Type:=8)
        On Error GoTo 0
    End If

    If dataArea Is Nothing Then Exit Sub
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

    If firstDataRow > lastRow Then
        MsgBox "분석할 숫자 데이터가 없습니다.", vbExclamation, "이상치 대체"
        Exit Sub
    End If

    Dim criteria As String
    Dim method As String
    If Len(Trim$(criteriaOverride)) > 0 Then
        If Not OD_MapEasyCriteria(criteriaOverride, criteria) Then
            MsgBox "알 수 없는 이상치 기준입니다: " & criteriaOverride, vbExclamation, "이상치 대체"
            Exit Sub
        End If
    Else
        If Not OD_SelectEasyCriteria(criteria) Then Exit Sub
    End If

    If Len(Trim$(methodOverride)) > 0 Then
        If Not OD_MapEasyMethod(methodOverride, method) Then
            MsgBox "알 수 없는 대체 방식입니다: " & methodOverride, vbExclamation, "이상치 대체"
            Exit Sub
        End If
    Else
        If Not OD_SelectEasyMethod(method) Then Exit Sub
    End If

    If Not suppressPrompts Then
        If MsgBox("선택 범위: " & targetRange.Address(False, False) & vbCrLf & _
                  "변수: " & varName & vbCrLf & _
                  "이상치 기준: " & criteria & vbCrLf & _
                  "대체 방식: " & method & vbCrLf & vbCrLf & _
                  "원본 오른쪽에 결과 열을 삽입하고, 이상치 셀을 강조합니다." & vbCrLf & _
                  "계속하시겠습니까?", vbQuestion + vbYesNo, "이상치 대체 확인") = vbNo Then Exit Sub
    End If

    Dim rowCount As Long: rowCount = lastRow - firstDataRow + 1
    Dim values() As Double: ReDim values(1 To rowCount)
    Dim rowMap() As Long: ReDim rowMap(1 To rowCount)
    Dim numCount As Long: numCount = 0
    Dim r As Long

    For r = firstDataRow To lastRow
        If Not ws.Rows(r).Hidden Then
            Dim cellValue As Variant
            cellValue = ws.Cells(r, targetRange.Column).Value
            If Not IsError(cellValue) And Not IsEmpty(cellValue) And IsNumeric(cellValue) Then
                numCount = numCount + 1
                values(numCount) = CDbl(cellValue)
                rowMap(numCount) = r
            End If
        End If
    Next r

    If numCount = 0 Then
        MsgBox "선택 범위에서 유효한 숫자값을 찾지 못했습니다.", vbExclamation, "이상치 대체"
        Exit Sub
    End If

    ReDim Preserve values(1 To numCount)
    ReDim Preserve rowMap(1 To numCount)

    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction
    Dim q1 As Double: q1 = wf.Percentile(values, 0.25)
    Dim q3 As Double: q3 = wf.Percentile(values, 0.75)
    Dim iqrVal As Double: iqrVal = q3 - q1
    Dim medV As Double: medV = wf.Median(values)
    Dim meanV As Double: meanV = wf.Average(values)
    Dim minV As Double: minV = wf.Min(values)
    Dim maxV As Double: maxV = wf.Max(values)
    Dim lo As Double, hi As Double

    Select Case criteria
        Case OD_IQR15: lo = q1 - 1.5 * iqrVal: hi = q3 + 1.5 * iqrVal
        Case OD_IQR20: lo = q1 - 2# * iqrVal: hi = q3 + 2# * iqrVal
        Case OD_P95: lo = wf.Percentile(values, 0.05): hi = wf.Percentile(values, 0.95)
        Case OD_P99: lo = wf.Percentile(values, 0.01): hi = wf.Percentile(values, 0.99)
    End Select

    Dim replBase As Double
    If method = OD_MEDIAN Or method = OD_MEAN Then
        replBase = OD_ReplacementBase(values, numCount, lo, hi, method, medV, meanV)
    End If

    Dim outlierCount As Long: outlierCount = 0
    Dim outputCol As Long: outputCol = targetRange.Column + 1

    Application.ScreenUpdating = False

    ws.Columns(outputCol).Insert Shift:=xlToRight
    If headerRow > 0 Then
        ws.Cells(headerRow, outputCol).Value = varName & "_이상치대체"
        ws.Cells(headerRow, outputCol).Font.Bold = True
        ws.Cells(headerRow, outputCol).Interior.Color = RGB(255, 235, 185)
    End If

    Dim sourceDataRange As Range
    Dim resultDataRange As Range
    Set sourceDataRange = ws.Range(ws.Cells(firstDataRow, targetRange.Column), ws.Cells(lastRow, targetRange.Column))
    Set resultDataRange = ws.Range(ws.Cells(firstDataRow, outputCol), ws.Cells(lastRow, outputCol))
    resultDataRange.NumberFormat = sourceDataRange.NumberFormat
    resultDataRange.Value = sourceDataRange.Value

    Dim i As Long
    For i = 1 To numCount
        Dim v As Double: v = values(i)
        Dim outR As Long: outR = rowMap(i)
        If v < lo Or v > hi Then
            outlierCount = outlierCount + 1
            ws.Cells(outR, targetRange.Column).Interior.Color = RGB(255, 220, 220)
            ws.Cells(outR, targetRange.Column).Font.Color = RGB(160, 0, 0)
            ws.Cells(outR, outputCol).Interior.Color = RGB(255, 245, 190)
            ws.Cells(outR, outputCol).Font.Bold = True
            If method = OD_CAP Then
                If v < lo Then ws.Cells(outR, outputCol).Value = lo Else ws.Cells(outR, outputCol).Value = hi
            ElseIf method = OD_REVIEW Then
                ws.Cells(outR, outputCol).AddComment "이상치 검토 필요: " & criteria & " 기준 범위 [" & Format(lo, "0.####") & ", " & Format(hi, "0.####") & "]"
            Else
                ws.Cells(outR, outputCol).Value = replBase
            End If
        End If
    Next i

    OD_ApplyOutlierConditionalFormatting sourceDataRange, lo, hi
    OD_ApplyOutlierConditionalFormatting resultDataRange, lo, hi
    OD_WriteEasyStats ws, headerRow, firstDataRow, outputCol + 2, _
                      varName, criteria, method, numCount, outlierCount, _
                      minV, maxV, meanV, medV, q1, q3, iqrVal, lo, hi, replBase

    Dim displayTopRow As Long
    If headerRow > 0 Then displayTopRow = headerRow Else displayTopRow = firstDataRow
    ws.Range(ws.Cells(displayTopRow, outputCol), ws.Cells(lastRow, outputCol)).Columns.AutoFit
    ws.Range(ws.Cells(displayTopRow, outputCol + 2), ws.Cells(displayTopRow + 24, outputCol + 3)).Columns.AutoFit
    Application.ScreenUpdating = True

    If Not suppressPrompts Then
        MsgBox "이상치 대체 열을 만들었습니다." & vbCrLf & _
               "유효값: " & numCount & "개" & vbCrLf & _
               "이상치: " & outlierCount & "개" & vbCrLf & _
               "결과 열: " & ws.Cells(1, outputCol).EntireColumn.Address(False, False), _
               vbInformation, "이상치 대체 완료"
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

Private Function OD_SelectEasyCriteria(ByRef criteria As String) As Boolean
    Dim criteriaInput As String
    criteriaInput = InputBox( _
        "이상치 기준을 선택하세요." & vbCrLf & vbCrLf & _
        "1. " & OD_IQR15 & " : 일반적인 윈저라이징 기준" & vbCrLf & _
        "2. " & OD_IQR20 & " : 조금 더 완화된 IQR 기준" & vbCrLf & _
        "3. " & OD_P95 & " : 하위 5% / 상위 5% 기준" & vbCrLf & _
        "4. " & OD_P99 & " : 하위 1% / 상위 1% 기준", _
        "이상치 대체 - 기준 선택", "1")
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

    Dim prompt As String
    prompt = "선택 범위에서 숫자형 열을 찾았습니다." & vbCrLf & _
             "이상치를 탐지할 열 번호를 선택하세요." & vbCrLf & vbCrLf

    For c = 1 To numericCols.Count
        Dim listItem As Variant
        listItem = numericCols(c)
        prompt = prompt & c & ". " & CStr(listItem(1)) & _
                 " (유효값 " & CLng(listItem(2)) & _
                 "개, IQR 이상치 " & CLng(listItem(3)) & "개)" & vbCrLf
    Next c

    Dim userInput As String
    userInput = InputBox(prompt, "이상치 대체 - 숫자 열 선택", "1")
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

Private Function OD_CountNumericValues(ByVal ws As Worksheet, _
                                       ByVal absoluteCol As Long, _
                                       ByVal firstDataRow As Long, _
                                       ByVal lastRow As Long) As Long
    Dim r As Long
    For r = firstDataRow To lastRow
        If Not ws.Rows(r).Hidden Then
            Dim v As Variant
            v = ws.Cells(r, absoluteCol).Value
            If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
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
            If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
                idx = idx + 1
                values(idx) = CDbl(v)
            End If
        End If
    Next r

    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    Dim q1 As Double: q1 = wf.Percentile(values, 0.25)
    Dim q3 As Double: q3 = wf.Percentile(values, 0.75)
    Dim lo As Double: lo = q1 - 1.5 * (q3 - q1)
    Dim hi As Double: hi = q3 + 1.5 * (q3 - q1)

    For idx = 1 To countN
        If values(idx) < lo Or values(idx) > hi Then OD_CountIqrOutliers = OD_CountIqrOutliers + 1
    Next idx
End Function

Private Function OD_SelectEasyMethod(ByRef method As String) As Boolean
    Dim methodInput As String
    methodInput = InputBox( _
        "이상치 대체 방식을 선택하세요." & vbCrLf & vbCrLf & _
        "1. " & OD_CAP & " : 하한보다 작으면 하한, 상한보다 크면 상한으로 대체" & vbCrLf & _
        "2. " & OD_MEDIAN & " : 정상 범위 데이터의 중앙값으로 대체" & vbCrLf & _
        "3. " & OD_MEAN & " : 정상 범위 데이터의 평균으로 대체" & vbCrLf & _
        "4. " & OD_REVIEW & " : 옆 열에 원본을 복사하고 이상치만 표시", _
        "이상치 대체 - 방식 선택", "1")
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
        OD_ReplacementBase = Application.WorksheetFunction.Median(trimVals)
    Else
        OD_ReplacementBase = Application.WorksheetFunction.Average(trimVals)
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
    On Error Resume Next
    Set selectedRange = Application.InputBox( _
        "① 헤더 행을 포함한 분석 범위를 선택하세요." & vbCrLf & vbCrLf & _
        "예: A1:H500 처럼 첫 행에 변수명이 있는 범위" & vbCrLf & _
        "필터가 적용된 경우 보이는 행만 분석합니다.", _
        "이상치 탐지 - 데이터 범위", Type:=8)
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Function

    If selectedRange.Rows.Count < 2 Then
        MsgBox "헤더 1행과 데이터 1행 이상을 포함해서 선택하세요.", vbExclamation, "이상치 탐지"
        Exit Function
    End If

    Set dataArea = selectedRange
    Set srcWs = dataArea.Worksheet

    Dim headerList As String
    Dim defaultIndex As Long: defaultIndex = 0
    Dim c As Long
    For c = 1 To dataArea.Columns.Count
        Dim headerText As String
        headerText = Trim$(CStr(dataArea.Cells(1, c).Value))
        If Len(headerText) = 0 Then headerText = "열" & c
        headerList = headerList & c & ". " & headerText & vbCrLf
        If defaultIndex = 0 And OD_ColumnHasNumericValue(dataArea, c) Then defaultIndex = c
    Next c
    If defaultIndex = 0 Then defaultIndex = 1

    Dim varInput As String
    varInput = InputBox( _
        "② 이상치를 탐지할 분석 변수를 선택하세요." & vbCrLf & vbCrLf & _
        "아래 목록에서 번호 또는 변수명을 입력하세요." & vbCrLf & vbCrLf & headerList, _
        "이상치 탐지 - 분석 변수", CStr(defaultIndex))
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
    criteriaInput = InputBox( _
        "③ 이상치 기준을 선택하세요." & vbCrLf & vbCrLf & _
        "1. " & OD_IQR15 & vbCrLf & _
        "2. " & OD_IQR20 & vbCrLf & _
        "3. " & OD_P95 & vbCrLf & _
        "4. " & OD_P99, _
        "이상치 탐지 - 기준", "1")
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
    methodInput = InputBox( _
        "④ 이상치 대체 방식을 선택하세요." & vbCrLf & vbCrLf & _
        "1. " & OD_CAP & vbCrLf & _
        "2. " & OD_MEDIAN & vbCrLf & _
        "3. " & OD_MEAN, _
        "이상치 탐지 - 대체 방식", "1")
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
            If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
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
' ============================================================
' Ribbon callbacks
' 설  명  : .xlam 리본 탭 버튼에서 기존 매크로를 호출하기 위한 콜백
' ============================================================
Public Sub Ribbon_ToolkitAbout(control As IRibbonControl)
    Toolkit_About
End Sub

Public Sub Ribbon_ConvertTextNumbersInSelection(control As IRibbonControl)
    ConvertTextNumbersInSelection
End Sub

Public Sub Ribbon_ConvertTextNumbersInUsedRange(control As IRibbonControl)
    ConvertTextNumbersInUsedRange
End Sub

Public Sub Ribbon_ConvertTextNumbersInColumns(control As IRibbonControl)
    ConvertTextNumbersInColumns
End Sub

Public Sub Ribbon_NormalizePhonesInSelection(control As IRibbonControl)
    NormalizePhonesInSelection
End Sub

Public Sub Ribbon_NormalizePhonesInColumns(control As IRibbonControl)
    NormalizePhonesInColumns
End Sub

Public Sub Ribbon_BizNoSetApiKey(control As IRibbonControl)
    BizNo_SetApiKey
End Sub

Public Sub Ribbon_BizNoMergeThreeColumns(control As IRibbonControl)
    BizNo_MergeThreeColumns
End Sub

Public Sub Ribbon_BizNoQueryStatus(control As IRibbonControl)
    BizNo_QueryStatusForSelection
End Sub

Public Sub Ribbon_OutlierDetectSetup(control As IRibbonControl)
    OutlierDetect_Setup
End Sub

Public Sub Ribbon_OutlierDetectRun(control As IRibbonControl)
    OutlierDetect_ApplyToSelection
End Sub

Public Sub Ribbon_SPSSDeleteSelectedBanner(control As IRibbonControl)
    SPSS_DeleteSelectedBanner
End Sub






