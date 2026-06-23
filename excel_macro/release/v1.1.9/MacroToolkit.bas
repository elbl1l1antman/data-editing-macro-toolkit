Attribute VB_Name = "MacroToolkit"
' ============================================================
' 모듈명  : MacroToolkit
' 설  명  : 엑셀 데이터 편집을 위한 매크로 도구 모음
'           - 텍스트 형식 숫자를 실제 숫자로 일괄 변환
'           - 혼용된 한국 전화번호 형식을 표준 형식으로 정규화
' ============================================================
Option Explicit

' 현재 툴킷 버전 (외부 모듈에서도 참조 가능하도록 Public 선언)
Public Const TOOLKIT_VERSION As String = "1.1.9"

' 버전 정보를 기록할 숨김 시트 이름 (언더스코어 접두어로 내부 시트임을 표시)
Private Const VERSION_SHEET As String = "_MacroToolkit"
Private Const LOG_SHEET As String = "_MacroToolkit_Log"
Private Const BIZ_API_BASE_URL As String = "https://api.odcloud.kr/api/nts-businessman/v1/status"
Private Const BIZ_APP_NAME As String = "DataEditingMacroToolkit"
Private Const BIZ_APP_SECTION As String = "BusinessApi"
Private Const BIZ_APP_KEY As String = "ServiceKey"
Private Const NAVER_API_BASE_URL As String = "https://openapi.naver.com/v1/search/local.json"
Private Const NAVER_APP_NAME As String = "DataEditingMacroToolkit"
Private Const NAVER_APP_SECTION As String = "NaverLocalSearch"
Private Const NAVER_CLIENT_ID_KEY As String = "ClientId"
Private Const NAVER_CLIENT_SECRET_KEY As String = "ClientSecret"
Private Const NAVER_DONG_MASTER_SHEET As String = "_NaverLocalDongMaster"
Private Const NAVER_USER_DONG_MASTER_SHEET As String = "행정동_마스터"
Private Const NAVER_LEGAL_DONG_MASTER_SHEET As String = "_NaverLocalLegalDongMaster"
Private Const NAVER_USER_LEGAL_DONG_MASTER_SHEET As String = "법정동_마스터"
Private Const API_RESOLVE_TIMEOUT_MS As Long = 5000
Private Const API_CONNECT_TIMEOUT_MS As Long = 10000
Private Const API_SEND_TIMEOUT_MS As Long = 15000
Private Const API_RECEIVE_TIMEOUT_MS As Long = 30000

Private mToolkitOperationDepth As Long
Private mToolkitPrevScreenUpdating As Boolean
Private mToolkitPrevEnableEvents As Boolean
Private mToolkitPrevDisplayAlerts As Boolean
Private mToolkitPrevCalculation As XlCalculation
Private mToolkitPrevStatusBar As Variant

Private gProgressForm As Object

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
           "- 네이버 지역검색 API 기반 업체 리스트 생성/주소 검증" & vbCrLf & _
           "- 선택 범위 기반 Q1/Q3 이상치 확인 보조열 생성" & vbCrLf & _
           "- 표본설계 시트 생성 및 표본할당 계산", _
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
    ws.Range("A8").Value  = "ConvertTextNumbersInCurrentSheetDataRange"
    ws.Range("A9").Value  = "ConvertTextNumbersInColumns"
    ws.Range("A10").Value = "NormalizePhonesInSelection"
    ws.Range("A11").Value = "NormalizePhonesInColumns"
    ws.Range("A12").Value = "RemovePhoneDashesToCopiedColumns"
    ws.Range("A13").Value = "RemoveBizNoDashesToCopiedColumns"
    ws.Range("A14").Value = "CompanyName_CleanWithOptions"
    ws.Range("A15").Value = "Address_ClassifyAddressType"
    ws.Range("A16").Value = "OutlierDetect_Setup"
    ws.Range("A17").Value = "OutlierDetect_Run"
    ws.Range("A18").Value = "SPSS_DeleteSelectedBanner"
    ws.Range("A19").Value = "BizNo_SetApiKey"
    ws.Range("A20").Value = "BizNo_MergeThreeColumns"
    ws.Range("A21").Value = "BizNo_QueryStatusForSelection"
    ws.Range("A22").Value = "OutlierDetect_ApplyToSelection"
    ws.Range("A23").Value = "SPSS_RebuildTablesFromIndex"
    ws.Range("A24").Value = "SPSS_DeleteVerticalBannersFromIndex"
    ws.Range("A25").Value = "Toolkit_ResetAllSheetsToTop"
    ws.Range("A26").Value = "SPSS_CreateIndexForActiveSheet"
    ws.Range("A27").Value = "Toolkit_SelectA1OnAllSheets"
    ws.Range("A28").Value = "NaverLocal_SetApiCredentials"
    ws.Range("A29").Value = "NaverLocal_BuildListByDong"
    ws.Range("A30").Value = "NaverLocal_VerifyExistingData"
    ws.Range("A31").Value = "Address_SplitDongAndRoad"
    ws.Range("A32").Value = "NaverLocal_CreateDongMasterTemplate"
    ws.Range("A33").Value = "Roster_DetectDuplicates"
    ws.Range("A34").Value = "Roster_QualityReport"
    ws.Range("A35").Value = "Roster_MatchCompanyNames"
    ws.Range("A36").Value = "Address_SplitAdministrativeParts"
    ws.Range("A37").Value = "Phone_ValidateSelection"
    ws.Range("A38").Value = "SPSS_ExtractTableMetadata"
    ws.Range("A39").Value = "SampleDesign_CreateTemplate"
    ws.Range("A40").Value = "SampleDesign_RunAllocation"
    ws.Range("A41").Value = "SampleDesign_CreateProgressTracker"
    ws.Range("A42").Value = "SampleDesign_ListifyPopulationTable"
    ws.Range("A43").Value = "Database_MergeSheetsByHeaders"
    ws.Range("A44").Value = "Database_SplitMergedSheetBySource"
    ws.Range("A45").Value = "NaverLocal_CreateLegalDongMasterTemplate"
    ws.Range("A46").Value = "Address_CombineDongAndRoad"

    ' 내용에 맞게 열 너비 자동 조절
    ws.Columns("A:B").AutoFit
End Sub

Public Sub Toolkit_ResetAllSheetsToTop()
    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim originalSheet As Worksheet
    If TypeName(ActiveSheet) = "Worksheet" Then Set originalSheet = ActiveSheet

    Dim ws As Worksheet
    Dim resetCount As Long

    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        ws.Activate
        ws.Range("A1").Select
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
        resetCount = resetCount + 1
    Next ws

    If Not originalSheet Is Nothing Then
        originalSheet.Activate
        originalSheet.Range("A1").Select
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    End If

    Application.ScreenUpdating = True
    MsgBox resetCount & "개 시트의 화면 위치를 최상단으로 정렬했습니다.", vbInformation, "모든 시트 최상단"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "모든 시트 최상단 정렬 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "모든 시트 최상단"
End Sub

Private Sub Toolkit_ProgressShow(ByVal title As String, ByVal message As String, ByVal totalCount As Long)
    On Error Resume Next
    Set gProgressForm = VBA.UserForms.Add("frmProgressStatus")
    If Not gProgressForm Is Nothing Then
        gProgressForm.Configure title, totalCount, message
        gProgressForm.Show vbModeless
    End If
    Application.StatusBar = title & " - " & message
    DoEvents
    On Error GoTo 0
End Sub

Private Sub Toolkit_ProgressUpdate(ByVal currentCount As Long, ByVal totalCount As Long, ByVal message As String)
    On Error Resume Next
    If Not gProgressForm Is Nothing Then
        gProgressForm.UpdateProgress currentCount, totalCount, message
    End If

    Dim percentValue As Double
    If totalCount > 0 Then percentValue = currentCount / totalCount
    Application.StatusBar = Format(percentValue, "0%") & " - " & message & " (" & currentCount & "/" & totalCount & ")"
    DoEvents
    On Error GoTo 0
End Sub

Private Sub Toolkit_ProgressClose()
    On Error Resume Next
    If Not gProgressForm Is Nothing Then
        Unload gProgressForm
        Set gProgressForm = Nothing
    End If
    Application.StatusBar = False
    DoEvents
    On Error GoTo 0
End Sub

Public Sub Toolkit_SelectA1OnAllSheets()
    Toolkit_ResetAllSheetsToTop
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
    NormalizePhonesToCopiedColumns Selection
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
    NormalizePhonesToCopiedColumns targetRange
End Sub

Public Sub NormalizePhonesWithForm()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("전화번호 정리", _
        "한국 전화번호 형식을 하이픈 포함 표준 형식으로 정리합니다. 정리할 셀/열/범위를 선택하세요.", _
        targetRange) Then Exit Sub
    NormalizePhonesToCopiedColumns targetRange
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

Private Function Toolkit_GetRangeFromUserForm(ByVal formTitle As String, _
                                              ByVal guideText As String, _
                                              ByRef targetRange As Range) As Boolean
    Dim defaultAddress As String
    If TypeName(Selection) = "Range" Then defaultAddress = Selection.Areas(1).Address(External:=True)

    Dim promptText As String
    promptText = guideText & vbCrLf & vbCrLf & _
                 "범위 입력칸을 클릭한 뒤 시트에서 마우스 드래그, Shift+방향키, Ctrl+Shift+방향키로 연속 범위를 지정할 수 있습니다." & vbCrLf & _
                 "떨어진 여러 범위는 한 번에 처리하지 않습니다. 필요한 경우 기능을 여러 번 실행하세요." & vbCrLf & _
                 "취소하면 작업을 실행하지 않습니다."

    On Error Resume Next
    Set targetRange = Application.InputBox(Prompt:=promptText, _
                                           Title:=formTitle, _
                                           Default:=defaultAddress, _
                                           Type:=8)
    If Err.Number <> 0 Then
        Err.Clear
        Set targetRange = Nothing
    End If
    On Error GoTo 0

    If targetRange Is Nothing Then Exit Function

    If targetRange.Areas.Count > 1 Then
        MsgBox "떨어진 여러 범위가 선택되었습니다." & vbCrLf & vbCrLf & _
               "현재 기능은 오류 방지를 위해 하나의 연속 범위만 처리합니다." & vbCrLf & _
               "한 덩어리의 범위를 다시 선택하거나 기능을 여러 번 나누어 실행하세요.", _
               vbExclamation, formTitle
        Set targetRange = Nothing
        Exit Function
    End If

    If Toolkit_RangeHasMergedCells(targetRange) Then
        Dim mergeAnswer As VbMsgBoxResult
        mergeAnswer = MsgBox("선택 범위에 병합된 셀이 포함되어 있습니다." & vbCrLf & vbCrLf & _
                             "병합 셀은 열 추가, 행 복사, 결과 쓰기 과정에서 예기치 않은 오류를 만들 수 있습니다." & vbCrLf & _
                             "그래도 계속 진행할까요?", _
                             vbQuestion + vbYesNo, formTitle)
        If mergeAnswer <> vbYes Then
            Set targetRange = Nothing
            Exit Function
        End If
    End If

    Toolkit_GetRangeFromUserForm = True
End Function

Private Function Toolkit_RangeHasMergedCells(ByVal targetRange As Range) As Boolean
    On Error GoTo SafeExit
    Dim mergeState As Variant
    mergeState = targetRange.MergeCells

    If IsNull(mergeState) Then
        Toolkit_RangeHasMergedCells = True
    ElseIf CBool(mergeState) Then
        Toolkit_RangeHasMergedCells = True
    End If
SafeExit:
End Function

Private Sub Toolkit_BeginOperation(Optional ByVal statusMessage As String = "")
    If mToolkitOperationDepth = 0 Then
        mToolkitPrevScreenUpdating = Application.ScreenUpdating
        mToolkitPrevEnableEvents = Application.EnableEvents
        mToolkitPrevDisplayAlerts = Application.DisplayAlerts
        mToolkitPrevCalculation = Application.Calculation
        mToolkitPrevStatusBar = Application.StatusBar

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    End If

    mToolkitOperationDepth = mToolkitOperationDepth + 1
    If Len(statusMessage) > 0 Then Application.StatusBar = statusMessage
End Sub

Private Sub Toolkit_EndOperation()
    On Error Resume Next

    If mToolkitOperationDepth > 0 Then mToolkitOperationDepth = mToolkitOperationDepth - 1
    If mToolkitOperationDepth = 0 Then
        Application.ScreenUpdating = mToolkitPrevScreenUpdating
        Application.EnableEvents = mToolkitPrevEnableEvents
        Application.DisplayAlerts = mToolkitPrevDisplayAlerts
        Application.Calculation = mToolkitPrevCalculation
        Application.StatusBar = mToolkitPrevStatusBar
    End If

    On Error GoTo 0
End Sub

Private Function Toolkit_ValidateWorksheet(ByVal ws As Worksheet, ByVal purpose As String, _
                                           Optional ByVal requiresWrite As Boolean = True) As Boolean
    If ws Is Nothing Then
        MsgBox "작업할 워크시트를 찾지 못했습니다.", vbExclamation, purpose
        Exit Function
    End If

    If requiresWrite And ws.ProtectContents Then
        MsgBox "현재 시트가 보호되어 있어 '" & purpose & "' 작업을 실행할 수 없습니다." & vbCrLf & _
               "시트 보호를 해제한 뒤 다시 실행하세요.", vbExclamation, purpose
        Exit Function
    End If

    Toolkit_ValidateWorksheet = True
End Function

Private Function Toolkit_CanInsertColumns(ByVal ws As Worksheet, ByVal startCol As Long, _
                                          ByVal colCount As Long, ByVal purpose As String) As Boolean
    If Not Toolkit_ValidateWorksheet(ws, purpose, True) Then Exit Function

    If startCol < 1 Or colCount < 1 Then
        MsgBox "추가할 열 위치를 계산하지 못했습니다.", vbExclamation, purpose
        Exit Function
    End If

    If startCol > ws.Columns.Count Or startCol + colCount - 1 > ws.Columns.Count Then
        MsgBox "오른쪽에 결과 열을 추가할 공간이 부족합니다." & vbCrLf & _
               "처리 범위를 왼쪽으로 옮기거나 빈 열을 확보한 뒤 다시 실행하세요.", _
               vbExclamation, purpose
        Exit Function
    End If

    If ws.AutoFilterMode Then
        MsgBox "현재 시트에 필터가 켜져 있습니다." & vbCrLf & _
               "필터 상태에서 열을 추가하면 결과 위치가 헷갈릴 수 있으니, 필요하면 필터를 해제한 뒤 실행하세요.", _
               vbInformation, purpose
    End If

    Toolkit_CanInsertColumns = True
End Function

Private Function Toolkit_CanDeleteColumns(ByVal ws As Worksheet, ByVal startCol As Long, _
                                          ByVal colCount As Long, ByVal purpose As String) As Boolean
    If Not Toolkit_ValidateWorksheet(ws, purpose, True) Then Exit Function

    If startCol < 1 Or colCount < 1 Then
        MsgBox "삭제할 열 위치를 계산하지 못했습니다.", vbExclamation, purpose
        Exit Function
    End If

    If startCol > ws.Columns.Count Or startCol + colCount - 1 > ws.Columns.Count Then
        MsgBox "삭제할 열 범위가 시트 범위를 벗어났습니다.", vbExclamation, purpose
        Exit Function
    End If

    If colCount >= ws.Columns.Count Then
        MsgBox "시트의 모든 열을 삭제할 수는 없습니다.", vbExclamation, purpose
        Exit Function
    End If

    Toolkit_CanDeleteColumns = True
End Function

Private Function Toolkit_CanInsertRows(ByVal ws As Worksheet, ByVal startRow As Long, _
                                       ByVal rowCount As Long, ByVal purpose As String) As Boolean
    If Not Toolkit_ValidateWorksheet(ws, purpose, True) Then Exit Function

    If startRow < 1 Or rowCount < 1 Then
        MsgBox "추가할 행 위치를 계산하지 못했습니다.", vbExclamation, purpose
        Exit Function
    End If

    If startRow > ws.Rows.Count Or startRow + rowCount - 1 > ws.Rows.Count Then
        MsgBox "아래쪽에 결과 행을 추가할 공간이 부족합니다.", vbExclamation, purpose
        Exit Function
    End If

    Toolkit_CanInsertRows = True
End Function

Private Function Toolkit_CanDeleteRows(ByVal ws As Worksheet, ByVal startRow As Long, _
                                       ByVal rowCount As Long, ByVal purpose As String) As Boolean
    If Not Toolkit_ValidateWorksheet(ws, purpose, True) Then Exit Function

    If startRow < 1 Or rowCount < 1 Then
        MsgBox "삭제할 행 위치를 계산하지 못했습니다.", vbExclamation, purpose
        Exit Function
    End If

    If startRow > ws.Rows.Count Or startRow + rowCount - 1 > ws.Rows.Count Then
        MsgBox "삭제할 행 범위가 시트 범위를 벗어났습니다.", vbExclamation, purpose
        Exit Function
    End If

    If rowCount >= ws.Rows.Count Then
        MsgBox "시트의 모든 행을 삭제할 수는 없습니다.", vbExclamation, purpose
        Exit Function
    End If

    Toolkit_CanDeleteRows = True
End Function

Private Sub Toolkit_LogApiEvent(ByVal wb As Workbook, ByVal sourceName As String, ByVal actionName As String, _
                                ByVal targetText As String, ByVal statusText As String, ByVal detailText As String)
    On Error Resume Next

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = wb.Worksheets(LOG_SHEET)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = LOG_SHEET
        ws.Range("A1:F1").Value = Array("일시", "기능", "작업", "대상", "상태", "상세")
        ws.Rows(1).Font.Bold = True
        ws.Visible = xlSheetVeryHidden
    End If

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = sourceName
    ws.Cells(nextRow, 3).Value = actionName
    ws.Cells(nextRow, 4).Value = Left$(targetText, 250)
    ws.Cells(nextRow, 5).Value = Left$(statusText, 120)
    ws.Cells(nextRow, 6).Value = Left$(detailText, 1000)
    ws.Visible = xlSheetVeryHidden

    On Error GoTo 0
End Sub

Private Function Toolkit_GetActualDataRange(ByVal ws As Worksheet) As Range
    Dim firstCell As Range
    Dim lastRowCell As Range
    Dim lastColCell As Range

    On Error Resume Next
    Set firstCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                  MatchCase:=False)
    Set lastRowCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    Set lastColCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    On Error GoTo 0

    If firstCell Is Nothing Or lastRowCell Is Nothing Or lastColCell Is Nothing Then Exit Function

    Set Toolkit_GetActualDataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRowCell.Row, lastColCell.Column))
End Function

Private Function Toolkit_TrimRangeToActualData(ByVal target As Range) As Range
    Dim actualData As Range
    Set actualData = Toolkit_GetActualDataRange(target.Worksheet)
    If actualData Is Nothing Then Exit Function
    Set Toolkit_TrimRangeToActualData = Intersect(target, actualData)
End Function

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
' 프로시저 : NormalizePhonesToCopiedColumns (내부 전용)
' 설  명  : 선택 범위를 오른쪽 새 열에 복사한 뒤 복사본만 전화번호
'           표준 형식으로 정규화한다.
' ============================================================
Private Sub NormalizePhonesToCopiedColumns(ByVal target As Range)
    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(target)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 처리할 데이터를 찾지 못했습니다.", vbInformation, "전화번호 정리"
        Exit Sub
    End If

    Dim outputRange As Range
    Set outputRange = CopyRangeToAdjacentColumns(workRange, "전화번호_정리")
    If outputRange Is Nothing Then Exit Sub

    NormalizePhonesInRange outputRange, "전화번호 정리 복사본"
End Sub

' ============================================================
' 프로시저 : NormalizePhonesInRange (내부 전용)
' 설  명  : 주어진 Range 내 각 셀을 순회하며
'           한국 전화번호를 표준 형식으로 정규화한다.
'           셀 서식을 텍스트(@)로 변경해 앞자리 0이 유지되도록 한다.
'           수식 셀은 건너뛰고 완료 후 통계를 표시한다.
' 파라미터 : target - 정규화 대상 Range
' ============================================================
Private Sub NormalizePhonesInRange(ByVal target As Range, Optional ByVal resultTitle As String = "전화번호 정규화")
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
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다." & vbCrLf & _
           "원본 범위는 변경하지 않았습니다.", _
           vbInformation, resultTitle
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, resultTitle
End Sub

Private Function CopyRangeToAdjacentColumns(ByVal sourceRange As Range, ByVal headerSuffix As String) As Range
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

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    For Each cell In outputRange.Cells
        If cell.HasFormula Then
            skippedFormulaCount = skippedFormulaCount + 1
        Else
            digits = DigitsOnly(CStr(cell.Value))
            If Len(digits) > 0 And CStr(cell.Value) <> digits Then
                cell.NumberFormat = "@"
                cell.Value = digits
                changedCount = changedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox changedCount & "개 셀에서 대시/구분기호를 제거했습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다." & vbCrLf & _
           "원본 범위는 변경하지 않았습니다.", vbInformation, resultTitle
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, resultTitle
End Sub

Public Sub CompanyName_CleanWithOptions()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("기업체명/사업체명 정리", _
        "정리할 기업체명 또는 사업체명 범위를 선택하세요. 원본 오른쪽에 복사본 열을 만들고 선택한 규칙을 적용합니다.", _
        targetRange) Then Exit Sub

    Dim options As Object
    Set options = CompanyName_GetOptionsFromUser()
    If options Is Nothing Then Exit Sub

    CompanyName_CleanRangeToCopiedColumns targetRange, options
End Sub

Private Function CompanyName_GetOptionsFromUser() As Object
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmCompanyNameOptions")
    form.Show

    If form.Cancelled Then
        Unload form
        Exit Function
    End If

    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    options("trimOuter") = CBool(form.TrimOuter)
    options("removeSpaces") = CBool(form.RemoveSpaces)
    options("normalizeCorpSymbol") = CBool(form.NormalizeCorpSymbol)
    options("normalizeCorpText") = CBool(form.NormalizeCorpText)
    options("moveCorpToFront") = CBool(form.MoveCorpToFront)
    options("moveCorpToEnd") = CBool(form.MoveCorpToEnd)
    options("removeCorpMarkers") = CBool(form.RemoveCorpMarkers)
    options("addCorpPrefix") = CBool(form.AddCorpPrefix)
    options("addCorpSuffix") = CBool(form.AddCorpSuffix)
    options("cutBeforePossessive") = CBool(form.CutBeforePossessive)
    options("pronounceEnglishWords") = CBool(form.PronounceEnglishWords)
    options("spellAlphabetKorean") = CBool(form.SpellAlphabetKorean)

    Unload form
    Set CompanyName_GetOptionsFromUser = options
    Exit Function

Fallback:
    Set CompanyName_GetOptionsFromUser = CompanyName_DefaultOptions()
End Function

Private Function CompanyName_DefaultOptions() As Object
    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    options("trimOuter") = False
    options("removeSpaces") = True
    options("normalizeCorpSymbol") = False
    options("normalizeCorpText") = False
    options("moveCorpToFront") = False
    options("moveCorpToEnd") = False
    options("removeCorpMarkers") = True
    options("addCorpPrefix") = False
    options("addCorpSuffix") = False
    options("cutBeforePossessive") = True
    options("pronounceEnglishWords") = False
    options("spellAlphabetKorean") = True
    Set CompanyName_DefaultOptions = options
End Function

Private Sub CompanyName_CleanRangeToCopiedColumns(ByVal target As Range, ByVal options As Object)
    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(target)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 처리할 기업체명/사업체명을 찾지 못했습니다.", vbInformation, "기업체명 정리"
        Exit Sub
    End If

    Dim outputRange As Range
    Set outputRange = CopyRangeToAdjacentColumns(workRange, "기업명_정리")
    If outputRange Is Nothing Then Exit Sub

    Dim cell As Range
    Dim changedCount As Long
    Dim skippedFormulaCount As Long
    Dim cleaned As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    For Each cell In outputRange.Cells
        If cell.HasFormula Then
            skippedFormulaCount = skippedFormulaCount + 1
        Else
            cleaned = CompanyName_CleanOne(CStr(cell.Value), options)
            If CStr(cell.Value) <> cleaned Then
                cell.NumberFormat = "@"
                cell.Value = cleaned
                changedCount = changedCount + 1
            End If
        End If
    Next cell

    outputRange.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox changedCount & "개 기업체명/사업체명을 정리했습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다." & vbCrLf & _
           "원본 범위는 변경하지 않았습니다.", vbInformation, "기업체명 정리"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "기업체명 정리 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "기업체명 정리"
End Sub

Private Function CompanyName_CleanOne(ByVal valueIn As String, ByVal options As Object) As String
    Dim textValue As String
    textValue = CStr(valueIn)

    If CBool(options("cutBeforePossessive")) Then
        Dim cutPos As Long
        cutPos = InStr(1, textValue, "의 ", vbTextCompare)
        If cutPos > 0 Then textValue = Left$(textValue, cutPos - 1)
    End If

    If CBool(options("trimOuter")) Then textValue = Trim$(textValue)
    If CBool(options("pronounceEnglishWords")) Then textValue = CompanyName_PronounceEnglishWords(textValue)
    If CBool(options("removeSpaces")) And Not CBool(options("pronounceEnglishWords")) Then textValue = CompanyName_RemoveSpaces(textValue)
    If CBool(options("normalizeCorpSymbol")) Then textValue = Replace(textValue, "㈜", "(주)")
    If CBool(options("normalizeCorpText")) Then textValue = Replace(textValue, "주식회사", "(주)")
    If CBool(options("moveCorpToFront")) Then textValue = CompanyName_MoveCorpMarker(textValue, True)
    If CBool(options("moveCorpToEnd")) Then textValue = CompanyName_MoveCorpMarker(textValue, False)
    If CBool(options("removeCorpMarkers")) Then textValue = CompanyName_RemoveCorpMarkers(textValue)
    If CBool(options("addCorpPrefix")) Then
        If Not CompanyName_HasCorpMarker(textValue) Then textValue = "(주)" & textValue
    End If
    If CBool(options("addCorpSuffix")) Then
        If Not CompanyName_HasCorpMarker(textValue) Then textValue = textValue & "(주)"
    End If
    If CBool(options("spellAlphabetKorean")) Then textValue = CompanyName_SpellAlphabetKorean(textValue)

    CompanyName_CleanOne = textValue
End Function

Private Function CompanyName_RemoveSpaces(ByVal textValue As String) As String
    textValue = Replace(textValue, ChrW(160), "")
    textValue = Replace(textValue, vbTab, "")
    textValue = Replace(textValue, " ", "")
    CompanyName_RemoveSpaces = textValue
End Function

Private Function CompanyName_HasCorpMarker(ByVal textValue As String) As Boolean
    CompanyName_HasCorpMarker = (InStr(1, textValue, "(주)", vbTextCompare) > 0 _
                             Or InStr(1, textValue, "㈜", vbTextCompare) > 0 _
                             Or InStr(1, textValue, "주식회사", vbTextCompare) > 0)
End Function

Private Function CompanyName_RemoveCorpMarkers(ByVal textValue As String) As String
    textValue = Replace(textValue, "(주)", "")
    textValue = Replace(textValue, "㈜", "")
    textValue = Replace(textValue, "주식회사", "")
    CompanyName_RemoveCorpMarkers = textValue
End Function

Private Function CompanyName_MoveCorpMarker(ByVal textValue As String, ByVal moveToFront As Boolean) As String
    If Not CompanyName_HasCorpMarker(textValue) Then
        CompanyName_MoveCorpMarker = textValue
        Exit Function
    End If

    Dim baseName As String
    baseName = CompanyName_RemoveCorpMarkers(textValue)
    If moveToFront Then
        CompanyName_MoveCorpMarker = "(주)" & baseName
    Else
        CompanyName_MoveCorpMarker = baseName & "(주)"
    End If
End Function

Private Function CompanyName_PronounceEnglishWords(ByVal textValue As String) As String
    Dim result As String
    Dim token As String
    Dim i As Long

    For i = 1 To Len(textValue)
        Dim ch As String
        ch = Mid$(textValue, i, 1)
        If CompanyName_IsAsciiLetter(ch) Then
            token = token & ch
        Else
            If Len(token) > 0 Then
                result = result & CompanyName_EnglishWordKorean(token)
                token = ""
            End If
            result = result & ch
        End If
    Next i

    If Len(token) > 0 Then result = result & CompanyName_EnglishWordKorean(token)
    CompanyName_PronounceEnglishWords = result
End Function

Private Function CompanyName_IsAsciiLetter(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Dim codeValue As Long
    codeValue = AscW(ch)
    CompanyName_IsAsciiLetter = ((codeValue >= 65 And codeValue <= 90) Or (codeValue >= 97 And codeValue <= 122))
End Function

Private Function CompanyName_EnglishWordKorean(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "ai": CompanyName_EnglishWordKorean = "에이아이"
        Case "app": CompanyName_EnglishWordKorean = "앱"
        Case "bank": CompanyName_EnglishWordKorean = "뱅크"
        Case "bio": CompanyName_EnglishWordKorean = "바이오"
        Case "cloud": CompanyName_EnglishWordKorean = "클라우드"
        Case "commerce": CompanyName_EnglishWordKorean = "커머스"
        Case "communication", "communications": CompanyName_EnglishWordKorean = "커뮤니케이션"
        Case "company": CompanyName_EnglishWordKorean = "컴퍼니"
        Case "consulting": CompanyName_EnglishWordKorean = "컨설팅"
        Case "contents", "content": CompanyName_EnglishWordKorean = "콘텐츠"
        Case "data": CompanyName_EnglishWordKorean = "데이터"
        Case "design": CompanyName_EnglishWordKorean = "디자인"
        Case "digital": CompanyName_EnglishWordKorean = "디지털"
        Case "edu", "education": CompanyName_EnglishWordKorean = "에듀"
        Case "energy": CompanyName_EnglishWordKorean = "에너지"
        Case "entertainment": CompanyName_EnglishWordKorean = "엔터테인먼트"
        Case "factory": CompanyName_EnglishWordKorean = "팩토리"
        Case "finance", "financial": CompanyName_EnglishWordKorean = "파이낸스"
        Case "food": CompanyName_EnglishWordKorean = "푸드"
        Case "game", "games", "gaming": CompanyName_EnglishWordKorean = "게임"
        Case "global": CompanyName_EnglishWordKorean = "글로벌"
        Case "group": CompanyName_EnglishWordKorean = "그룹"
        Case "health", "healthcare": CompanyName_EnglishWordKorean = "헬스케어"
        Case "holdings", "holding": CompanyName_EnglishWordKorean = "홀딩스"
        Case "hotel": CompanyName_EnglishWordKorean = "호텔"
        Case "inc": CompanyName_EnglishWordKorean = "인크"
        Case "innovation": CompanyName_EnglishWordKorean = "이노베이션"
        Case "international": CompanyName_EnglishWordKorean = "인터내셔널"
        Case "investment": CompanyName_EnglishWordKorean = "인베스트먼트"
        Case "it": CompanyName_EnglishWordKorean = "아이티"
        Case "korea": CompanyName_EnglishWordKorean = "코리아"
        Case "lab", "labs": CompanyName_EnglishWordKorean = "랩"
        Case "life": CompanyName_EnglishWordKorean = "라이프"
        Case "logistics": CompanyName_EnglishWordKorean = "로지스틱스"
        Case "mall": CompanyName_EnglishWordKorean = "몰"
        Case "marketing": CompanyName_EnglishWordKorean = "마케팅"
        Case "media": CompanyName_EnglishWordKorean = "미디어"
        Case "medical": CompanyName_EnglishWordKorean = "메디컬"
        Case "mobility": CompanyName_EnglishWordKorean = "모빌리티"
        Case "mobile": CompanyName_EnglishWordKorean = "모바일"
        Case "network", "networks": CompanyName_EnglishWordKorean = "네트워크"
        Case "online": CompanyName_EnglishWordKorean = "온라인"
        Case "partners": CompanyName_EnglishWordKorean = "파트너스"
        Case "platform": CompanyName_EnglishWordKorean = "플랫폼"
        Case "plus": CompanyName_EnglishWordKorean = "플러스"
        Case "power": CompanyName_EnglishWordKorean = "파워"
        Case "project": CompanyName_EnglishWordKorean = "프로젝트"
        Case "research": CompanyName_EnglishWordKorean = "리서치"
        Case "service", "services": CompanyName_EnglishWordKorean = "서비스"
        Case "shop": CompanyName_EnglishWordKorean = "샵"
        Case "smart": CompanyName_EnglishWordKorean = "스마트"
        Case "soft", "software": CompanyName_EnglishWordKorean = "소프트웨어"
        Case "solution", "solutions": CompanyName_EnglishWordKorean = "솔루션"
        Case "studio": CompanyName_EnglishWordKorean = "스튜디오"
        Case "system", "systems": CompanyName_EnglishWordKorean = "시스템"
        Case "tech": CompanyName_EnglishWordKorean = "테크"
        Case "technology", "technologies": CompanyName_EnglishWordKorean = "테크놀로지"
        Case "telecom": CompanyName_EnglishWordKorean = "텔레콤"
        Case "tour": CompanyName_EnglishWordKorean = "투어"
        Case "trade", "trading": CompanyName_EnglishWordKorean = "트레이딩"
        Case "venture", "ventures": CompanyName_EnglishWordKorean = "벤처스"
        Case "vision": CompanyName_EnglishWordKorean = "비전"
        Case "works": CompanyName_EnglishWordKorean = "웍스"
        Case "world": CompanyName_EnglishWordKorean = "월드"
        Case Else: CompanyName_EnglishWordKorean = wordText
    End Select
End Function

Private Function CompanyName_SpellAlphabetKorean(ByVal textValue As String) As String
    Dim result As String
    Dim i As Long
    For i = 1 To Len(textValue)
        result = result & CompanyName_AlphabetKorean(Mid$(textValue, i, 1))
    Next i
    CompanyName_SpellAlphabetKorean = result
End Function

Private Function CompanyName_AlphabetKorean(ByVal ch As String) As String
    Select Case UCase$(ch)
        Case "A": CompanyName_AlphabetKorean = "에이"
        Case "B": CompanyName_AlphabetKorean = "비"
        Case "C": CompanyName_AlphabetKorean = "씨"
        Case "D": CompanyName_AlphabetKorean = "디"
        Case "E": CompanyName_AlphabetKorean = "이"
        Case "F": CompanyName_AlphabetKorean = "에프"
        Case "G": CompanyName_AlphabetKorean = "지"
        Case "H": CompanyName_AlphabetKorean = "에이치"
        Case "I": CompanyName_AlphabetKorean = "아이"
        Case "J": CompanyName_AlphabetKorean = "제이"
        Case "K": CompanyName_AlphabetKorean = "케이"
        Case "L": CompanyName_AlphabetKorean = "엘"
        Case "M": CompanyName_AlphabetKorean = "엠"
        Case "N": CompanyName_AlphabetKorean = "엔"
        Case "O": CompanyName_AlphabetKorean = "오"
        Case "P": CompanyName_AlphabetKorean = "피"
        Case "Q": CompanyName_AlphabetKorean = "큐"
        Case "R": CompanyName_AlphabetKorean = "알"
        Case "S": CompanyName_AlphabetKorean = "에스"
        Case "T": CompanyName_AlphabetKorean = "티"
        Case "U": CompanyName_AlphabetKorean = "유"
        Case "V": CompanyName_AlphabetKorean = "브이"
        Case "W": CompanyName_AlphabetKorean = "더블유"
        Case "X": CompanyName_AlphabetKorean = "엑스"
        Case "Y": CompanyName_AlphabetKorean = "와이"
        Case "Z": CompanyName_AlphabetKorean = "지"
        Case Else: CompanyName_AlphabetKorean = ch
    End Select
End Function

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
    Next r

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(lastRow, outputCol + 1)).EntireColumn.AutoFit
    Application.ScreenUpdating = True

    MsgBox "주소 유형 판별을 완료했습니다." & vbCrLf & _
           "도로명주소: " & roadCount & "건" & vbCrLf & _
           "법정/행정동 기준 주소: " & dongCount & "건" & vbCrLf & _
           "확인 필요: " & unknownCount & "건", vbInformation, "주소 유형 판별"
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

    For r = dataStartRow To lastRow
        sourceAddress = Trim$(CStr(ws.Cells(r, workRange.Column).Value))
        If Len(sourceAddress) > 0 Then
            splitResult = Address_SplitOneAddress(sourceAddress, records)
            ws.Cells(r, outputCol).Value = CStr(splitResult(0))
            ws.Cells(r, outputCol + 1).Value = CStr(splitResult(1))
            If CBool(splitResult(2)) Then matchedCount = matchedCount + 1 Else heuristicCount = heuristicCount + 1
        End If
    Next r

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(lastRow, outputCol + 1)).EntireColumn.AutoFit
    Application.ScreenUpdating = True

    MsgBox "주소 체계 분리를 완료했습니다." & vbCrLf & _
           "변환표 매칭: " & matchedCount & "건" & vbCrLf & _
           "주소 형태 기반 분리: " & heuristicCount & "건" & vbCrLf & _
           "변환표가 있으면 지번/도로명 양방향 매칭 정확도가 높아집니다.", vbInformation, "주소 체계 분리"
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
    Next r

    ws.Columns(outputCol).AutoFit
    Application.ScreenUpdating = True

    MsgBox "분할 주소 통합을 완료했습니다." & vbCrLf & _
           "도로명주소 사용: " & roadSelectedCount & "건" & vbCrLf & _
           "동 기준 주소 사용: " & dongSelectedCount & "건" & vbCrLf & _
           "빈값: " & blankCount & "건", vbInformation, "주소 통합"
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
    Next r

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(lastRow, outputCol + 2)).EntireColumn.AutoFit
    Application.ScreenUpdating = True

    MsgBox "도로명주소 변환을 완료했습니다." & vbCrLf & _
           "성공: " & successCount & "건" & vbCrLf & _
           "실패: " & failCount & "건", vbInformation, "도로명주소 변환"
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
    Dim textValue As String
    textValue = addressText
    textValue = Replace(textValue, "서울시", "서울특별시")
    textValue = Replace(textValue, "부산시", "부산광역시")
    textValue = Replace(textValue, "대구시", "대구광역시")
    textValue = Replace(textValue, "인천시", "인천광역시")
    textValue = Replace(textValue, "광주시", "광주광역시")
    textValue = Replace(textValue, "대전시", "대전광역시")
    textValue = Replace(textValue, "울산시", "울산광역시")
    textValue = Replace(textValue, "세종시", "세종특별자치시")
    textValue = Replace(textValue, "경기 ", "경기도 ")
    textValue = Replace(textValue, "강원 ", "강원도 ")
    textValue = Replace(textValue, "충북 ", "충청북도 ")
    textValue = Replace(textValue, "충남 ", "충청남도 ")
    textValue = Replace(textValue, "전북 ", "전라북도 ")
    textValue = Replace(textValue, "전남 ", "전라남도 ")
    textValue = Replace(textValue, "경북 ", "경상북도 ")
    textValue = Replace(textValue, "경남 ", "경상남도 ")
    textValue = Replace(textValue, "제주도", "제주특별자치도")
    Address_NormalizeRegionNames = textValue
End Function

Private Function Address_IsRoadAddress(ByVal addressText As String) As Boolean
    Dim textValue As String
    textValue = " " & Replace(Trim$(CStr(addressText)), "  ", " ") & " "
    Address_IsRoadAddress = (InStr(1, textValue, "로 ", vbTextCompare) > 0 _
                          Or InStr(1, textValue, "길 ", vbTextCompare) > 0 _
                          Or InStr(1, textValue, "대로 ", vbTextCompare) > 0 _
                          Or InStr(1, textValue, "번길 ", vbTextCompare) > 0)
End Function

Private Function Address_ClassifyOneAddress(ByVal addressText As String) As Variant
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

    If Not Toolkit_GetRangeFromUserForm("사업자등록번호 3열 병합", _
        "사업자등록번호가 3개 열로 나뉜 범위를 선택하세요. Ctrl+Shift+방향키, Shift+방향키, 마우스 드래그를 모두 사용할 수 있습니다.", _
        targetRange) Then Exit Sub

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

    If Not Toolkit_CanInsertColumns(ws, outputCol, 1, "사업자등록번호 병합") Then Exit Sub
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
    Dim selectedFields As Collection
    Dim responseCache As Object

    If Not Toolkit_GetRangeFromUserForm("사업자등록 상태조회", _
        "조회할 사업자등록번호 범위를 선택하세요. Ctrl+Shift+방향키, Shift+방향키, 마우스 드래그를 모두 사용할 수 있습니다.", _
        targetRange) Then Exit Sub

    Set targetRange = targetRange.Areas(1)
    If targetRange.Columns.Count > 1 Then
        Set targetRange = targetRange.Columns(1)
    End If

    apiKey = Biz_GetApiKey()
    If Len(apiKey) = 0 Then
        MsgBox "저장된 인증키가 없습니다. 먼저 [인증키 설정]을 실행하세요.", vbExclamation, "사업자등록 상태조회"
        Exit Sub
    End If

    Set selectedFields = Biz_GetStatusFieldOptions()
    If selectedFields Is Nothing Then Exit Sub
    If selectedFields.Count = 0 Then
        MsgBox "반환받을 값을 하나 이상 선택하세요.", vbExclamation, "사업자등록 상태조회"
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
    If Not Toolkit_CanInsertColumns(ws, outputCol, selectedFields.Count, "사업자등록 상태조회") Then Exit Sub
    ws.Columns(outputCol).Resize(, selectedFields.Count).Insert Shift:=xlToRight
    If headerRow > 0 Then Biz_WriteStatusHeaders ws, headerRow, outputCol, selectedFields

    Application.ScreenUpdating = False
    Toolkit_ProgressShow "사업자등록 상태조회", "조회 준비 중", lastRow - dataStartRow + 1
    On Error GoTo CleanFail
    Set responseCache = CreateObject("Scripting.Dictionary")

    For rowIndex = dataStartRow To lastRow
        Toolkit_ProgressUpdate rowIndex - dataStartRow + 1, lastRow - dataStartRow + 1, "사업자등록번호 확인 중: " & CStr(rowIndex) & "행"
        Dim bizNo As String
        bizNo = DigitsOnly(CStr(ws.Cells(rowIndex, targetRange.Column).Value))
        If Len(bizNo) = 10 Then
            requestCount = requestCount + 1
            If responseCache.Exists(bizNo) Then
                Dim cachedStatus As Variant
                cachedStatus = responseCache(bizNo)
                Biz_WriteStatusRow ws, rowIndex, outputCol, CStr(cachedStatus(0)), bizNo, CStr(cachedStatus(1)), selectedFields
                successCount = successCount + 1
            Else
                batchCount = batchCount + 1
                batchNos(batchCount) = bizNo
                batchRows(batchCount) = rowIndex

                If batchCount = 100 Then
                    Biz_QueryStatusBatch ws, outputCol, batchNos, batchRows, batchCount, apiKey, selectedFields, responseCache, successCount, failCount
                    batchCount = 0
                End If
            End If
        ElseIf Len(Trim$(CStr(ws.Cells(rowIndex, targetRange.Column).Value))) > 0 Then
            failCount = failCount + 1
            Biz_WriteStatusError ws, rowIndex, outputCol, CStr(ws.Cells(rowIndex, targetRange.Column).Value), "사업자등록번호는 숫자 10자리여야 합니다.", selectedFields
        End If
    Next rowIndex

    If batchCount > 0 Then
        Biz_QueryStatusBatch ws, outputCol, batchNos, batchRows, batchCount, apiKey, selectedFields, responseCache, successCount, failCount
    End If

    ws.Range(ws.Cells(firstRow, outputCol), ws.Cells(firstRow, outputCol + selectedFields.Count - 1)).EntireColumn.AutoFit
    Toolkit_ProgressClose
    Application.ScreenUpdating = True

    MsgBox "사업자등록 상태조회를 완료했습니다." & vbCrLf & _
           "요청: " & requestCount & "건" & vbCrLf & _
           "성공: " & successCount & "건" & vbCrLf & _
           "확인 필요: " & failCount & "건", vbInformation, "사업자등록 상태조회"
    Exit Sub

CleanFail:
    Toolkit_ProgressClose
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

Private Function Biz_GetStatusFieldOptions() As Collection
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmBizStatusFieldsOptions")
    form.Show
    If form.Cancelled Then
        Unload form
        Exit Function
    End If

    Set Biz_GetStatusFieldOptions = Biz_FieldsFromCsv(CStr(form.SelectedKeys))
    Unload form
    Exit Function

Fallback:
    Set Biz_GetStatusFieldOptions = Biz_DefaultStatusFields()
End Function

Private Function Biz_DefaultStatusFields() As Collection
    Dim fields As New Collection
    Dim keys As Variant
    Dim i As Long
    keys = Array("status_code", "requested_b_no", "b_no", "b_stt", "b_stt_cd", "tax_type", "tax_type_cd", "end_dt", "utcc_yn", "tax_type_change_dt", "invoice_apply_dt", "rbf_tax_type", "rbf_tax_type_cd")
    For i = LBound(keys) To UBound(keys)
        fields.Add CStr(keys(i))
    Next i
    Set Biz_DefaultStatusFields = fields
End Function

Private Function Biz_FieldsFromCsv(ByVal csvText As String) As Collection
    Dim fields As New Collection
    Dim parts As Variant
    Dim i As Long
    parts = Split(csvText, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(CStr(parts(i)))) > 0 Then fields.Add Trim$(CStr(parts(i)))
    Next i
    Set Biz_FieldsFromCsv = fields
End Function

Private Sub Biz_WriteStatusHeaders(ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal selectedFields As Collection)
    Dim i As Long
    For i = 1 To selectedFields.Count
        ws.Cells(rowIndex, colIndex + i - 1).Value = Biz_StatusFieldHeader(CStr(selectedFields(i)))
    Next i
End Sub

Private Function Biz_StatusFieldHeader(ByVal fieldKey As String) As String
    Select Case fieldKey
        Case "status_code": Biz_StatusFieldHeader = "status_code(응답상태코드)"
        Case "requested_b_no": Biz_StatusFieldHeader = "requested_b_no(요청 사업자등록번호)"
        Case "b_no": Biz_StatusFieldHeader = "b_no(사업자등록번호)"
        Case "b_stt": Biz_StatusFieldHeader = "b_stt(납세자상태)"
        Case "b_stt_cd": Biz_StatusFieldHeader = "b_stt_cd(납세자상태코드)"
        Case "tax_type": Biz_StatusFieldHeader = "tax_type(과세유형)"
        Case "tax_type_cd": Biz_StatusFieldHeader = "tax_type_cd(과세유형코드)"
        Case "end_dt": Biz_StatusFieldHeader = "end_dt(폐업일자)"
        Case "utcc_yn": Biz_StatusFieldHeader = "utcc_yn(단위과세전환폐업여부)"
        Case "tax_type_change_dt": Biz_StatusFieldHeader = "tax_type_change_dt(과세유형전환일자)"
        Case "invoice_apply_dt": Biz_StatusFieldHeader = "invoice_apply_dt(세금계산서적용일자)"
        Case "rbf_tax_type": Biz_StatusFieldHeader = "rbf_tax_type(직전과세유형)"
        Case "rbf_tax_type_cd": Biz_StatusFieldHeader = "rbf_tax_type_cd(직전과세유형코드)"
        Case Else: Biz_StatusFieldHeader = fieldKey
    End Select
End Function

Private Sub Biz_QueryStatusBatch(ws As Worksheet, ByVal outputCol As Long, ByRef batchNos() As String, ByRef batchRows() As Long, ByVal batchCount As Long, ByVal apiKey As String, ByVal selectedFields As Collection, ByVal responseCache As Object, ByRef successCount As Long, ByRef failCount As Long)
    Dim responseText As String
    Dim statusCode As String
    Dim dataObjects As Collection
    Dim i As Long

    On Error GoTo BatchFail
    responseText = Biz_PostStatus(batchNos, batchCount, apiKey)
    statusCode = Biz_JsonString(responseText, "status_code")
    If Len(statusCode) = 0 Then statusCode = "UNKNOWN"

    Set dataObjects = Biz_ExtractDataObjects(responseText)
    For i = 1 To batchCount
        If i <= dataObjects.Count Then
            Biz_WriteStatusRow ws, batchRows(i), outputCol, statusCode, batchNos(i), CStr(dataObjects(i)), selectedFields
            If Not responseCache Is Nothing Then responseCache(batchNos(i)) = Array(statusCode, CStr(dataObjects(i)))
            successCount = successCount + 1
        Else
            failCount = failCount + 1
            Biz_WriteStatusError ws, batchRows(i), outputCol, batchNos(i), "응답 데이터가 없습니다. status_code=" & statusCode, selectedFields
        End If
    Next i
    Exit Sub

BatchFail:
    For i = 1 To batchCount
        failCount = failCount + 1
        Biz_WriteStatusError ws, batchRows(i), outputCol, batchNos(i), "API 조회 실패: " & Err.Description, selectedFields
        Toolkit_LogApiEvent ws.Parent, "사업자등록 상태조회", "batch", batchNos(i), "ERROR", Err.Description
    Next i
    Err.Clear
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
    http.setTimeouts API_RESOLVE_TIMEOUT_MS, API_CONNECT_TIMEOUT_MS, API_SEND_TIMEOUT_MS, API_RECEIVE_TIMEOUT_MS
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

Private Sub Biz_WriteStatusRow(ws As Worksheet, ByVal rowIndex As Long, ByVal outputCol As Long, ByVal statusCode As String, ByVal requestedNo As String, ByVal dataObject As String, ByVal selectedFields As Collection)
    Dim i As Long
    Dim fieldKey As String
    For i = 1 To selectedFields.Count
        fieldKey = CStr(selectedFields(i))
        If fieldKey = "requested_b_no" Or fieldKey = "b_no" Then ws.Cells(rowIndex, outputCol + i - 1).NumberFormat = "@"
        ws.Cells(rowIndex, outputCol + i - 1).Value = Biz_StatusFieldValue(fieldKey, statusCode, requestedNo, dataObject)
    Next i
End Sub

Private Sub Biz_WriteStatusError(ws As Worksheet, ByVal rowIndex As Long, ByVal outputCol As Long, ByVal requestedNo As String, ByVal message As String, ByVal selectedFields As Collection)
    Dim i As Long
    Dim fieldKey As String
    For i = 1 To selectedFields.Count
        fieldKey = CStr(selectedFields(i))
        If fieldKey = "requested_b_no" Or fieldKey = "b_no" Then ws.Cells(rowIndex, outputCol + i - 1).NumberFormat = "@"

        Select Case fieldKey
            Case "status_code"
                ws.Cells(rowIndex, outputCol + i - 1).Value = "ERROR"
            Case "requested_b_no"
                ws.Cells(rowIndex, outputCol + i - 1).Value = Biz_FormatNo(requestedNo)
            Case "b_stt", "tax_type"
                ws.Cells(rowIndex, outputCol + i - 1).Value = message
        End Select
    Next i
End Sub

Private Function Biz_StatusFieldValue(ByVal fieldKey As String, ByVal statusCode As String, ByVal requestedNo As String, ByVal dataObject As String) As String
    Select Case fieldKey
        Case "status_code"
            Biz_StatusFieldValue = statusCode
        Case "requested_b_no"
            Biz_StatusFieldValue = Biz_FormatNo(requestedNo)
        Case "b_no"
            Biz_StatusFieldValue = Biz_FormatNo(Biz_JsonString(dataObject, "b_no"))
        Case Else
            Biz_StatusFieldValue = Biz_JsonString(dataObject, fieldKey)
    End Select
End Function

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
' 네이버 지역검색 API
' ============================================================
Public Sub NaverLocal_SetApiCredentials()
    Dim clientId As String
    Dim clientSecret As String

    clientId = InputBox("네이버 Developers 애플리케이션의 Client ID를 입력하세요.", "네이버 지역검색 인증키", NaverLocal_GetClientId())
    If Len(Trim$(clientId)) = 0 Then Exit Sub

    clientSecret = InputBox("네이버 Developers 애플리케이션의 Client Secret을 입력하세요.", "네이버 지역검색 인증키", NaverLocal_GetClientSecret())
    If Len(Trim$(clientSecret)) = 0 Then Exit Sub

    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_CLIENT_ID_KEY, Trim$(clientId)
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_CLIENT_SECRET_KEY, Trim$(clientSecret)

    MsgBox "네이버 지역검색 API 인증키를 저장했습니다.", vbInformation, "네이버 지역검색"
End Sub

Public Sub NaverLocal_BuildListByDong()
    If Not NaverLocal_HasCredentials() Then
        MsgBox "먼저 [네이버 인증키 설정] 버튼으로 Client ID와 Client Secret을 저장하세요.", vbExclamation, "네이버 지역검색"
        Exit Sub
    End If

    Dim options As Object
    Set options = NaverLocal_GetBuildOptions()
    If options Is Nothing Then Exit Sub

    Dim basisLabel As String
    basisLabel = NaverLocal_DongBasisLabel(options)

    Dim dongNames As Collection
    Set dongNames = NaverLocal_GetDongNamesForSearch(ActiveWorkbook, options)
    If dongNames.Count = 0 Then
        MsgBox "사용할 " & basisLabel & " 마스터를 찾지 못했거나 필터 조건에 맞는 " & basisLabel & "이 없습니다." & vbCrLf & _
               "[" & basisLabel & " 양식] 버튼으로 마스터 시트를 만든 뒤 다시 실행하세요.", vbExclamation, "네이버 지역검색"
        Exit Sub
    End If

    If dongNames.Count > 300 Then
        If MsgBox("검색 대상 " & basisLabel & "이 " & dongNames.Count & "개입니다." & vbCrLf & _
                  "네이버 API 요청이 많이 발생할 수 있습니다. 계속할까요?", _
                  vbQuestion + vbYesNo, "네이버 지역검색") <> vbYes Then Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Dim wsOut As Worksheet
    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = SPSS_UniqueSheetName(wb, "네이버_신규리스트")

    Dim headers As Variant
    headers = Array("업종", basisLabel, "검색어", "검색순위", "업체명", "업체유형", "지번주소", "도로명주소", "URL", "검색일시")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        wsOut.Cells(1, i + 1).Value = headers(i)
    Next i

    Dim rowOut As Long: rowOut = 2
    Dim requestCount As Long
    Dim resultCount As Long
    Dim failCount As Long
    Dim dongName As Variant
    Dim queryText As String
    Dim items As Collection
    Dim item As Object

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_ProgressShow "네이버 신규 리스트", "검색 준비 중", dongNames.Count

    For Each dongName In dongNames
        If Len(Trim$(CStr(dongName))) > 0 Then
            requestCount = requestCount + 1
            Toolkit_ProgressUpdate requestCount, dongNames.Count, "검색 중: " & Trim$(CStr(dongName))
            queryText = Trim$(CStr(dongName)) & " " & CStr(options("industry"))

            Set items = Nothing
            On Error Resume Next
            Set items = NaverLocal_Search(queryText, CLng(options("displayCount")), CStr(options("sortMode")))
            If Err.Number <> 0 Then
                failCount = failCount + 1
                Toolkit_LogApiEvent wb, "네이버 신규 리스트", "search", queryText, "ERROR", Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler

            If Not items Is Nothing Then
                For i = 1 To items.Count
                    Set item = items(i)
                    wsOut.Cells(rowOut, 1).Value = CStr(options("industry"))
                    wsOut.Cells(rowOut, 2).Value = Trim$(CStr(dongName))
                    wsOut.Cells(rowOut, 3).Value = queryText
                    wsOut.Cells(rowOut, 4).Value = i
                    wsOut.Cells(rowOut, 5).Value = item("title")
                    wsOut.Cells(rowOut, 6).Value = item("category")
                    wsOut.Cells(rowOut, 7).Value = item("address")
                    wsOut.Cells(rowOut, 8).Value = item("roadAddress")
                    wsOut.Cells(rowOut, 9).Value = item("link")
                    wsOut.Cells(rowOut, 10).Value = Now
                    rowOut = rowOut + 1
                    resultCount = resultCount + 1
                Next i
            End If

            NaverLocal_Delay CDbl(options("delaySeconds"))
        End If
    Next dongName

    wsOut.Range("A1:J1").Font.Bold = True
    wsOut.Columns("A:J").AutoFit
    Toolkit_ProgressClose
    Application.ScreenUpdating = True

    MsgBox "네이버 지역검색 신규 리스트 생성을 완료했습니다." & vbCrLf & _
           "요청 " & basisLabel & ": " & requestCount & "건" & vbCrLf & _
           "생성 결과: " & resultCount & "건" & vbCrLf & _
           "조회 오류: " & failCount & "건", vbInformation, "네이버 지역검색"
    Exit Sub

ErrHandler:
    Toolkit_ProgressClose
    Application.ScreenUpdating = True
    MsgBox "네이버 지역검색 신규 리스트 생성 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical, "네이버 지역검색"
End Sub

Public Sub NaverLocal_CreateDongMasterTemplate()
    NaverLocal_CreateMasterTemplate NAVER_USER_DONG_MASTER_SHEET, NaverLocal_GetEmbeddedDongMasterSheet(), "행정동"
End Sub

Public Sub NaverLocal_CreateLegalDongMasterTemplate()
    NaverLocal_CreateMasterTemplate NAVER_USER_LEGAL_DONG_MASTER_SHEET, NaverLocal_GetEmbeddedLegalDongMasterSheet(), "법정동"
End Sub

Private Sub NaverLocal_CreateMasterTemplate(ByVal targetSheetName As String, ByVal sourceWs As Worksheet, ByVal basisLabel As String)
    If ActiveWorkbook Is Nothing Then Exit Sub
    If sourceWs Is Nothing Then
        MsgBox "추가기능 내부 " & basisLabel & " 마스터를 찾지 못했습니다.", vbExclamation, basisLabel & " 마스터 양식"
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim existing As Worksheet
    On Error Resume Next
    Set existing = wb.Worksheets(targetSheetName)
    On Error GoTo 0

    If Not existing Is Nothing Then
        If MsgBox("현재 통합문서에 '" & targetSheetName & "' 시트가 이미 있습니다." & vbCrLf & _
                  "기존 시트를 지우고 내장 " & basisLabel & " 마스터 양식으로 다시 만들까요?", _
                  vbQuestion + vbYesNo, basisLabel & " 마스터 양식") <> vbYes Then Exit Sub
        Application.DisplayAlerts = False
        existing.Delete
        Application.DisplayAlerts = True
    End If

    Dim outWs As Worksheet
    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = targetSheetName

    sourceWs.UsedRange.Copy Destination:=outWs.Range("A1")
    Application.CutCopyMode = False
    outWs.UsedRange.Columns.AutoFit
    outWs.Rows(1).Font.Bold = True
    outWs.Activate

    MsgBox "'" & targetSheetName & "' 시트를 만들었습니다." & vbCrLf & _
           "컬럼 구조를 유지한 채 행을 추가/삭제하거나 사용여부를 Y/N으로 바꾸면 " & basisLabel & " 기준 신규 리스트 생성에 반영됩니다.", _
           vbInformation, basisLabel & " 마스터 양식"
End Sub

Public Sub NaverLocal_VerifyExistingData()
    If Not NaverLocal_HasCredentials() Then
        MsgBox "먼저 [네이버 인증키 설정] 버튼으로 Client ID와 Client Secret을 저장하세요.", vbExclamation, "네이버 지역검색"
        Exit Sub
    End If

    Dim options As Object
    Set options = NaverLocal_GetVerifyOptions()
    If options Is Nothing Then Exit Sub

    Dim nameRange As Range
    Dim addressRange As Range
    If CBool(options("useName")) Then Set nameRange = options("nameRange")
    If CBool(options("useAddress")) Then Set addressRange = options("addressRange")

    If nameRange Is Nothing And addressRange Is Nothing Then
        MsgBox "업체명 또는 주소 범위 중 하나 이상을 선택해야 합니다.", vbExclamation, "네이버 지역검색"
        Exit Sub
    End If

    If Not nameRange Is Nothing Then
        If nameRange.Areas.Count > 1 Then
            MsgBox "업체명 범위는 한 덩어리의 연속 범위만 지원합니다.", vbExclamation, "네이버 지역검색"
            Exit Sub
        End If
        If nameRange.Columns.Count > 1 Then Set nameRange = nameRange.Columns(1)
    End If

    If Not addressRange Is Nothing Then
        If addressRange.Areas.Count > 1 Then
            MsgBox "주소 범위는 한 덩어리의 연속 범위만 지원합니다.", vbExclamation, "네이버 지역검색"
            Exit Sub
        End If
        If addressRange.Columns.Count > 1 Then Set addressRange = addressRange.Columns(1)
    End If

    Dim ws As Worksheet
    If Not nameRange Is Nothing Then
        Set ws = nameRange.Worksheet
    Else
        Set ws = addressRange.Worksheet
    End If

    If Not nameRange Is Nothing And Not addressRange Is Nothing Then
        If Not addressRange.Worksheet Is ws Then
            MsgBox "업체명과 주소 범위는 같은 시트에 있어야 합니다.", vbExclamation, "네이버 지역검색"
            Exit Sub
        End If
        If nameRange.Rows.Count <> addressRange.Rows.Count Then
            MsgBox "업체명 범위와 주소 범위의 행 수가 다릅니다.", vbExclamation, "네이버 지역검색"
            Exit Sub
        End If
    End If

    Dim baseRange As Range
    If Not nameRange Is Nothing Then Set baseRange = nameRange Else Set baseRange = addressRange

    Dim firstRow As Long
    Dim startOffset As Long
    startOffset = IIf(CBool(options("hasHeader")), 1, 0)
    firstRow = baseRange.Row + startOffset
    If firstRow > baseRange.Row + baseRange.Rows.Count - 1 Then
        MsgBox "헤더를 제외하면 검증할 데이터가 없습니다.", vbInformation, "네이버 지역검색"
        Exit Sub
    End If

    Dim outputCol As Long
    If Not nameRange Is Nothing And Not addressRange Is Nothing Then
        outputCol = Application.Max(nameRange.Column + nameRange.Columns.Count, addressRange.Column + addressRange.Columns.Count)
    ElseIf Not nameRange Is Nothing Then
        outputCol = nameRange.Column + nameRange.Columns.Count
    Else
        outputCol = addressRange.Column + addressRange.Columns.Count
    End If
    If Not Toolkit_CanInsertColumns(ws, outputCol, 8, "네이버 주소 검증") Then Exit Sub
    ws.Columns(outputCol).Resize(, 8).Insert Shift:=xlToRight
    NaverLocal_WriteVerifyHeaders ws, baseRange.Row, outputCol

    Dim rowIndex As Long
    Dim nameValue As String
    Dim addressValue As String
    Dim queryText As String
    Dim items As Collection
    Dim bestItem As Object
    Dim nameScore As Double
    Dim addressScore As Double
    Dim riskLabel As String
    Dim requestCount As Long
    Dim noResultCount As Long
    Dim errorCount As Long
    Dim totalRows As Long
    totalRows = baseRange.Row + baseRange.Rows.Count - firstRow

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_ProgressShow "네이버 주소 검증", "검증 준비 중", totalRows

    For rowIndex = firstRow To baseRange.Row + baseRange.Rows.Count - 1
        Toolkit_ProgressUpdate rowIndex - firstRow + 1, totalRows, "검증 중: " & CStr(rowIndex) & "행"
        If Not nameRange Is Nothing Then nameValue = Trim$(CStr(ws.Cells(rowIndex, nameRange.Column).Value)) Else nameValue = ""
        If Not addressRange Is Nothing Then addressValue = Trim$(CStr(ws.Cells(rowIndex, addressRange.Column).Value)) Else addressValue = ""

        If Len(nameValue) > 0 Or Len(addressValue) > 0 Then
            requestCount = requestCount + 1
            queryText = NaverLocal_BuildVerifyQuery(nameValue, addressValue, CStr(options("queryMode")))

            Set items = Nothing
            Set bestItem = Nothing
            nameScore = 0
            addressScore = 0
            On Error Resume Next
            Set items = NaverLocal_Search(queryText, CLng(options("displayCount")), CStr(options("sortMode")))
            If Err.Number <> 0 Then
                errorCount = errorCount + 1
                riskLabel = "오류 위험성 높음 확인 필요"
                ws.Cells(rowIndex, outputCol + 7).Value = riskLabel
                Toolkit_LogApiEvent ws.Parent, "네이버 주소 검증", "search", queryText, "ERROR", Err.Description
                Err.Clear
                On Error GoTo ErrHandler
                NaverLocal_ColorRiskCell ws.Cells(rowIndex, outputCol + 7), riskLabel
                NaverLocal_Delay CDbl(options("delaySeconds"))
                GoTo NextVerifyRow
            End If
            On Error GoTo ErrHandler

            If Not items Is Nothing Then Set bestItem = NaverLocal_BestMatch(items, nameValue, addressValue, nameScore, addressScore)

            If bestItem Is Nothing Then
                noResultCount = noResultCount + 1
                riskLabel = "오류 위험성 높음 확인 필요"
                ws.Cells(rowIndex, outputCol + 7).Value = riskLabel
            Else
                riskLabel = NaverLocal_RiskLabelForMode(nameScore, addressScore, CStr(options("queryMode")))
                ws.Cells(rowIndex, outputCol).Value = bestItem("title")
                ws.Cells(rowIndex, outputCol + 1).Value = bestItem("category")
                ws.Cells(rowIndex, outputCol + 2).Value = bestItem("address")
                ws.Cells(rowIndex, outputCol + 3).Value = bestItem("roadAddress")
                ws.Cells(rowIndex, outputCol + 4).Value = bestItem("link")
                ws.Cells(rowIndex, outputCol + 5).Value = Round(nameScore, 1)
                ws.Cells(rowIndex, outputCol + 6).Value = Round(addressScore, 1)
                ws.Cells(rowIndex, outputCol + 7).Value = riskLabel
            End If

            NaverLocal_ColorRiskCell ws.Cells(rowIndex, outputCol + 7), riskLabel
            NaverLocal_Delay CDbl(options("delaySeconds"))
        End If
NextVerifyRow:
    Next rowIndex

    ws.Range(ws.Cells(baseRange.Row, outputCol), ws.Cells(baseRange.Row, outputCol + 7)).Font.Bold = True
    ws.Columns(outputCol).Resize(, 8).AutoFit
    Toolkit_ProgressClose
    Application.ScreenUpdating = True

    MsgBox "네이버 지역검색 기반 주소 검증을 완료했습니다." & vbCrLf & _
           "조회: " & requestCount & "건" & vbCrLf & _
           "검색결과 없음: " & noResultCount & "건" & vbCrLf & _
           "조회 오류: " & errorCount & "건", vbInformation, "네이버 지역검색"
    Exit Sub

ErrHandler:
    Toolkit_ProgressClose
    Application.ScreenUpdating = True
    MsgBox "네이버 주소 검증 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical, "네이버 지역검색"
End Sub

Private Function NaverLocal_GetBuildOptions() As Object
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmNaverLocalBuildOptions")
    form.Show
    If form.Cancelled Then
        Unload form
        Exit Function
    End If

    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    options("industry") = CStr(form.IndustryText)
    options("displayCount") = CLng(form.DisplayCount)
    options("delaySeconds") = CDbl(form.DelaySeconds)
    options("sortMode") = CStr(form.SortMode)
    options("sidoFilter") = CStr(form.SidoFilter)
    options("sigunguFilter") = CStr(form.SigunguFilter)
    options("dongBasis") = CStr(form.DongBasis)
    Unload form
    Set NaverLocal_GetBuildOptions = options
    Exit Function

Fallback:
    Dim industry As String
    industry = InputBox("검색할 업종/키워드를 입력하세요. 예: 서점, 카페, 음식점", "네이버 신규 리스트")
    If Len(Trim$(industry)) = 0 Then Exit Function
    Set options = CreateObject("Scripting.Dictionary")
    options("industry") = Trim$(industry)
    options("displayCount") = 5
    options("delaySeconds") = 0.4
    options("sortMode") = "comment"
    options("sidoFilter") = ""
    options("sigunguFilter") = ""
    options("dongBasis") = "admin"
    Set NaverLocal_GetBuildOptions = options
End Function

Private Function NaverLocal_GetDongNamesForSearch(ByVal wb As Workbook, ByVal options As Object) As Collection
    Dim result As New Collection
    Dim ws As Worksheet
    If NaverLocal_DongBasis(options) = "legal" Then
        Set ws = NaverLocal_GetUserLegalDongMasterSheet(wb)
        If ws Is Nothing Then Set ws = NaverLocal_GetEmbeddedLegalDongMasterSheet()
    Else
        Set ws = NaverLocal_GetUserDongMasterSheet(wb)
        If ws Is Nothing Then Set ws = NaverLocal_GetEmbeddedDongMasterSheet()
    End If
    If ws Is Nothing Then
        Set NaverLocal_GetDongNamesForSearch = result
        Exit Function
    End If

    Dim headerRow As Long: headerRow = 1
    Dim sidoCol As Long, sigunguCol As Long, searchCol As Long, useCol As Long
    sidoCol = NaverLocal_FindHeaderColumn(ws, headerRow, "광역시도")
    If sidoCol = 0 Then sidoCol = NaverLocal_FindHeaderColumn(ws, headerRow, "시도명")
    sigunguCol = NaverLocal_FindHeaderColumn(ws, headerRow, "시군구")
    If sigunguCol = 0 Then sigunguCol = NaverLocal_FindHeaderColumn(ws, headerRow, "시군구명")
    searchCol = NaverLocal_FindHeaderColumn(ws, headerRow, "검색주소")
    If searchCol = 0 Then searchCol = NaverLocal_FindHeaderColumn(ws, headerRow, "원본")
    If searchCol = 0 Then searchCol = NaverLocal_FindHeaderColumn(ws, headerRow, "행정동전체명")
    If searchCol = 0 Then searchCol = NaverLocal_FindHeaderColumn(ws, headerRow, "검색어")
    useCol = NaverLocal_FindHeaderColumn(ws, headerRow, "사용여부")

    If searchCol = 0 Then
        Set NaverLocal_GetDongNamesForSearch = result
        Exit Function
    End If

    Dim sidoFilter As String, sigunguFilter As String
    sidoFilter = Trim$(CStr(options("sidoFilter")))
    sigunguFilter = Trim$(CStr(options("sigunguFilter")))

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, searchCol).End(xlUp).Row

    Dim r As Long
    For r = headerRow + 1 To lastRow
        Dim searchText As String
        searchText = Trim$(CStr(ws.Cells(r, searchCol).Value))
        If Len(searchText) > 0 Then
            Dim useValue As String
            If useCol > 0 Then useValue = UCase$(Trim$(CStr(ws.Cells(r, useCol).Value))) Else useValue = "Y"
            If useValue <> "N" And useValue <> "NO" And useValue <> "0" And useValue <> "FALSE" Then
                If NaverLocal_MasterRowMatches(ws, r, sidoCol, sigunguCol, sidoFilter, sigunguFilter) Then
                    result.Add searchText
                End If
            End If
        End If
    Next r

    Set NaverLocal_GetDongNamesForSearch = result
End Function

Private Function NaverLocal_MasterRowMatches(ws As Worksheet, ByVal rowIndex As Long, ByVal sidoCol As Long, ByVal sigunguCol As Long, ByVal sidoFilter As String, ByVal sigunguFilter As String) As Boolean
    NaverLocal_MasterRowMatches = True
    If Len(sidoFilter) > 0 Then
        If sidoCol = 0 Then
            NaverLocal_MasterRowMatches = False
            Exit Function
        End If
        If InStr(1, CStr(ws.Cells(rowIndex, sidoCol).Value), sidoFilter, vbTextCompare) = 0 Then
            NaverLocal_MasterRowMatches = False
            Exit Function
        End If
    End If

    If Len(sigunguFilter) > 0 Then
        If sigunguCol = 0 Then
            NaverLocal_MasterRowMatches = False
            Exit Function
        End If
        If InStr(1, CStr(ws.Cells(rowIndex, sigunguCol).Value), sigunguFilter, vbTextCompare) = 0 Then
            NaverLocal_MasterRowMatches = False
            Exit Function
        End If
    End If
End Function

Private Function NaverLocal_GetUserDongMasterSheet(ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set NaverLocal_GetUserDongMasterSheet = wb.Worksheets(NAVER_USER_DONG_MASTER_SHEET)
    On Error GoTo 0
End Function

Private Function NaverLocal_GetEmbeddedDongMasterSheet() As Worksheet
    On Error Resume Next
    Set NaverLocal_GetEmbeddedDongMasterSheet = ThisWorkbook.Worksheets(NAVER_DONG_MASTER_SHEET)
    On Error GoTo 0
End Function

Private Function NaverLocal_GetUserLegalDongMasterSheet(ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set NaverLocal_GetUserLegalDongMasterSheet = wb.Worksheets(NAVER_USER_LEGAL_DONG_MASTER_SHEET)
    On Error GoTo 0
End Function

Private Function NaverLocal_GetEmbeddedLegalDongMasterSheet() As Worksheet
    On Error Resume Next
    Set NaverLocal_GetEmbeddedLegalDongMasterSheet = ThisWorkbook.Worksheets(NAVER_LEGAL_DONG_MASTER_SHEET)
    On Error GoTo 0
End Function

Private Function NaverLocal_FindHeaderColumn(ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If Replace(Trim$(CStr(ws.Cells(headerRow, c).Value)), " ", "") = Replace(headerText, " ", "") Then
            NaverLocal_FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function NaverLocal_GetVerifyOptions() As Object
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmNaverLocalVerifyOptions")
    form.Show
    If form.Cancelled Then
        Unload form
        Exit Function
    End If

    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    options("queryMode") = CStr(form.QueryMode)
    options("displayCount") = CLng(form.DisplayCount)
    options("delaySeconds") = CDbl(form.DelaySeconds)
    options("sortMode") = CStr(form.SortMode)
    options("hasHeader") = CBool(form.HasHeader)
    options("useName") = CBool(form.UseName)
    options("useAddress") = CBool(form.UseAddress)
    If CBool(form.UseName) Then Set options("nameRange") = form.NameRange
    If CBool(form.UseAddress) Then Set options("addressRange") = form.AddressRange
    Unload form
    Set NaverLocal_GetVerifyOptions = options
    Exit Function

Fallback:
    MsgBox "주소 검증 옵션창을 열 수 없습니다." & vbCrLf & Err.Description, vbCritical, "네이버 지역검색"
End Function

Private Function NaverLocal_HasCredentials() As Boolean
    NaverLocal_HasCredentials = (Len(NaverLocal_GetClientId()) > 0 And Len(NaverLocal_GetClientSecret()) > 0)
End Function

Private Function NaverLocal_GetClientId() As String
    NaverLocal_GetClientId = GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_CLIENT_ID_KEY, "")
End Function

Private Function NaverLocal_GetClientSecret() As String
    NaverLocal_GetClientSecret = GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_CLIENT_SECRET_KEY, "")
End Function

Private Function NaverLocal_Search(ByVal queryText As String, ByVal displayCount As Long, ByVal sortMode As String) As Collection
    If displayCount < 1 Then displayCount = 1
    If displayCount > 5 Then displayCount = 5
    If sortMode <> "random" Then sortMode = "comment"

    Static searchCache As Object
    If searchCache Is Nothing Then Set searchCache = CreateObject("Scripting.Dictionary")

    Dim cacheKey As String
    cacheKey = NaverLocal_GetClientId() & "|" & queryText & "|" & CStr(displayCount) & "|" & sortMode
    If searchCache.Exists(cacheKey) Then
        Set NaverLocal_Search = searchCache(cacheKey)
        Exit Function
    End If

    Dim url As String
    url = NAVER_API_BASE_URL & "?query=" & NaverLocal_UrlEncode(queryText) & _
          "&display=" & CStr(displayCount) & "&start=1&sort=" & sortMode

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts API_RESOLVE_TIMEOUT_MS, API_CONNECT_TIMEOUT_MS, API_SEND_TIMEOUT_MS, API_RECEIVE_TIMEOUT_MS
    http.Open "GET", url, False
    http.setRequestHeader "X-Naver-Client-Id", NaverLocal_GetClientId()
    http.setRequestHeader "X-Naver-Client-Secret", NaverLocal_GetClientSecret()
    http.setRequestHeader "Accept", "application/json"
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 781, "NaverLocal_Search", "HTTP " & http.Status & ": " & http.responseText
    End If

    Set NaverLocal_Search = NaverLocal_ExtractItems(CStr(http.responseText))
    searchCache.Add cacheKey, NaverLocal_Search
End Function

Private Function NaverLocal_DongBasis(ByVal options As Object) As String
    On Error Resume Next
    NaverLocal_DongBasis = LCase$(Trim$(CStr(options("dongBasis"))))
    On Error GoTo 0
    If NaverLocal_DongBasis <> "legal" Then NaverLocal_DongBasis = "admin"
End Function

Private Function NaverLocal_DongBasisLabel(ByVal options As Object) As String
    If NaverLocal_DongBasis(options) = "legal" Then
        NaverLocal_DongBasisLabel = "법정동"
    Else
        NaverLocal_DongBasisLabel = "행정동"
    End If
End Function

Private Function NaverLocal_ExtractItems(ByVal jsonText As String) As Collection
    Dim result As New Collection
    Dim itemObjects As Collection
    Set itemObjects = NaverLocal_ExtractArrayObjects(jsonText, "items")

    Dim itemJson As Variant
    Dim item As Object
    For Each itemJson In itemObjects
        Set item = CreateObject("Scripting.Dictionary")
        item("title") = NaverLocal_StripHtml(Biz_JsonString(CStr(itemJson), "title"))
        item("link") = Biz_JsonString(CStr(itemJson), "link")
        item("category") = Biz_JsonString(CStr(itemJson), "category")
        item("address") = Biz_JsonString(CStr(itemJson), "address")
        item("roadAddress") = Biz_JsonString(CStr(itemJson), "roadAddress")
        result.Add item
    Next itemJson

    Set NaverLocal_ExtractItems = result
End Function

Private Function NaverLocal_ExtractArrayObjects(ByVal jsonText As String, ByVal arrayKey As String) As Collection
    Dim result As New Collection
    Dim keyPos As Long
    Dim arrStart As Long
    Dim i As Long
    Dim depth As Long
    Dim objStart As Long
    Dim inString As Boolean
    Dim ch As String
    Dim prevCh As String

    keyPos = InStr(1, jsonText, """" & arrayKey & """", vbTextCompare)
    If keyPos = 0 Then
        Set NaverLocal_ExtractArrayObjects = result
        Exit Function
    End If

    arrStart = InStr(keyPos, jsonText, "[")
    If arrStart = 0 Then
        Set NaverLocal_ExtractArrayObjects = result
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

    Set NaverLocal_ExtractArrayObjects = result
End Function

Private Function NaverLocal_UrlEncode(ByVal textValue As String) As String
    On Error Resume Next
    NaverLocal_UrlEncode = Application.WorksheetFunction.EncodeURL(textValue)
    If Err.Number <> 0 Or Len(NaverLocal_UrlEncode) = 0 Then
        Err.Clear
        NaverLocal_UrlEncode = Replace(textValue, " ", "%20")
    End If
    On Error GoTo 0
End Function

Private Function NaverLocal_StripHtml(ByVal textValue As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "<[^>]+>"
    NaverLocal_StripHtml = Trim$(re.Replace(textValue, ""))
End Function

Private Function NaverLocal_BuildVerifyQuery(ByVal nameValue As String, ByVal addressValue As String, ByVal queryMode As String) As String
    Select Case queryMode
        Case "name_only"
            NaverLocal_BuildVerifyQuery = nameValue
        Case "address_only"
            NaverLocal_BuildVerifyQuery = addressValue
        Case Else
            NaverLocal_BuildVerifyQuery = Trim$(nameValue & " " & addressValue)
    End Select
End Function

Private Function NaverLocal_BestMatch(ByVal items As Collection, ByVal nameValue As String, ByVal addressValue As String, ByRef bestNameScore As Double, ByRef bestAddressScore As Double) As Object
    Dim item As Object
    Dim nameScore As Double
    Dim addressScore As Double
    Dim combinedScore As Double
    Dim bestScore As Double

    If items Is Nothing Then Exit Function
    For Each item In items
        nameScore = NaverLocal_Similarity(nameValue, CStr(item("title")))
        addressScore = Application.Max( _
            NaverLocal_Similarity(addressValue, CStr(item("roadAddress"))), _
            NaverLocal_Similarity(addressValue, CStr(item("address"))))
        combinedScore = (nameScore * 0.45) + (addressScore * 0.55)
        If combinedScore > bestScore Then
            bestScore = combinedScore
            bestNameScore = nameScore
            bestAddressScore = addressScore
            Set NaverLocal_BestMatch = item
        End If
    Next item
End Function

Private Function NaverLocal_RiskLabel(ByVal nameScore As Double, ByVal addressScore As Double) As String
    If nameScore >= 85 And addressScore >= 85 Then
        NaverLocal_RiskLabel = "정상"
    ElseIf nameScore >= 75 And addressScore >= 70 Then
        NaverLocal_RiskLabel = "오류 위험성 낮음"
    ElseIf nameScore >= 55 Or addressScore >= 55 Then
        NaverLocal_RiskLabel = "오류 위험성 보통"
    Else
        NaverLocal_RiskLabel = "오류 위험성 높음 확인 필요"
    End If
End Function

Private Function NaverLocal_RiskLabelForMode(ByVal nameScore As Double, ByVal addressScore As Double, ByVal queryMode As String) As String
    Select Case queryMode
        Case "name_only"
            NaverLocal_RiskLabelForMode = NaverLocal_RiskLabelSingle(nameScore)
        Case "address_only"
            NaverLocal_RiskLabelForMode = NaverLocal_RiskLabelSingle(addressScore)
        Case Else
            NaverLocal_RiskLabelForMode = NaverLocal_RiskLabel(nameScore, addressScore)
    End Select
End Function

Private Function NaverLocal_RiskLabelSingle(ByVal score As Double) As String
    If score >= 85 Then
        NaverLocal_RiskLabelSingle = "정상"
    ElseIf score >= 75 Then
        NaverLocal_RiskLabelSingle = "오류 위험성 낮음"
    ElseIf score >= 55 Then
        NaverLocal_RiskLabelSingle = "오류 위험성 보통"
    Else
        NaverLocal_RiskLabelSingle = "오류 위험성 높음 확인 필요"
    End If
End Function

Private Sub NaverLocal_WriteVerifyHeaders(ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long)
    Dim headers As Variant
    headers = Array("검색결과 업체명", "검색결과 업체유형", "검색결과 지번주소", "검색결과 도로명주소", "검색결과 URL", "업체명 유사도", "주소 유사도", "검증등급")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowIndex, colIndex + i).Value = headers(i)
    Next i
End Sub

Private Sub NaverLocal_ColorRiskCell(ByVal cell As Range, ByVal riskLabel As String)
    Select Case riskLabel
        Case "정상"
            cell.Interior.Color = RGB(198, 239, 206)
        Case "오류 위험성 낮음"
            cell.Interior.Color = RGB(255, 235, 156)
        Case "오류 위험성 보통"
            cell.Interior.Color = RGB(244, 176, 132)
        Case Else
            cell.Interior.Color = RGB(255, 199, 206)
    End Select
End Sub

Private Function NaverLocal_Similarity(ByVal leftText As String, ByVal rightText As String) As Double
    Dim a As String
    Dim b As String
    a = NaverLocal_NormalizeForMatch(leftText)
    b = NaverLocal_NormalizeForMatch(rightText)

    If Len(a) = 0 And Len(b) = 0 Then
        NaverLocal_Similarity = 100
    ElseIf Len(a) = 0 Or Len(b) = 0 Then
        NaverLocal_Similarity = 0
    Else
        NaverLocal_Similarity = 100# * (1# - (NaverLocal_Levenshtein(a, b) / Application.Max(Len(a), Len(b))))
    End If
End Function

Private Function NaverLocal_NormalizeForMatch(ByVal textValue As String) As String
    textValue = LCase$(NaverLocal_StripHtml(CStr(textValue)))
    textValue = Replace(textValue, "주식회사", "")
    textValue = Replace(textValue, "(주)", "")
    textValue = Replace(textValue, "㈜", "")
    textValue = Replace(textValue, " ", "")
    textValue = Replace(textValue, vbTab, "")
    textValue = Replace(textValue, "-", "")
    textValue = Replace(textValue, "(", "")
    textValue = Replace(textValue, ")", "")
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, ".", "")
    NaverLocal_NormalizeForMatch = textValue
End Function

Private Function NaverLocal_Levenshtein(ByVal a As String, ByVal b As String) As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim d() As Long

    ReDim d(0 To Len(a), 0 To Len(b))
    For i = 0 To Len(a)
        d(i, 0) = i
    Next i
    For j = 0 To Len(b)
        d(0, j) = j
    Next j

    For i = 1 To Len(a)
        For j = 1 To Len(b)
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then cost = 0 Else cost = 1
            d(i, j) = Application.Min(Application.Min(d(i - 1, j) + 1, d(i, j - 1) + 1), d(i - 1, j - 1) + cost)
        Next j
    Next i

    NaverLocal_Levenshtein = d(Len(a), Len(b))
End Function

Private Sub NaverLocal_Delay(ByVal seconds As Double)
    If seconds <= 0 Then Exit Sub
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

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
    MsgBox "이상치 탐지는 이제 설정 시트를 만들지 않고 바로 범위를 선택하는 방식으로 통합되었습니다." & vbCrLf & vbCrLf & _
           "[데이터 툴킷] 탭의 [분석 열 선택] 버튼 또는 OutlierDetect_Run을 실행하세요.", _
           vbInformation, "이상치 탐지"
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

Private Function SPSS_ParseIndexSelection(ByVal inputText As String, ByVal maxIndex As Long) As Object
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

Private Function SPSS_UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
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

    Application.ScreenUpdating = False
    For r1 = dataRange.Row + 1 To lastRow - 1
        For r2 = r1 + 1 To lastRow
            score = Roster_DuplicateScore(dataRange.Worksheet, r1, r2, dataRange, nameCol, addressCol, phoneCol, bizCol, reason)
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
    Next r1
    outWs.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "중복 업체 탐지를 완료했습니다." & vbCrLf & "의심 쌍: " & maxPairs & "건", vbInformation, "중복 업체 탐지"
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

    outRow = Roster_WriteMissingSummary(outWs, outRow, ws, dataRange, nameCol, "업체명 빈값")
    outRow = Roster_WriteMissingSummary(outWs, outRow, ws, dataRange, addressCol, "주소 빈값")
    outRow = Roster_WriteMissingSummary(outWs, outRow, ws, dataRange, phoneCol, "전화번호 빈값")
    outRow = Roster_WriteMissingSummary(outWs, outRow, ws, dataRange, bizCol, "사업자번호 빈값")

    If phoneCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 확인 필요", Roster_CountInvalidPhones(ws, dataRange, phoneCol), "숫자 길이/대표번호 패턴 기준"): outRow = outRow + 1
    If bizCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("사업자번호 확인 필요", Roster_CountInvalidBizNo(ws, dataRange, bizCol), "숫자 10자리 여부 기준"): outRow = outRow + 1
    If addressCol > 0 Then
        Roster_WriteArray outWs, outRow, 1, Array("도로명주소", Roster_CountAddressType(ws, dataRange, addressCol, "도로명주소"), "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("법정/행정동 기준 주소", Roster_CountAddressType(ws, dataRange, addressCol, "법정/행정동 기준 주소"), "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("주소 확인 필요", Roster_CountAddressType(ws, dataRange, addressCol, "확인 필요"), "주소 유형 판별 기준"): outRow = outRow + 1
    End If

    outWs.Columns.AutoFit
    MsgBox "명부 품질 진단 리포트를 생성했습니다.", vbInformation, "명부 품질 진단"
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

    If selectedItems.Exists("1") Then Roster_WriteArray outWs, outRow, 1, Array("업체명 빈값", Roster_CountBlankRange(nameRange, hasHeader), "수동 선택 범위 빈값 여부"): outRow = outRow + 1
    If selectedItems.Exists("2") Then Roster_WriteArray outWs, outRow, 1, Array("주소 빈값", Roster_CountBlankRange(addressRange, hasHeader), "수동 선택 범위 빈값 여부"): outRow = outRow + 1
    If selectedItems.Exists("3") Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 확인 필요", Roster_CountInvalidPhonesRange(phoneRange, hasHeader), "숫자 길이/대표번호 패턴 기준"): outRow = outRow + 1
    If selectedItems.Exists("4") Then Roster_WriteArray outWs, outRow, 1, Array("사업자번호 확인 필요", Roster_CountInvalidBizNoRange(bizRange, hasHeader), "숫자 10자리 여부 기준"): outRow = outRow + 1
    If selectedItems.Exists("5") Then
        Roster_WriteArray outWs, outRow, 1, Array("도로명주소", Roster_CountAddressTypeRange(addressRange, hasHeader, "도로명주소"), "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("법정/행정동 기준 주소", Roster_CountAddressTypeRange(addressRange, hasHeader, "법정/행정동 기준 주소"), "주소 유형 판별 기준"): outRow = outRow + 1
        Roster_WriteArray outWs, outRow, 1, Array("주소 확인 필요", Roster_CountAddressTypeRange(addressRange, hasHeader, "확인 필요"), "주소 유형 판별 기준"): outRow = outRow + 1
    End If

    outWs.Columns.AutoFit
    MsgBox "수동 선택 기반 명부 품질 진단 리포트를 생성했습니다." & vbCrLf & outWs.Name, vbInformation, "명부 품질 진단"
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

    Application.ScreenUpdating = False
    For r1 = firstDataRow To lastRow - 1
        For r2 = r1 + 1 To lastRow
            score = Roster_DuplicateScoreManual(ws, r1, r2, nameRange, addressRange, phoneRange, bizRange, reason)
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
    Next r1
    outWs.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "수동 선택 기반 중복 업체 탐지를 완료했습니다." & vbCrLf & "의심 쌍: " & maxPairs & "건", vbInformation, "중복 업체 탐지"
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
    ws.Columns(outCol).Resize(, 2).Insert Shift:=xlToRight
    ws.Cells(firstRow, outCol).Value = "업체명 표준화"
    ws.Cells(firstRow, outCol + 1).Value = "유사그룹"

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, key As String, groupNo As Long
    For r = dataStart To lastRow
        key = Roster_NormalizeCompanyName(CStr(ws.Cells(r, workRange.Column).Value))
        ws.Cells(r, outCol).Value = key
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                groupNo = groupNo + 1
                dict(key) = "G" & Format(groupNo, "0000")
            End If
            ws.Cells(r, outCol + 1).Value = dict(key)
        End If
    Next r
    ws.Columns(outCol).Resize(, 2).AutoFit
    MsgBox "업체명 표준화와 유사그룹 생성을 완료했습니다.", vbInformation, "업체명 유사도 매칭"
End Sub

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
    ws.Columns(outCol).Resize(, 4).Insert Shift:=xlToRight
    Roster_WriteArray ws, firstRow, outCol, Array("광역시도", "시군구", "읍면동", "상세주소")

    Dim r As Long, parts As Variant
    For r = dataStart To lastRow
        parts = Address_ParseAdministrativeParts(CStr(ws.Cells(r, addressRange.Column).Value))
        Roster_WriteArray ws, r, outCol, parts
    Next r
    ws.Columns(outCol).Resize(, 4).AutoFit
    MsgBox "행정구역 분해를 완료했습니다.", vbInformation, "행정구역 분해"
End Sub

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
    ws.Columns(outCol).Resize(, 2).Insert Shift:=xlToRight
    Roster_WriteArray ws, firstRow, outCol, Array("전화번호 유형", "전화번호 유효성")

    Dim r As Long, result As Variant
    For r = dataStart To lastRow
        result = Phone_ClassifyOne(CStr(ws.Cells(r, phoneRange.Column).Value))
        Roster_WriteArray ws, r, outCol, result
    Next r
    ws.Columns(outCol).Resize(, 2).AutoFit
    MsgBox "전화번호 유효성 검사를 완료했습니다.", vbInformation, "전화번호 유효성 검사"
End Sub

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
        p1 = DigitsOnly(Roster_CellText(ws, r1, dataRange, phoneCol))
        p2 = DigitsOnly(Roster_CellText(ws, r2, dataRange, phoneCol))
        If Len(p1) >= 8 And p1 = p2 Then score = score + 35: reason = reason & "전화번호 일치; "
    End If
    If nameCol > 0 Then
        Dim ns As Double
        ns = NaverLocal_Similarity(Roster_CellText(ws, r1, dataRange, nameCol), Roster_CellText(ws, r2, dataRange, nameCol))
        If ns >= 85 Then score = score + 35: reason = reason & "업체명 유사; "
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

Private Sub Roster_WriteArray(ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal values As Variant)
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
        p1 = DigitsOnly(Roster_RangeTextAtRow(phoneRange, r1))
        p2 = DigitsOnly(Roster_RangeTextAtRow(phoneRange, r2))
        If Len(p1) >= 7 And p1 = p2 Then score = score + 45: reason = reason & "전화번호 일치; "
    End If

    If Not nameRange Is Nothing Then
        ns = NaverLocal_Similarity(Roster_RangeTextAtRow(nameRange, r1), Roster_RangeTextAtRow(nameRange, r2))
        score = score + ns * 0.35
        If ns >= 85 Then reason = reason & "업체명 유사; "
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

Private Function Phone_ClassifyOne(ByVal phoneText As String) As Variant
    Dim digits As String: digits = DigitsOnly(phoneText)
    If Len(digits) = 0 Then
        Phone_ClassifyOne = Array("빈값", "확인 필요")
    ElseIf Left$(digits, 3) = "010" And Len(digits) = 11 Then
        Phone_ClassifyOne = Array("휴대폰", "정상")
    ElseIf Left$(digits, 2) = "02" And (Len(digits) = 9 Or Len(digits) = 10) Then
        Phone_ClassifyOne = Array("서울 지역번호", "정상")
    ElseIf Left$(digits, 2) = "15" Or Left$(digits, 2) = "16" Or Left$(digits, 2) = "18" Then
        If Len(digits) = 8 Then Phone_ClassifyOne = Array("대표번호", "정상") Else Phone_ClassifyOne = Array("대표번호", "확인 필요")
    ElseIf Len(digits) = 10 Or Len(digits) = 11 Then
        Phone_ClassifyOne = Array("지역/일반번호", "정상")
    Else
        Phone_ClassifyOne = Array("알 수 없음", "확인 필요")
    End If
End Function

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

' ============================================================
' 표본설계/표본할당 도구
' ============================================================
Public Sub SampleDesign_CreateTemplate()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "표본설계 시트를 만들 통합문서를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim inputWs As Worksheet
    Dim resultWs As Worksheet
    Dim reviewWs As Worksheet
    Set inputWs = SampleDesign_PrepareSheet(wb, "표본설계_입력", "표본설계 입력 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?")
    If inputWs Is Nothing Then Exit Sub
    Set resultWs = SampleDesign_PrepareSheet(wb, "표본설계_할당결과", "표본설계 할당결과 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?")
    If resultWs Is Nothing Then Exit Sub
    Set reviewWs = SampleDesign_PrepareSheet(wb, "표본설계_검토", "표본설계 검토 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?")
    If reviewWs Is Nothing Then Exit Sub

    SampleDesign_WriteInputTemplate inputWs
    SampleDesign_WriteEmptyResult resultWs
    SampleDesign_WriteEmptyReview reviewWs

    inputWs.Activate
    inputWs.Range("B4").Select
    MsgBox "표본설계 입력/결과/검토 시트를 만들었습니다." & vbCrLf & _
           "목표 표본 수와 모집단 표를 입력한 뒤 [할당 계산]을 실행하세요.", _
           vbInformation, "표본설계"
End Sub

Public Sub SampleDesign_RunAllocation()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "작업할 통합문서를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim inputWs As Worksheet
    Set inputWs = SampleDesign_FindInputSheet(wb)
    If inputWs Is Nothing Then
        MsgBox "'표본설계_입력' 시트를 찾지 못했습니다." & vbCrLf & _
               "먼저 [설계 시트 생성]을 실행하세요.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim targetN As Long
    targetN = CLng(Val(inputWs.Range("B4").Value))
    If targetN <= 0 Then
        MsgBox "목표 표본 수(B4)를 1 이상의 숫자로 입력하세요.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim minN As Long
    Dim maxN As Long
    Dim alpha As Double
    Dim methodName As String
    minN = CLng(Val(inputWs.Range("B5").Value))
    maxN = CLng(Val(inputWs.Range("B6").Value))
    methodName = Trim$(CStr(inputWs.Range("B7").Value))
    alpha = Val(inputWs.Range("B8").Value)
    If alpha <= 0 Then alpha = 0.5

    If Not SampleDesign_SelectMethod(methodName, alpha) Then Exit Sub

    Dim records As Collection
    Dim issues As Collection
    Set records = New Collection
    Set issues = New Collection
    SampleDesign_ReadPopulation inputWs, records, issues

    If records.Count = 0 Then
        MsgBox "계산할 모집단 행을 찾지 못했습니다." & vbCrLf & _
               "표본설계_입력 시트의 14행부터 모집단 정보를 입력하세요.", _
               vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim resultWs As Worksheet
    Dim reviewWs As Worksheet
    Set resultWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_할당결과")
    Set reviewWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_검토")

    Dim allocation As Collection
    Set allocation = SampleDesign_Allocate(records, targetN, methodName, alpha, minN, maxN, issues)

    Application.ScreenUpdating = False
    SampleDesign_WriteAllocationResult resultWs, allocation, targetN, methodName
    SampleDesign_WriteReviewReport reviewWs, allocation, issues, targetN, methodName, alpha
    Application.ScreenUpdating = True

    resultWs.Activate
    MsgBox "표본할당 계산을 완료했습니다." & vbCrLf & _
           "배분 방식: " & methodName & vbCrLf & _
           "목표 표본 수: " & targetN & vbCrLf & _
           "최종 배정 표본 수: " & SampleDesign_TotalFinalN(allocation), _
           vbInformation, "표본설계"
End Sub

Public Sub SampleDesign_CreateProgressTracker()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Dim resultWs As Worksheet
    Set resultWs = SampleDesign_FindSheet(wb, "표본설계_할당결과")
    If resultWs Is Nothing Then
        MsgBox "먼저 [할당 계산]을 실행해 표본설계_할당결과 시트를 만드세요.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim outWs As Worksheet
    Set outWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_진행관리")
    outWs.Range("A1").Value = "표본설계 진행관리"
    outWs.Range("A1").Font.Bold = True
    outWs.Range("A1").Font.Size = 14
    SampleDesign_WriteArray outWs, 3, 1, Array("쿼터ID", "지역", "업종", "종사자규모", "성별", "연령", "최종할당", "회수", "부족", "초과", "달성률", "상태")

    Dim lastRow As Long
    lastRow = resultWs.Cells(resultWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "할당결과 시트에 진행관리로 옮길 데이터가 없습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim r As Long
    Dim outRow As Long
    outRow = 4
    For r = 4 To lastRow
        If Trim$(CStr(resultWs.Cells(r, 1).Value)) <> "총계" And Len(Trim$(CStr(resultWs.Cells(r, 1).Value))) > 0 Then
            outWs.Cells(outRow, 1).Resize(1, 6).Value = resultWs.Cells(r, 1).Resize(1, 6).Value
            outWs.Cells(outRow, 7).Value = resultWs.Cells(r, 11).Value
            outWs.Cells(outRow, 8).Value = 0
            outWs.Cells(outRow, 9).FormulaR1C1 = "=MAX(RC[-2]-RC[-1],0)"
            outWs.Cells(outRow, 10).FormulaR1C1 = "=MAX(RC[-2]-RC[-3],0)"
            outWs.Cells(outRow, 11).FormulaR1C1 = "=IF(RC[-4]=0,0,RC[-3]/RC[-4])"
            outWs.Cells(outRow, 12).FormulaR1C1 = "=IF(RC[-5]=0,""할당없음"",IF(RC[-4]=0,""미착수"",IF(RC[-2]>0,""초과"",IF(RC[-3]>0,""미달"",""충족""))))"
            outRow = outRow + 1
        End If
    Next r

    With outWs.Range(outWs.Cells(3, 1), outWs.Cells(Application.Max(3, outRow - 1), 12))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    outWs.Range("A3:L3").Font.Bold = True
    outWs.Range("A3:L3").Interior.Color = RGB(220, 235, 250)
    outWs.Range("K:K").NumberFormat = "0.0%"
    SampleDesign_ApplyProgressFormatting outWs, 4, outRow - 1
    outWs.Activate
    MsgBox "표본설계 진행관리 시트를 생성했습니다.", vbInformation, "표본설계"
End Sub

Public Sub SampleDesign_ReviewQuota()
    SampleDesign_RunAllocation
End Sub

Public Sub SampleDesign_ListifyPopulationTable()
    Dim sourceRange As Range
    If Not Toolkit_GetRangeFromUserForm("모집단 구성표 리스트화", _
        "리스트화할 집계표들이 포함된 전체 범위를 선택하세요." & vbCrLf & _
        "표 제목, 가로 헤더, 세로 헤더, 값 영역, 합계 행/열이 함께 포함되어도 됩니다.", _
        sourceRange) Then Exit Sub
    If sourceRange Is Nothing Then Exit Sub
    Set sourceRange = sourceRange.Areas(1)

    If sourceRange.Rows.Count < 3 Or sourceRange.Columns.Count < 3 Then
        MsgBox "집계표 리스트화에는 3행 x 3열 이상의 범위가 필요합니다.", vbExclamation, "모집단 구성표 리스트화"
        Exit Sub
    End If

    Dim blocks As Collection
    Set blocks = PopList_DetectBlocks(sourceRange)
    If blocks.Count = 0 Then
        MsgBox "선택 범위에서 리스트화할 집계표 구조를 찾지 못했습니다." & vbCrLf & _
               "표 제목, 헤더 행, 숫자 값 영역이 함께 포함되도록 범위를 다시 선택하세요.", _
               vbExclamation, "모집단 구성표 리스트화"
        Exit Sub
    End If

    Dim columnAxisName As String
    Dim valueName As String
    Dim outputSheetName As String
    columnAxisName = "가로구분"
    valueName = "값"
    outputSheetName = "모집단_리스트"

    If Not PopList_ShowPreviewForm(blocks, sourceRange, columnAxisName, valueName, outputSheetName) Then Exit Sub

    outputSheetName = PopList_CleanSheetName(outputSheetName, "모집단_리스트")
    Dim outWs As Worksheet
    Set outWs = SampleDesign_PrepareOutputSheet(sourceRange.Worksheet.Parent, outputSheetName)

    Dim writtenRows As Long
    writtenRows = PopList_WriteOutput(blocks, sourceRange, outWs, columnAxisName, valueName)
    outWs.Activate

    MsgBox blocks.Count & "개 집계표에서 " & writtenRows & "개 리스트 행을 생성했습니다." & vbCrLf & _
           "결과 시트: " & outWs.Name, vbInformation, "모집단 구성표 리스트화"
End Sub

Private Function PopList_DetectBlocks(ByVal sourceRange As Range) As Collection
    Dim blocks As Collection
    Set blocks = New Collection

    Dim relRow As Long
    relRow = 1
    Do While relRow < sourceRange.Rows.Count
        Dim block As Object
        Set block = PopList_TryDetectBlock(sourceRange, relRow)
        If Not block Is Nothing Then
            blocks.Add block
            relRow = CLng(block("dataEndRelRow")) + 1
        Else
            relRow = relRow + 1
        End If
    Loop

    Set PopList_DetectBlocks = blocks
End Function

Private Function PopList_TryDetectBlock(ByVal sourceRange As Range, ByVal headerRelRow As Long) As Object
    Dim firstValueCol As Long
    firstValueCol = PopList_FirstValueHeaderCol(sourceRange, headerRelRow)
    If firstValueCol <= 1 Then Exit Function

    Dim rowAxisCount As Long
    rowAxisCount = firstValueCol - 1

    Dim valueCols As Collection
    Set valueCols = PopList_DetectValueColumns(sourceRange, headerRelRow, firstValueCol)
    If valueCols.Count = 0 Then Exit Function

    Dim nonNumericHeaderCount As Long
    nonNumericHeaderCount = PopList_CountNonNumericHeaders(sourceRange, headerRelRow, valueCols)
    If Not PopList_RowAxisLooksLikeHeader(sourceRange, headerRelRow, rowAxisCount) _
       And nonNumericHeaderCount = 0 Then Exit Function

    Dim dataEndRow As Long
    dataEndRow = PopList_FindDataEndRow(sourceRange, headerRelRow, valueCols)
    If dataEndRow <= headerRelRow Then Exit Function
    If PopList_CountUsableDataRows(sourceRange, headerRelRow + 1, dataEndRow, rowAxisCount, valueCols) = 0 Then Exit Function

    Dim rowAxisHeaders As Collection
    Set rowAxisHeaders = PopList_ReadRowAxisHeaders(sourceRange, headerRelRow, rowAxisCount)

    Dim block As Object
    Set block = CreateObject("Scripting.Dictionary")
    block.Add "title", PopList_FindTitle(sourceRange, headerRelRow)
    block.Add "headerRelRow", headerRelRow
    block.Add "dataStartRelRow", headerRelRow + 1
    block.Add "dataEndRelRow", dataEndRow
    block.Add "rowAxisCount", rowAxisCount
    block.Add "firstValueRelCol", firstValueCol
    block.Add "valueCols", valueCols
    block.Add "rowAxisHeaders", rowAxisHeaders
    Set PopList_TryDetectBlock = block
End Function

Private Function PopList_FirstValueHeaderCol(ByVal sourceRange As Range, ByVal headerRelRow As Long) As Long
    Dim relCol As Long
    For relCol = 2 To sourceRange.Columns.Count
        Dim headerText As String
        headerText = PopList_CellText(sourceRange.Cells(headerRelRow, relCol))
        If Len(headerText) > 0 And Not PopList_IsTotalLabel(headerText) Then
            If PopList_ColumnHasNumericBelow(sourceRange, headerRelRow, relCol) Then
                PopList_FirstValueHeaderCol = relCol
                Exit Function
            End If
        End If
    Next relCol
End Function

Private Function PopList_DetectValueColumns(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal firstValueCol As Long) As Collection
    Dim cols As Collection
    Set cols = New Collection

    Dim relCol As Long
    For relCol = firstValueCol To sourceRange.Columns.Count
        Dim headerText As String
        headerText = PopList_CellText(sourceRange.Cells(headerRelRow, relCol))
        If Len(headerText) > 0 And Not PopList_IsTotalLabel(headerText) Then
            If PopList_ColumnHasNumericBelow(sourceRange, headerRelRow, relCol) Then cols.Add relCol
        End If
    Next relCol

    Set PopList_DetectValueColumns = cols
End Function

Private Function PopList_ColumnHasNumericBelow(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal relCol As Long) As Boolean
    Dim relRow As Long
    Dim scanLimit As Long
    scanLimit = headerRelRow + 80
    If scanLimit > sourceRange.Rows.Count Then scanLimit = sourceRange.Rows.Count

    For relRow = headerRelRow + 1 To scanLimit
        If PopList_IsNumberLike(sourceRange.Cells(relRow, relCol).Value) Then
            PopList_ColumnHasNumericBelow = True
            Exit Function
        End If
    Next relRow
End Function

Private Function PopList_FindDataEndRow(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal valueCols As Collection) As Long
    Dim seenData As Boolean
    Dim relRow As Long

    For relRow = headerRelRow + 1 To sourceRange.Rows.Count
        If PopList_RowHasNumericInValueCols(sourceRange, relRow, valueCols) Then
            seenData = True
            PopList_FindDataEndRow = relRow
        ElseIf seenData Then
            Exit Function
        End If
    Next relRow
End Function

Private Function PopList_RowHasNumericInValueCols(ByVal sourceRange As Range, ByVal relRow As Long, ByVal valueCols As Collection) As Boolean
    Dim i As Long
    For i = 1 To valueCols.Count
        If PopList_IsNumberLike(sourceRange.Cells(relRow, CLng(valueCols(i))).Value) Then
            PopList_RowHasNumericInValueCols = True
            Exit Function
        End If
    Next i
End Function

Private Function PopList_CountUsableDataRows(ByVal sourceRange As Range, ByVal firstRelRow As Long, ByVal lastRelRow As Long, _
                                             ByVal rowAxisCount As Long, ByVal valueCols As Collection) As Long
    Dim relRow As Long
    For relRow = firstRelRow To lastRelRow
        If Not PopList_RowIsTotalByCurrentCells(sourceRange, relRow, rowAxisCount) Then
            If PopList_RowHasNumericInValueCols(sourceRange, relRow, valueCols) Then PopList_CountUsableDataRows = PopList_CountUsableDataRows + 1
        End If
    Next relRow
End Function

Private Function PopList_ReadRowAxisHeaders(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal rowAxisCount As Long) As Collection
    Dim headers As Collection
    Set headers = New Collection

    Dim relCol As Long
    For relCol = 1 To rowAxisCount
        Dim headerText As String
        headerText = PopList_CellText(sourceRange.Cells(headerRelRow, relCol))
        If Len(headerText) = 0 Then headerText = "행축" & relCol
        headers.Add headerText
    Next relCol

    Set PopList_ReadRowAxisHeaders = headers
End Function

Private Function PopList_RowAxisLooksLikeHeader(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal rowAxisCount As Long) As Boolean
    Dim joinedText As String
    Dim relCol As Long
    For relCol = 1 To rowAxisCount
        joinedText = joinedText & PopList_NormalizeLabel(PopList_CellText(sourceRange.Cells(headerRelRow, relCol))) & "|"
    Next relCol

    If InStr(1, joinedText, "구분", vbTextCompare) > 0 _
       Or InStr(1, joinedText, "분류", vbTextCompare) > 0 _
       Or InStr(1, joinedText, "항목", vbTextCompare) > 0 _
       Or InStr(1, joinedText, "업종", vbTextCompare) > 0 _
       Or InStr(1, joinedText, "지역", vbTextCompare) > 0 _
       Or InStr(1, joinedText, "성별", vbTextCompare) > 0 _
       Or InStr(1, joinedText, "연령", vbTextCompare) > 0 _
       Or InStr(1, joinedText, "규모", vbTextCompare) > 0 Then
        PopList_RowAxisLooksLikeHeader = True
    End If
End Function

Private Function PopList_CountNonNumericHeaders(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal valueCols As Collection) As Long
    Dim i As Long
    For i = 1 To valueCols.Count
        Dim headerText As String
        headerText = PopList_CellText(sourceRange.Cells(headerRelRow, CLng(valueCols(i))))
        If Len(headerText) > 0 And Not PopList_IsNumberLike(headerText) Then PopList_CountNonNumericHeaders = PopList_CountNonNumericHeaders + 1
    Next i
End Function

Private Function PopList_FindTitle(ByVal sourceRange As Range, ByVal headerRelRow As Long) As String
    Dim relRow As Long
    For relRow = headerRelRow - 1 To 1 Step -1
        If PopList_RowNonEmptyCount(sourceRange, relRow) > 0 Then
            PopList_FindTitle = PopList_FirstTextInRow(sourceRange, relRow)
            If Len(PopList_FindTitle) > 0 Then Exit Function
        End If
    Next relRow
    PopList_FindTitle = sourceRange.Worksheet.Name & "!" & sourceRange.Rows(headerRelRow).Address(False, False)
End Function

Private Function PopList_RowNonEmptyCount(ByVal sourceRange As Range, ByVal relRow As Long) As Long
    Dim relCol As Long
    For relCol = 1 To sourceRange.Columns.Count
        If Len(PopList_CellText(sourceRange.Cells(relRow, relCol))) > 0 Then PopList_RowNonEmptyCount = PopList_RowNonEmptyCount + 1
    Next relCol
End Function

Private Function PopList_FirstTextInRow(ByVal sourceRange As Range, ByVal relRow As Long) As String
    Dim relCol As Long
    For relCol = 1 To sourceRange.Columns.Count
        PopList_FirstTextInRow = PopList_CellText(sourceRange.Cells(relRow, relCol))
        If Len(PopList_FirstTextInRow) > 0 Then Exit Function
    Next relCol
End Function

Private Function PopList_ShowPreviewForm(ByVal blocks As Collection, ByVal sourceRange As Range, _
                                         ByRef columnAxisName As String, ByRef valueName As String, _
                                         ByRef outputSheetName As String) As Boolean
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmPopulationListifyPreview")
    form.Configure PopList_BuildSummary(blocks, sourceRange), _
                   PopList_BuildPreview(blocks, sourceRange, columnAxisName, valueName, 25), _
                   columnAxisName, valueName, outputSheetName
    form.Show

    If CBool(form.Confirmed) Then
        columnAxisName = Trim$(CStr(form.ColumnAxisName))
        valueName = Trim$(CStr(form.ValueName))
        outputSheetName = Trim$(CStr(form.OutputSheetName))
        If Len(columnAxisName) = 0 Then columnAxisName = "가로구분"
        If Len(valueName) = 0 Then valueName = "값"
        If Len(outputSheetName) = 0 Then outputSheetName = "모집단_리스트"
        PopList_ShowPreviewForm = True
    End If
    Unload form
    Exit Function

Fallback:
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    On Error GoTo 0

    If MsgBox("미리보기 창을 열지 못했습니다. 탐지 결과를 기준으로 바로 실행할까요?" & vbCrLf & vbCrLf & _
              PopList_BuildSummary(blocks, sourceRange), _
              vbQuestion + vbYesNo, "모집단 구성표 리스트화") = vbYes Then
        PopList_ShowPreviewForm = True
    End If
End Function

Private Function PopList_BuildSummary(ByVal blocks As Collection, ByVal sourceRange As Range) As String
    Dim result As String
    result = "선택 범위: " & sourceRange.Worksheet.Name & "!" & sourceRange.Address(False, False) & vbCrLf & _
             "탐지된 집계표: " & blocks.Count & "개" & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To blocks.Count
        Dim block As Object
        Set block = blocks(i)
        result = result & i & ". " & CStr(block("title")) & vbCrLf & _
                 "   헤더 행: " & (sourceRange.Row + CLng(block("headerRelRow")) - 1) & _
                 " / 세로축: " & PopList_JoinCollection(block("rowAxisHeaders"), ", ") & _
                 " / 가로 항목: " & CStr(block("valueCols").Count) & _
                 " / 데이터 행: " & PopList_CountUsableDataRows(sourceRange, CLng(block("dataStartRelRow")), CLng(block("dataEndRelRow")), CLng(block("rowAxisCount")), block("valueCols")) & vbCrLf
    Next i

    PopList_BuildSummary = result
End Function

Private Function PopList_BuildPreview(ByVal blocks As Collection, ByVal sourceRange As Range, _
                                      ByVal columnAxisName As String, ByVal valueName As String, _
                                      ByVal maxRows As Long) As String
    Dim result As String
    result = "표제목" & vbTab & "행축명" & vbTab & "행축값" & vbTab & columnAxisName & vbTab & valueName & vbCrLf

    Dim madeRows As Long
    Dim i As Long
    For i = 1 To blocks.Count
        Dim block As Object
        Set block = blocks(i)

        Dim lastSeen() As String
        ReDim lastSeen(1 To CLng(block("rowAxisCount")))

        Dim relRow As Long
        For relRow = CLng(block("dataStartRelRow")) To CLng(block("dataEndRelRow"))
            If Not PopList_RowIsTotalByCurrentCells(sourceRange, relRow, CLng(block("rowAxisCount"))) Then
                PopList_UpdateLastSeen sourceRange, relRow, CLng(block("rowAxisCount")), lastSeen

                Dim valueCols As Collection
                Set valueCols = block("valueCols")
                Dim c As Long
                For c = 1 To valueCols.Count
                    Dim relCol As Long
                    relCol = CLng(valueCols(c))
                    If PopList_IsNumberLike(sourceRange.Cells(relRow, relCol).Value) Then
                        result = result & CStr(block("title")) & vbTab & _
                                 PopList_CollectionText(block("rowAxisHeaders"), 1) & vbTab & _
                                 lastSeen(1) & vbTab & _
                                 PopList_CellText(sourceRange.Cells(CLng(block("headerRelRow")), relCol)) & vbTab & _
                                 PopList_ToNumber(sourceRange.Cells(relRow, relCol).Value) & vbCrLf
                        madeRows = madeRows + 1
                        If madeRows >= maxRows Then
                            PopList_BuildPreview = result & "... (미리보기 " & maxRows & "행까지만 표시)"
                            Exit Function
                        End If
                    End If
                Next c
            End If
        Next relRow
    Next i

    PopList_BuildPreview = result
End Function

Private Function PopList_WriteOutput(ByVal blocks As Collection, ByVal sourceRange As Range, ByVal outWs As Worksheet, _
                                     ByVal columnAxisName As String, ByVal valueName As String) As Long
    outWs.Cells.Clear
    outWs.Range("A1").Value = "모집단 구성표 리스트"
    outWs.Range("A1").Font.Bold = True
    outWs.Range("A1").Font.Size = 14
    outWs.Range("A2").Value = "원본 시트"
    outWs.Range("B2").Value = sourceRange.Worksheet.Name
    outWs.Range("A3").Value = "원본 범위"
    outWs.Range("B3").Value = sourceRange.Address(False, False)

    Dim maxAxisCount As Long
    maxAxisCount = PopList_MaxRowAxisCount(blocks)

    Dim headerRow As Long
    headerRow = 5
    Dim outCol As Long
    outCol = 1
    outWs.Cells(headerRow, outCol).Value = "표제목"
    outCol = outCol + 1

    Dim axisIndex As Long
    For axisIndex = 1 To maxAxisCount
        outWs.Cells(headerRow, outCol).Value = "행축" & axisIndex & "명"
        outWs.Cells(headerRow, outCol + 1).Value = "행축" & axisIndex & "값"
        outCol = outCol + 2
    Next axisIndex

    Dim columnAxisCol As Long
    Dim valueCol As Long
    Dim sourceRowCol As Long
    Dim sourceColCol As Long
    columnAxisCol = outCol
    valueCol = outCol + 1
    sourceRowCol = outCol + 2
    sourceColCol = outCol + 3
    outWs.Cells(headerRow, columnAxisCol).Value = columnAxisName
    outWs.Cells(headerRow, valueCol).Value = valueName
    outWs.Cells(headerRow, sourceRowCol).Value = "원본행"
    outWs.Cells(headerRow, sourceColCol).Value = "원본열"

    Dim outRow As Long
    outRow = headerRow + 1

    Dim i As Long
    For i = 1 To blocks.Count
        Dim block As Object
        Set block = blocks(i)

        Dim lastSeen() As String
        ReDim lastSeen(1 To CLng(block("rowAxisCount")))

        Dim relRow As Long
        For relRow = CLng(block("dataStartRelRow")) To CLng(block("dataEndRelRow"))
            If Not PopList_RowIsTotalByCurrentCells(sourceRange, relRow, CLng(block("rowAxisCount"))) Then
                PopList_UpdateLastSeen sourceRange, relRow, CLng(block("rowAxisCount")), lastSeen

                Dim valueCols As Collection
                Set valueCols = block("valueCols")
                Dim valueIndex As Long
                For valueIndex = 1 To valueCols.Count
                    Dim relCol As Long
                    relCol = CLng(valueCols(valueIndex))
                    If PopList_IsNumberLike(sourceRange.Cells(relRow, relCol).Value) Then
                        outWs.Cells(outRow, 1).Value = CStr(block("title"))

                        outCol = 2
                        For axisIndex = 1 To maxAxisCount
                            If axisIndex <= CLng(block("rowAxisCount")) Then
                                outWs.Cells(outRow, outCol).Value = PopList_CollectionText(block("rowAxisHeaders"), axisIndex)
                                outWs.Cells(outRow, outCol + 1).Value = lastSeen(axisIndex)
                            End If
                            outCol = outCol + 2
                        Next axisIndex

                        outWs.Cells(outRow, columnAxisCol).Value = PopList_CellText(sourceRange.Cells(CLng(block("headerRelRow")), relCol))
                        outWs.Cells(outRow, valueCol).Value = PopList_ToNumber(sourceRange.Cells(relRow, relCol).Value)
                        outWs.Cells(outRow, sourceRowCol).Value = sourceRange.Row + relRow - 1
                        outWs.Cells(outRow, sourceColCol).Value = sourceRange.Column + relCol - 1

                        outRow = outRow + 1
                    End If
                Next valueIndex
            End If
        Next relRow
    Next i

    Dim lastCol As Long
    lastCol = sourceColCol
    With outWs.Range(outWs.Cells(headerRow, 1), outWs.Cells(Application.Max(headerRow, outRow - 1), lastCol))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    outWs.Range(outWs.Cells(headerRow, 1), outWs.Cells(headerRow, lastCol)).Font.Bold = True
    outWs.Range(outWs.Cells(headerRow, 1), outWs.Cells(headerRow, lastCol)).Interior.Color = RGB(220, 235, 250)
    outWs.Columns(valueCol).NumberFormat = "#,##0"

    PopList_WriteOutput = outRow - headerRow - 1
End Function

Private Sub PopList_UpdateLastSeen(ByVal sourceRange As Range, ByVal relRow As Long, ByVal rowAxisCount As Long, ByRef lastSeen() As String)
    Dim axisIndex As Long
    For axisIndex = 1 To rowAxisCount
        Dim axisValue As String
        axisValue = PopList_CellText(sourceRange.Cells(relRow, axisIndex))
        If Len(axisValue) > 0 Then lastSeen(axisIndex) = axisValue
    Next axisIndex
End Sub

Private Function PopList_RowIsTotalByCurrentCells(ByVal sourceRange As Range, ByVal relRow As Long, ByVal rowAxisCount As Long) As Boolean
    Dim axisIndex As Long
    For axisIndex = 1 To rowAxisCount
        If PopList_IsTotalLabel(PopList_CellText(sourceRange.Cells(relRow, axisIndex))) Then
            PopList_RowIsTotalByCurrentCells = True
            Exit Function
        End If
    Next axisIndex
End Function

Private Function PopList_MaxRowAxisCount(ByVal blocks As Collection) As Long
    Dim i As Long
    For i = 1 To blocks.Count
        If CLng(blocks(i)("rowAxisCount")) > PopList_MaxRowAxisCount Then PopList_MaxRowAxisCount = CLng(blocks(i)("rowAxisCount"))
    Next i
End Function

Private Function PopList_JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To items.Count
        If Len(PopList_JoinCollection) > 0 Then PopList_JoinCollection = PopList_JoinCollection & delimiter
        PopList_JoinCollection = PopList_JoinCollection & CStr(items(i))
    Next i
End Function

Private Function PopList_CollectionText(ByVal items As Collection, ByVal index As Long) As String
    If index >= 1 And index <= items.Count Then PopList_CollectionText = CStr(items(index))
End Function

Private Function PopList_CellText(ByVal targetCell As Range) As String
    If IsError(targetCell.Value) Or IsEmpty(targetCell.Value) Then Exit Function
    PopList_CellText = Trim$(CStr(targetCell.Value))
End Function

Private Function PopList_IsNumberLike(ByVal valueIn As Variant) As Boolean
    If IsError(valueIn) Or IsEmpty(valueIn) Then Exit Function

    Dim textValue As String
    textValue = Trim$(CStr(valueIn))
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, " ", "")
    If Len(textValue) = 0 Then Exit Function
    PopList_IsNumberLike = IsNumeric(textValue)
End Function

Private Function PopList_ToNumber(ByVal valueIn As Variant) As Double
    Dim textValue As String
    textValue = Trim$(CStr(valueIn))
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, " ", "")
    If IsNumeric(textValue) Then PopList_ToNumber = CDbl(textValue)
End Function

Private Function PopList_IsTotalLabel(ByVal labelText As String) As Boolean
    Dim normalizedText As String
    normalizedText = PopList_NormalizeLabel(labelText)
    PopList_IsTotalLabel = (normalizedText = "계" _
                            Or normalizedText = "합계" _
                            Or normalizedText = "총계" _
                            Or normalizedText = "전체" _
                            Or UCase$(normalizedText) = "TOTAL")
End Function

Private Function PopList_NormalizeLabel(ByVal labelText As String) As String
    PopList_NormalizeLabel = Replace(Trim$(CStr(labelText)), " ", "")
End Function

Private Function PopList_CleanSheetName(ByVal sheetName As String, ByVal defaultName As String) As String
    Dim result As String
    result = Trim$(sheetName)
    If Len(result) = 0 Then result = defaultName

    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    result = Replace(result, ":", "_")

    If Len(result) > 31 Then result = Left$(result, 31)
    If Len(result) = 0 Then result = defaultName
    PopList_CleanSheetName = result
End Function

Private Function SampleDesign_PrepareSheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal overwritePrompt As String) As Worksheet
    Dim ws As Worksheet
    Set ws = SampleDesign_FindSheet(wb, sheetName)
    If Not ws Is Nothing Then
        If MsgBox(overwritePrompt, vbQuestion + vbYesNo, "표본설계") = vbYes Then
            ws.Cells.Clear
            Set SampleDesign_PrepareSheet = ws
        Else
            Set SampleDesign_PrepareSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            SampleDesign_PrepareSheet.Name = SPSS_UniqueSheetName(wb, sheetName)
        End If
    Else
        Set SampleDesign_PrepareSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        SampleDesign_PrepareSheet.Name = sheetName
    End If
End Function

Private Function SampleDesign_PrepareOutputSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = SampleDesign_FindSheet(wb, sheetName)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    Set SampleDesign_PrepareOutputSheet = ws
End Function

Private Function SampleDesign_FindInputSheet(ByVal wb As Workbook) As Worksheet
    Set SampleDesign_FindInputSheet = SampleDesign_FindSheet(wb, "표본설계_입력")
End Function

Private Function SampleDesign_FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SampleDesign_FindSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub SampleDesign_WriteInputTemplate(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 입력"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A3").Value = "프로젝트명"
    ws.Range("A4").Value = "목표 표본 수"
    ws.Range("A5").Value = "기본 최소 셀 표본"
    ws.Range("A6").Value = "기본 최대 셀 표본"
    ws.Range("A7").Value = "배분 방식"
    ws.Range("A8").Value = "멱배분 alpha"
    ws.Range("A9").Value = "잔여 배정"
    ws.Range("C3").Value = "선택"
    ws.Range("C4").Value = "필수"
    ws.Range("C5").Value = "선택"
    ws.Range("C6").Value = "선택"
    ws.Range("C7").Value = "필수"
    ws.Range("C8").Value = "멱배분 선택 시 필수"
    ws.Range("C9").Value = "자동"
    ws.Range("B3").Value = "신규 조사"
    ws.Range("B4").Value = 1000
    ws.Range("B5").Value = 0
    ws.Range("B6").Value = 0
    ws.Range("B7").Value = "단순 비례"
    ws.Range("B8").Value = 0.5
    ws.Range("B9").Value = "소수점 잔차 큰 순"
    ws.Range("D3").Value = "배분 방식 입력값"
    ws.Range("D4").Value = "단순 비례"
    ws.Range("D5").Value = "제곱근 비례"
    ws.Range("D6").Value = "멱배분"
    ws.Range("D7").Value = "균등 배분"
    ws.Range("D8").Value = "사용자 지정 가중"
    ws.Range("A11").Value = "모집단 표"
    ws.Range("A11").Font.Bold = True
    ws.Range("A12").Value = "쿼터 사용 여부"
    SampleDesign_WriteArray ws, 12, 2, Array("사용", "사용", "사용", "미사용", "미사용")
    ws.Range("G12").Value = "필수"
    ws.Range("I12").Value = "선택"
    ws.Range("J12").Value = "선택"
    ws.Range("K12").Value = "사용자 지정 가중 선택 시 사용"
    SampleDesign_WriteArray ws, 13, 1, Array("쿼터ID", "지역", "업종", "종사자규모", "성별", "연령", "모집단수", "우선배정여부", "최소표본", "최대표본", "사용자가중치", "비고")
    SampleDesign_WriteArray ws, 14, 1, Array("Q001", "서울", "제조업", "10인 미만", "", "", 1000, "", "", "", "", "")
    SampleDesign_WriteArray ws, 15, 1, Array("Q002", "부산", "서비스업", "10~49인", "", "", 500, "", "", "", "", "")
    With ws.Range("A13:L13")
        .Font.Bold = True
        .Interior.Color = RGB(220, 235, 250)
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A12:K12")
        .Interior.Color = RGB(255, 245, 210)
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("B12:F12").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="사용,미사용"
    End With
    ws.Range("A3:C9").Borders.LineStyle = xlContinuous
    ws.Columns("A:L").AutoFit
End Sub

Private Sub SampleDesign_WriteEmptyResult(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 할당결과"
    ws.Range("A1").Font.Bold = True
    SampleDesign_WriteArray ws, 3, 1, SampleDesign_ResultHeaders()
    ws.Columns("A:O").AutoFit
End Sub

Private Sub SampleDesign_WriteEmptyReview(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 검토"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "할당 계산 후 검토 요약이 표시됩니다."
    ws.Columns("A:D").AutoFit
End Sub

Private Function SampleDesign_SelectMethod(ByRef methodName As String, ByRef alpha As Double) As Boolean
    Dim result As String
    result = Toolkit_ButtonMenu( _
        "표본설계 - 배분 방식", _
        "배분 방식을 선택하세요." & vbCrLf & "현재 입력값: " & IIf(Len(methodName) = 0, "단순 비례", methodName), _
        "1" & vbTab & "단순 비례" & vbLf & _
        "2" & vbTab & "제곱근 비례" & vbLf & _
        "3" & vbTab & "멱배분" & vbLf & _
        "4" & vbTab & "균등 배분" & vbLf & _
        "5" & vbTab & "사용자 지정 가중", _
        SampleDesign_MethodToChoice(methodName))
    If Len(Trim$(result)) = 0 Then Exit Function
    Select Case Trim$(CStr(result))
        Case "1", "단순 비례", "비례": methodName = "단순 비례"
        Case "2", "제곱근 비례", "제곱근": methodName = "제곱근 비례": alpha = 0.5
        Case "3", "멱배분"
            methodName = "멱배분"
            Dim alphaInput As Variant
            alphaInput = Application.InputBox("멱배분 지수 alpha를 입력하세요." & vbCrLf & _
                                              "1=단순 비례, 0.5=제곱근 비례, 0=균등 배분", _
                                              "표본설계 - alpha", CStr(alpha), Type:=1)
            If VarType(alphaInput) = vbBoolean And alphaInput = False Then Exit Function
            alpha = CDbl(alphaInput)
        Case "4", "균등 배분", "균등": methodName = "균등 배분": alpha = 0
        Case "5", "사용자 지정 가중", "사용자 가중": methodName = "사용자 지정 가중"
        Case Else
            MsgBox "배분 방식은 1~5 중 하나를 선택하세요.", vbExclamation, "표본설계"
            Exit Function
    End Select
    SampleDesign_SelectMethod = True
End Function

Private Function SampleDesign_MethodToChoice(ByVal methodName As String) As String
    Select Case Trim$(methodName)
        Case "제곱근 비례": SampleDesign_MethodToChoice = "2"
        Case "멱배분": SampleDesign_MethodToChoice = "3"
        Case "균등 배분": SampleDesign_MethodToChoice = "4"
        Case "사용자 지정 가중": SampleDesign_MethodToChoice = "5"
        Case Else: SampleDesign_MethodToChoice = "1"
    End Select
End Function

Private Sub SampleDesign_ReadPopulation(ByVal ws As Worksheet, ByRef records As Collection, ByRef issues As Collection)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If lastRow < 14 Then Exit Sub
    Dim r As Long
    For r = 14 To lastRow
        Dim popN As Double
        popN = Val(ws.Cells(r, 7).Value)
        If Len(Trim$(CStr(ws.Cells(r, 1).Value & ws.Cells(r, 2).Value & ws.Cells(r, 3).Value & ws.Cells(r, 4).Value & ws.Cells(r, 5).Value & ws.Cells(r, 6).Value & ws.Cells(r, 7).Value))) = 0 Then
        ElseIf popN <= 0 Then
            issues.Add Array("제외", r, "모집단수가 0 이하이거나 숫자가 아닙니다.")
        Else
            Dim key As String
            key = SampleDesign_BuildQuotaKey(ws, r)
            Dim rec As Variant
            If dict.Exists(key) Then
                rec = dict(key)
                rec(6) = CDbl(rec(6)) + popN
                rec(13) = CStr(rec(13)) & ", " & r
                dict(key) = rec
                issues.Add Array("중복합산", r, "같은 쿼터 조합을 기존 행에 합산했습니다.")
            Else
                rec = Array(CStr(ws.Cells(r, 1).Value), SampleDesign_QuotaValue(ws, r, 2), SampleDesign_QuotaValue(ws, r, 3), _
                            SampleDesign_QuotaValue(ws, r, 4), SampleDesign_QuotaValue(ws, r, 5), SampleDesign_QuotaValue(ws, r, 6), _
                            popN, CStr(ws.Cells(r, 8).Value), CLng(Val(ws.Cells(r, 9).Value)), _
                            CLng(Val(ws.Cells(r, 10).Value)), CDbl(Val(ws.Cells(r, 11).Value)), _
                            CStr(ws.Cells(r, 12).Value), key, CStr(r))
                dict.Add key, rec
            End If
        End If
    Next r
    Dim item As Variant
    For Each item In dict.Items
        records.Add item
    Next item
End Sub

Private Function SampleDesign_BuildQuotaKey(ByVal ws As Worksheet, ByVal rowIndex As Long) As String
    SampleDesign_BuildQuotaKey = SampleDesign_QuotaValue(ws, rowIndex, 2) & "|" & _
                                SampleDesign_QuotaValue(ws, rowIndex, 3) & "|" & _
                                SampleDesign_QuotaValue(ws, rowIndex, 4) & "|" & _
                                SampleDesign_QuotaValue(ws, rowIndex, 5) & "|" & _
                                SampleDesign_QuotaValue(ws, rowIndex, 6)
End Function

Private Function SampleDesign_QuotaValue(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    If SampleDesign_QuotaEnabled(ws, colIndex) Then
        SampleDesign_QuotaValue = Trim$(CStr(ws.Cells(rowIndex, colIndex).Value))
    Else
        SampleDesign_QuotaValue = ""
    End If
End Function

Private Function SampleDesign_QuotaEnabled(ByVal ws As Worksheet, ByVal colIndex As Long) As Boolean
    Dim flagText As String
    flagText = Replace(Trim$(CStr(ws.Cells(12, colIndex).Value)), " ", "")
    SampleDesign_QuotaEnabled = (flagText = "사용" Or UCase$(flagText) = "Y" Or UCase$(flagText) = "YES")
End Function

Private Function SampleDesign_Allocate(ByVal records As Collection, ByVal targetN As Long, _
                                       ByVal methodName As String, ByVal alpha As Double, _
                                       ByVal defaultMin As Long, ByVal defaultMax As Long, _
                                       ByRef issues As Collection) As Collection
    Dim result As New Collection
    Dim i As Long, totalPop As Double, totalWeight As Double
    For i = 1 To records.Count
        Dim rec As Variant
        rec = records(i)
        totalPop = totalPop + CDbl(rec(6))
        totalWeight = totalWeight + SampleDesign_RecordWeight(rec, methodName, alpha)
    Next i
    If totalWeight <= 0 Then
        issues.Add Array("오류", 0, "배분 가중치 합계가 0입니다.")
        Set SampleDesign_Allocate = result
        Exit Function
    End If
    Dim finals() As Long, residuals() As Double, raws() As Double
    ReDim finals(1 To records.Count)
    ReDim residuals(1 To records.Count)
    ReDim raws(1 To records.Count)
    Dim assigned As Long
    For i = 1 To records.Count
        rec = records(i)
        Dim weight As Double
        weight = SampleDesign_RecordWeight(rec, methodName, alpha)
        raws(i) = targetN * weight / totalWeight
        finals(i) = Fix(raws(i))
        residuals(i) = raws(i) - finals(i)
        assigned = assigned + finals(i)
    Next i
    Dim remain As Long
    remain = targetN - assigned
    Do While remain > 0
        Dim addIndex As Long
        addIndex = SampleDesign_LargestResidualIndex(residuals, finals, records, defaultMax)
        If addIndex = 0 Then Exit Do
        finals(addIndex) = finals(addIndex) + 1
        residuals(addIndex) = -1
        remain = remain - 1
    Loop
    For i = 1 To records.Count
        rec = records(i)
        Dim minN As Long, maxN As Long
        minN = CLng(rec(8)): maxN = CLng(rec(9))
        If minN <= 0 Then minN = defaultMin
        If maxN <= 0 Then maxN = defaultMax
        If minN > 0 And finals(i) < minN Then
            issues.Add Array("최소보정", rec(13), "최소표본 " & minN & " 적용")
            finals(i) = minN
        End If
        If maxN > 0 And finals(i) > maxN Then
            issues.Add Array("최대보정", rec(13), "최대표본 " & maxN & " 적용")
            finals(i) = maxN
        End If
    Next i
    SampleDesign_AdjustToTarget finals, raws, records, targetN, defaultMax, issues
    For i = 1 To records.Count
        rec = records(i)
        Dim noteText As String
        noteText = CStr(rec(11))
        If finals(i) = 0 Then noteText = Trim$(noteText & " 최종할당0")
        result.Add Array(rec(0), rec(1), rec(2), rec(3), rec(4), rec(5), CDbl(rec(6)), _
                         CDbl(rec(6)) / totalPop, methodName, raws(i), finals(i), _
                         IIf(targetN > 0, finals(i) / targetN, 0), CDbl(rec(10)), _
                         SampleDesign_ConstraintLabel(rec, finals(i), defaultMin, defaultMax), noteText)
    Next i
    Set SampleDesign_Allocate = result
End Function

Private Function SampleDesign_RecordWeight(ByVal rec As Variant, ByVal methodName As String, ByVal alpha As Double) As Double
    Select Case methodName
        Case "제곱근 비례": SampleDesign_RecordWeight = Sqr(CDbl(rec(6)))
        Case "멱배분": SampleDesign_RecordWeight = CDbl(rec(6)) ^ alpha
        Case "균등 배분": SampleDesign_RecordWeight = 1
        Case "사용자 지정 가중"
            If CDbl(rec(10)) > 0 Then
                SampleDesign_RecordWeight = CDbl(rec(10))
            Else
                SampleDesign_RecordWeight = CDbl(rec(6))
            End If
        Case Else: SampleDesign_RecordWeight = CDbl(rec(6))
    End Select
End Function

Private Function SampleDesign_LargestResidualIndex(ByRef residuals() As Double, ByRef finalN() As Long, ByVal records As Collection, ByVal defaultMax As Long) As Long
    Dim bestIdx As Long, bestVal As Double, i As Long
    bestVal = -2
    For i = LBound(residuals) To UBound(residuals)
        Dim rec As Variant
        rec = records(i)
        Dim maxN As Long
        maxN = CLng(rec(9))
        If maxN <= 0 Then maxN = defaultMax
        If (maxN <= 0 Or finalN(i) < maxN) And residuals(i) > bestVal Then
            bestVal = residuals(i)
            bestIdx = i
        End If
    Next i
    SampleDesign_LargestResidualIndex = bestIdx
End Function

Private Sub SampleDesign_AdjustToTarget(ByRef finalN() As Long, ByRef rawN() As Double, ByVal records As Collection, _
                                        ByVal targetN As Long, ByVal defaultMax As Long, ByRef issues As Collection)
    Dim currentTotal As Long
    currentTotal = SampleDesign_ArrayTotal(finalN)
    Dim guard As Long
    Do While currentTotal < targetN And guard < 100000
        Dim addIdx As Long
        addIdx = SampleDesign_AdjustAddIndex(finalN, rawN, records, defaultMax)
        If addIdx = 0 Then Exit Do
        finalN(addIdx) = finalN(addIdx) + 1
        currentTotal = currentTotal + 1
        guard = guard + 1
    Loop
    Do While currentTotal > targetN And guard < 100000
        Dim subIdx As Long
        subIdx = SampleDesign_AdjustSubtractIndex(finalN, rawN, records)
        If subIdx = 0 Then Exit Do
        finalN(subIdx) = finalN(subIdx) - 1
        currentTotal = currentTotal - 1
        guard = guard + 1
    Loop
    If currentTotal <> targetN Then
        issues.Add Array("합계불일치", 0, "최소/최대 제약 때문에 최종합계가 목표 표본 수와 다릅니다. 최종합계=" & currentTotal)
    End If
End Sub

Private Function SampleDesign_AdjustAddIndex(ByRef finalN() As Long, ByRef rawN() As Double, ByVal records As Collection, ByVal defaultMax As Long) As Long
    Dim bestIdx As Long, bestGap As Double, i As Long
    bestGap = -999999
    For i = LBound(finalN) To UBound(finalN)
        Dim rec As Variant
        rec = records(i)
        Dim maxN As Long
        maxN = CLng(rec(9))
        If maxN <= 0 Then maxN = defaultMax
        If maxN <= 0 Or finalN(i) < maxN Then
            If rawN(i) - finalN(i) > bestGap Then
                bestGap = rawN(i) - finalN(i)
                bestIdx = i
            End If
        End If
    Next i
    SampleDesign_AdjustAddIndex = bestIdx
End Function

Private Function SampleDesign_AdjustSubtractIndex(ByRef finalN() As Long, ByRef rawN() As Double, ByVal records As Collection) As Long
    Dim bestIdx As Long, bestGap As Double, i As Long
    bestGap = -999999
    For i = LBound(finalN) To UBound(finalN)
        Dim rec As Variant
        rec = records(i)
        Dim minN As Long
        minN = CLng(rec(8))
        If finalN(i) > minN And finalN(i) > 0 Then
            If finalN(i) - rawN(i) > bestGap Then
                bestGap = finalN(i) - rawN(i)
                bestIdx = i
            End If
        End If
    Next i
    SampleDesign_AdjustSubtractIndex = bestIdx
End Function

Private Function SampleDesign_ArrayTotal(ByRef values() As Long) As Long
    Dim i As Long
    For i = LBound(values) To UBound(values)
        SampleDesign_ArrayTotal = SampleDesign_ArrayTotal + values(i)
    Next i
End Function

Private Function SampleDesign_ConstraintLabel(ByVal rec As Variant, ByVal finalN As Long, ByVal defaultMin As Long, ByVal defaultMax As Long) As String
    Dim minN As Long, maxN As Long
    minN = CLng(rec(8)): maxN = CLng(rec(9))
    If minN <= 0 Then minN = defaultMin
    If maxN <= 0 Then maxN = defaultMax
    If minN > 0 And finalN = minN Then
        SampleDesign_ConstraintLabel = "최소 적용"
    ElseIf maxN > 0 And finalN = maxN Then
        SampleDesign_ConstraintLabel = "최대 적용"
    Else
        SampleDesign_ConstraintLabel = ""
    End If
End Function

Private Sub SampleDesign_WriteAllocationResult(ByVal ws As Worksheet, ByVal allocation As Collection, ByVal targetN As Long, ByVal methodName As String)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 할당결과"
    ws.Range("B1").Value = "배분 방식": ws.Range("C1").Value = methodName
    ws.Range("D1").Value = "목표 표본 수": ws.Range("E1").Value = targetN
    SampleDesign_WriteArray ws, 3, 1, SampleDesign_ResultHeaders()
    Dim r As Long
    For r = 1 To allocation.Count
        SampleDesign_WriteArray ws, r + 3, 1, allocation(r)
    Next r
    Dim totalRow As Long
    totalRow = allocation.Count + 4
    ws.Cells(totalRow, 1).Value = "총계"
    ws.Cells(totalRow, 7).Formula = "=SUM(G4:G" & (totalRow - 1) & ")"
    ws.Cells(totalRow, 11).Formula = "=SUM(K4:K" & (totalRow - 1) & ")"
    ws.Cells(totalRow, 12).Formula = "=SUM(L4:L" & (totalRow - 1) & ")"
    ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, 15)).Font.Bold = True
    With ws.Range(ws.Cells(3, 1), ws.Cells(totalRow, 15))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    ws.Range("A3:O3").Font.Bold = True
    ws.Range("A3:O3").Interior.Color = RGB(220, 235, 250)
    ws.Range("H:H,L:L").NumberFormat = "0.0%"
    ws.Range("J:J").NumberFormat = "0.000"
    SampleDesign_ApplyResultFormatting ws, 4, totalRow - 1
End Sub

Private Function SampleDesign_ResultHeaders() As Variant
    SampleDesign_ResultHeaders = Array("쿼터ID", "지역", "업종", "종사자규모", "성별", "연령", "모집단수", _
                                       "모집단비율", "배분방식", "원시할당", "최종할당", "표본비율", _
                                       "사용자가중치", "제약 적용", "비고")
End Function

Private Sub SampleDesign_WriteReviewReport(ByVal ws As Worksheet, ByVal allocation As Collection, ByVal issues As Collection, _
                                           ByVal targetN As Long, ByVal methodName As String, ByVal alpha As Double)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 검토"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    Dim totalPop As Double, totalFinal As Long, zeroCount As Long, i As Long
    For i = 1 To allocation.Count
        Dim rec As Variant
        rec = allocation(i)
        totalPop = totalPop + CDbl(rec(6))
        totalFinal = totalFinal + CLng(rec(10))
        If CLng(rec(10)) = 0 Then zeroCount = zeroCount + 1
    Next i
    Dim rowNo As Long
    rowNo = 3
    SampleDesign_WriteKV ws, rowNo, "배분 방식", methodName
    SampleDesign_WriteKV ws, rowNo, "멱배분 alpha", alpha
    SampleDesign_WriteKV ws, rowNo, "총 모집단 수", totalPop
    SampleDesign_WriteKV ws, rowNo, "목표 표본 수", targetN
    SampleDesign_WriteKV ws, rowNo, "최종 배정 표본 수", totalFinal
    SampleDesign_WriteKV ws, rowNo, "쿼터 셀 수", allocation.Count
    SampleDesign_WriteKV ws, rowNo, "최종할당 0인 셀 수", zeroCount
    SampleDesign_WriteKV ws, rowNo, "검토/경고 건수", issues.Count
    rowNo = rowNo + 2
    ws.Cells(rowNo, 1).Value = "검토 항목": ws.Cells(rowNo, 1).Font.Bold = True
    rowNo = rowNo + 1
    SampleDesign_WriteArray ws, rowNo, 1, Array("유형", "원본행", "내용")
    ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 3)).Font.Bold = True
    rowNo = rowNo + 1
    For i = 1 To issues.Count
        SampleDesign_WriteArray ws, rowNo, 1, issues(i)
        rowNo = rowNo + 1
    Next i
    If issues.Count = 0 Then ws.Cells(rowNo, 1).Value = "검토 경고 없음"
    ws.Columns("A:D").AutoFit
End Sub

Private Sub SampleDesign_WriteKV(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal keyText As String, ByVal valueData As Variant)
    ws.Cells(rowNo, 1).Value = keyText
    ws.Cells(rowNo, 2).Value = valueData
    rowNo = rowNo + 1
End Sub

Private Sub SampleDesign_WriteArray(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal values As Variant)
    ws.Cells(rowIndex, colIndex).Resize(1, UBound(values) - LBound(values) + 1).Value = values
End Sub

Private Function SampleDesign_TotalFinalN(ByVal allocation As Collection) As Long
    Dim i As Long
    For i = 1 To allocation.Count
        Dim rec As Variant
        rec = allocation(i)
        SampleDesign_TotalFinalN = SampleDesign_TotalFinalN + CLng(rec(10))
    Next i
End Function

Private Sub SampleDesign_ApplyResultFormatting(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If lastRow < firstRow Then Exit Sub
    On Error Resume Next
    With ws.Range(ws.Cells(firstRow, 11), ws.Cells(lastRow, 11)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
        .Interior.Color = RGB(255, 230, 230)
        .Font.Color = RGB(160, 0, 0)
    End With
    With ws.Range(ws.Cells(firstRow, 14), ws.Cells(lastRow, 14)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=" & Chr$(34) & Chr$(34))
        .Interior.Color = RGB(255, 245, 190)
    End With
    On Error GoTo 0
End Sub

Private Sub SampleDesign_ApplyProgressFormatting(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If lastRow < firstRow Then Exit Sub
    On Error Resume Next
    With ws.Range(ws.Cells(firstRow, 12), ws.Cells(lastRow, 12)).FormatConditions.Add(Type:=xlTextString, String:="미달", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 245, 190)
    End With
    With ws.Range(ws.Cells(firstRow, 12), ws.Cells(lastRow, 12)).FormatConditions.Add(Type:=xlTextString, String:="충족", TextOperator:=xlContains)
        .Interior.Color = RGB(220, 245, 220)
    End With
    With ws.Range(ws.Cells(firstRow, 12), ws.Cells(lastRow, 12)).FormatConditions.Add(Type:=xlTextString, String:="초과", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 220, 220)
    End With
    On Error GoTo 0
End Sub

Public Sub Database_MergeSheetsByHeaders()
    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb.Worksheets.Count < 1 Then Exit Sub

    Dim selectedSheetNames As String
    Dim keyHeader As String
    Dim outputHeadersCsv As String
    Dim outputSheetName As String
    Dim outputPosition As String

    keyHeader = "ID"
    outputHeadersCsv = DBMerge_DefaultHeaderCsv(wb)
    outputSheetName = "통합데이터"
    outputPosition = "맨 뒤"

    If Not DBMerge_ShowSetupForm(wb, selectedSheetNames, keyHeader, outputHeadersCsv, outputSheetName, outputPosition) Then Exit Sub

    keyHeader = Trim$(keyHeader)
    outputSheetName = DBMerge_CleanSheetName(outputSheetName, "통합데이터")
    If Len(keyHeader) = 0 Then
        MsgBox "메인 KEY 헤더를 입력해야 합니다.", vbExclamation, "시트 데이터 통합"
        Exit Sub
    End If

    If DBMerge_NameIsSelected(selectedSheetNames, outputSheetName) Then
        MsgBox "결과 시트 이름이 통합 대상 시트와 같습니다." & vbCrLf & _
               "원본 시트를 덮어쓰지 않도록 다른 결과 시트명을 입력하세요.", vbExclamation, "시트 데이터 통합"
        Exit Sub
    End If

    Dim sheets As Collection
    Set sheets = DBMerge_SelectedWorksheets(wb, selectedSheetNames)
    If sheets.Count = 0 Then
        MsgBox "통합할 시트를 하나 이상 선택하세요.", vbExclamation, "시트 데이터 통합"
        Exit Sub
    End If

    Dim outputHeaders As Collection
    Set outputHeaders = DBMerge_ParseHeaderList(outputHeadersCsv, keyHeader)
    If outputHeaders.Count = 0 Then
        MsgBox "통합할 헤더를 하나 이상 입력하세요.", vbExclamation, "시트 데이터 통합"
        Exit Sub
    End If

    Dim plan As Object
    Set plan = DBMerge_BuildPlan(sheets, keyHeader, outputHeaders)
    If plan Is Nothing Then Exit Sub

    If Not DBMerge_ShowReviewForm(plan) Then Exit Sub

    Dim outWs As Worksheet
    Set outWs = DBMerge_PrepareOutputSheet(wb, outputSheetName, outputPosition)
    If outWs Is Nothing Then Exit Sub

    Dim writtenRows As Long
    writtenRows = DBMerge_WriteMergedSheet(plan, outWs)
    outWs.Activate
    MsgBox sheets.Count & "개 시트를 통합했습니다." & vbCrLf & _
           "생성 행 수: " & writtenRows & vbCrLf & _
           "결과 시트: " & outWs.Name, vbInformation, "시트 데이터 통합"
End Sub

Public Sub Database_SplitMergedSheetBySource()
    If ActiveWorkbook Is Nothing Or TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastCol As Long
    lastCol = DBMerge_LastUsedCol(ws)
    Dim lastRow As Long
    lastRow = DBMerge_LastUsedRow(ws)
    If lastRow < 3 Or lastCol < 1 Then
        MsgBox "분리할 통합 데이터가 없습니다.", vbExclamation, "통합 시트 분리"
        Exit Sub
    End If

    Dim sourceCol As Long
    sourceCol = DBMerge_FindManagementColumn(ws, "__DBMERGE_SOURCE_SHEET", lastCol)
    If sourceCol = 0 Then
        MsgBox "통합용 관리코드 행을 찾지 못했습니다." & vbCrLf & _
               "시트 데이터 통합 기능으로 생성한 시트에서 실행하세요.", vbExclamation, "통합 시트 분리"
        Exit Sub
    End If

    If MsgBox("현재 시트를 원본시트 값 기준으로 다시 분리합니다." & vbCrLf & _
              "관리코드 1행과 헤더 2행도 각 분리 시트에 복사됩니다." & vbCrLf & vbCrLf & _
              "계속할까요?", vbQuestion + vbYesNo, "통합 시트 분리") = vbNo Then Exit Sub

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 3 To lastRow
        Dim sourceName As String
        sourceName = Trim$(CStr(ws.Cells(r, sourceCol).Value))
        If Len(sourceName) = 0 Then sourceName = "미지정"
        If Not groups.Exists(sourceName) Then
            Dim rowList As Collection
            Set rowList = New Collection
            groups.Add sourceName, rowList
        End If
        groups(sourceName).Add r
    Next r

    If groups.Count = 0 Then
        MsgBox "분리할 데이터 행이 없습니다.", vbInformation, "통합 시트 분리"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim key As Variant
    For Each key In groups.Keys
        Dim outWs As Worksheet
        Set outWs = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
        outWs.Name = DBMerge_UniqueSheetName(ws.Parent, "분리_" & CStr(key))

        ws.Range(ws.Cells(1, 1), ws.Cells(2, lastCol)).Copy Destination:=outWs.Cells(1, 1)

        Dim outRow As Long
        outRow = 3
        Dim rowsForSheet As Collection
        Set rowsForSheet = groups(key)
        Dim i As Long
        For i = 1 To rowsForSheet.Count
            ws.Range(ws.Cells(CLng(rowsForSheet(i)), 1), ws.Cells(CLng(rowsForSheet(i)), lastCol)).Copy Destination:=outWs.Cells(outRow, 1)
            outRow = outRow + 1
        Next i
        outWs.Columns.AutoFit
    Next key
    Application.ScreenUpdating = True

    MsgBox groups.Count & "개 시트로 분리했습니다.", vbInformation, "통합 시트 분리"
End Sub

Private Function DBMerge_ShowSetupForm(ByVal wb As Workbook, ByRef selectedSheetNames As String, _
                                       ByRef keyHeader As String, ByRef outputHeadersCsv As String, _
                                       ByRef outputSheetName As String, ByRef outputPosition As String) As Boolean
    On Error GoTo ErrHandler

    Dim form As Object
    Set form = VBA.UserForms.Add("frmDatabaseMergeSetup")
    form.Configure DBMerge_WorksheetNamesText(wb), keyHeader, outputHeadersCsv, outputSheetName, outputPosition
    form.Show

    If CBool(form.Cancelled) Then
        Unload form
        Exit Function
    End If

    selectedSheetNames = CStr(form.SelectedSheets)
    keyHeader = CStr(form.KeyHeader)
    outputHeadersCsv = CStr(form.OutputHeaders)
    outputSheetName = CStr(form.OutputSheetName)
    outputPosition = CStr(form.OutputPosition)
    Unload form
    DBMerge_ShowSetupForm = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    On Error GoTo 0
    MsgBox "시트 통합 설정 창을 열지 못했습니다." & vbCrLf & Err.Description, vbExclamation, "시트 데이터 통합"
End Function

Private Function DBMerge_ShowReviewForm(ByVal plan As Object) As Boolean
    On Error GoTo ErrHandler

    Dim form As Object
    Set form = VBA.UserForms.Add("frmDatabaseMergeReview")
    form.Configure DBMerge_PlanSummary(plan), DBMerge_MappingSummary(plan)
    form.Show

    If Not CBool(form.Cancelled) Then DBMerge_ShowReviewForm = True
    Unload form
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    On Error GoTo 0
    MsgBox "시트 통합 검증 창을 열지 못했습니다." & vbCrLf & Err.Description, vbExclamation, "시트 데이터 통합"
End Function

Private Function DBMerge_WorksheetNamesText(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Len(DBMerge_WorksheetNamesText) > 0 Then DBMerge_WorksheetNamesText = DBMerge_WorksheetNamesText & vbLf
        DBMerge_WorksheetNamesText = DBMerge_WorksheetNamesText & ws.Name
    Next ws
End Function

Private Function DBMerge_DefaultHeaderCsv(ByVal wb As Workbook) As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim headerRow As Long
        headerRow = DBMerge_DetectHeaderRow(ws)
        If headerRow > 0 Then
            Dim lastCol As Long
            lastCol = DBMerge_LastUsedCol(ws)
            Dim c As Long
            For c = 1 To lastCol
                Dim headerText As String
                headerText = Trim$(CStr(ws.Cells(headerRow, c).Value))
                If Len(headerText) > 0 Then
                    Dim norm As String
                    norm = DBMerge_NormalizeHeader(headerText)
                    If Len(norm) > 0 And Not seen.Exists(norm) Then seen.Add norm, headerText
                End If
            Next c
        End If
    Next ws

    Dim key As Variant
    For Each key In seen.Keys
        If Len(DBMerge_DefaultHeaderCsv) > 0 Then DBMerge_DefaultHeaderCsv = DBMerge_DefaultHeaderCsv & ", "
        DBMerge_DefaultHeaderCsv = DBMerge_DefaultHeaderCsv & CStr(seen(key))
    Next key
End Function

Private Function DBMerge_SelectedWorksheets(ByVal wb As Workbook, ByVal selectedSheetNames As String) As Collection
    Dim sheets As Collection
    Set sheets = New Collection

    Dim lines As Variant
    lines = Split(Replace(selectedSheetNames, vbCr, vbLf), vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim sheetName As String
        sheetName = Trim$(CStr(lines(i)))
        If Len(sheetName) > 0 Then
            On Error Resume Next
            Dim ws As Worksheet
            Set ws = wb.Worksheets(sheetName)
            On Error GoTo 0
            If Not ws Is Nothing Then sheets.Add ws
            Set ws = Nothing
        End If
    Next i

    Set DBMerge_SelectedWorksheets = sheets
End Function

Private Function DBMerge_ParseHeaderList(ByVal headerText As String, ByVal keyHeader As String) As Collection
    Dim headers As Collection
    Set headers = New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    DBMerge_AddHeader headers, seen, keyHeader

    Dim normalizedText As String
    normalizedText = Replace(headerText, vbCrLf, ",")
    normalizedText = Replace(normalizedText, vbCr, ",")
    normalizedText = Replace(normalizedText, vbLf, ",")
    normalizedText = Replace(normalizedText, ";", ",")

    Dim parts As Variant
    parts = Split(normalizedText, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        DBMerge_AddHeader headers, seen, CStr(parts(i))
    Next i

    Set DBMerge_ParseHeaderList = headers
End Function

Private Sub DBMerge_AddHeader(ByVal headers As Collection, ByVal seen As Object, ByVal headerText As String)
    Dim cleaned As String
    cleaned = Trim$(headerText)
    If Len(cleaned) = 0 Then Exit Sub

    Dim norm As String
    norm = DBMerge_NormalizeHeader(cleaned)
    If Len(norm) = 0 Or seen.Exists(norm) Then Exit Sub
    seen.Add norm, True
    headers.Add cleaned
End Sub

Private Function DBMerge_BuildPlan(ByVal sheets As Collection, ByVal keyHeader As String, ByVal outputHeaders As Collection) As Object
    Dim plan As Object
    Set plan = CreateObject("Scripting.Dictionary")
    plan.Add "keyHeader", keyHeader
    plan.Add "headers", outputHeaders

    Dim sheetPlans As Collection
    Set sheetPlans = New Collection

    Dim i As Long
    For i = 1 To sheets.Count
        Dim ws As Worksheet
        Set ws = sheets(i)

        Dim headerRow As Long
        headerRow = DBMerge_DetectHeaderRow(ws)
        If headerRow = 0 Then
            MsgBox ws.Name & " 시트에서 헤더 행을 찾지 못했습니다.", vbExclamation, "시트 데이터 통합"
            Exit Function
        End If

        Dim headerMap As Object
        Set headerMap = DBMerge_HeaderMap(ws, headerRow)

        Dim sheetPlan As Object
        Set sheetPlan = CreateObject("Scripting.Dictionary")
        sheetPlan.Add "sheetName", ws.Name
        sheetPlan.Add "worksheet", ws
        sheetPlan.Add "headerRow", headerRow
        sheetPlan.Add "lastRow", DBMerge_LastDataRow(ws, headerRow, headerMap, outputHeaders)
        sheetPlan.Add "headerMap", headerMap
        sheetPlans.Add sheetPlan
    Next i

    plan.Add "sheets", sheetPlans
    Set DBMerge_BuildPlan = plan
End Function

Private Function DBMerge_DetectHeaderRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = DBMerge_LastUsedRow(ws)
    If lastRow = 0 Then Exit Function

    Dim scanLimit As Long
    scanLimit = lastRow
    If scanLimit > 30 Then scanLimit = 30

    Dim r As Long
    For r = 1 To scanLimit
        If DBMerge_RowTextCount(ws, r, DBMerge_LastUsedCol(ws)) >= 2 Then
            DBMerge_DetectHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function DBMerge_RowTextCount(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = 1 To lastCol
        If Len(Trim$(CStr(ws.Cells(rowNo, c).Value))) > 0 Then DBMerge_RowTextCount = DBMerge_RowTextCount + 1
    Next c
End Function

Private Function DBMerge_HeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = DBMerge_LastUsedCol(ws)
    Dim c As Long
    For c = 1 To lastCol
        Dim headerText As String
        headerText = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If Len(headerText) > 0 Then
            Dim norm As String
            norm = DBMerge_NormalizeHeader(headerText)
            If Len(norm) > 0 And Not map.Exists(norm) Then map.Add norm, c
        End If
    Next c

    Set DBMerge_HeaderMap = map
End Function

Private Function DBMerge_LastDataRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerMap As Object, ByVal outputHeaders As Collection) As Long
    Dim lastRow As Long
    lastRow = headerRow

    Dim i As Long
    For i = 1 To outputHeaders.Count
        Dim norm As String
        norm = DBMerge_NormalizeHeader(CStr(outputHeaders(i)))
        If headerMap.Exists(norm) Then
            Dim colNo As Long
            colNo = CLng(headerMap(norm))
            Dim rowNo As Long
            rowNo = ws.Cells(ws.Rows.Count, colNo).End(xlUp).Row
            If rowNo > lastRow Then lastRow = rowNo
        End If
    Next i

    If lastRow = headerRow Then lastRow = DBMerge_LastUsedRow(ws)
    DBMerge_LastDataRow = lastRow
End Function

Private Function DBMerge_PlanSummary(ByVal plan As Object) As String
    Dim result As String
    result = "메인 KEY: " & CStr(plan("keyHeader")) & vbCrLf & _
             "통합 헤더 수: " & CStr(plan("headers").Count) & vbCrLf & _
             "대상 시트 수: " & CStr(plan("sheets").Count) & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To plan("sheets").Count
        Dim sheetPlan As Object
        Set sheetPlan = plan("sheets")(i)
        result = result & i & ". " & CStr(sheetPlan("sheetName")) & _
                 " / 헤더 행: " & CStr(sheetPlan("headerRow")) & _
                 " / 마지막 행: " & CStr(sheetPlan("lastRow")) & vbCrLf
    Next i
    DBMerge_PlanSummary = result
End Function

Private Function DBMerge_MappingSummary(ByVal plan As Object) As String
    Dim result As String
    result = "통합 헤더" & vbTab & "시트별 매칭 결과" & vbCrLf

    Dim h As Long
    For h = 1 To plan("headers").Count
        Dim headerText As String
        headerText = CStr(plan("headers")(h))
        result = result & headerText & vbTab

        Dim parts As String
        Dim i As Long
        For i = 1 To plan("sheets").Count
            Dim sheetPlan As Object
            Set sheetPlan = plan("sheets")(i)
            If Len(parts) > 0 Then parts = parts & " / "

            Dim norm As String
            norm = DBMerge_NormalizeHeader(headerText)
            If sheetPlan("headerMap").Exists(norm) Then
                parts = parts & CStr(sheetPlan("sheetName")) & ":" & DBMerge_ColumnLetter(CLng(sheetPlan("headerMap")(norm)))
            Else
                parts = parts & CStr(sheetPlan("sheetName")) & ":없음"
            End If
        Next i
        result = result & parts & vbCrLf
    Next h
    DBMerge_MappingSummary = result
End Function

Private Function DBMerge_WriteMergedSheet(ByVal plan As Object, ByVal outWs As Worksheet) As Long
    outWs.Cells.Clear

    Dim dataHeaders As Collection
    Set dataHeaders = DBMerge_DataHeaders(plan("headers"), CStr(plan("keyHeader")))

    outWs.Cells(1, 1).Value = "__DBMERGE_SOURCE_SHEET"
    outWs.Cells(1, 2).Value = "__DBMERGE_SOURCE_ROW"
    outWs.Cells(1, 3).Value = "__DBMERGE_KEY"
    outWs.Cells(2, 1).Value = "원본시트"
    outWs.Cells(2, 2).Value = "원본행"
    outWs.Cells(2, 3).Value = CStr(plan("keyHeader"))

    Dim c As Long
    For c = 1 To dataHeaders.Count
        outWs.Cells(1, c + 3).Value = "__FIELD:" & CStr(dataHeaders(c))
        outWs.Cells(2, c + 3).Value = CStr(dataHeaders(c))
    Next c

    Dim outRow As Long
    outRow = 3

    Dim i As Long
    For i = 1 To plan("sheets").Count
        Dim sheetPlan As Object
        Set sheetPlan = plan("sheets")(i)
        Dim ws As Worksheet
        Set ws = sheetPlan("worksheet")
        Dim headerMap As Object
        Set headerMap = sheetPlan("headerMap")

        Dim r As Long
        For r = CLng(sheetPlan("headerRow")) + 1 To CLng(sheetPlan("lastRow"))
            If DBMerge_RowHasSelectedData(ws, r, headerMap, plan("headers")) Then
                outWs.Cells(outRow, 1).Value = ws.Name
                outWs.Cells(outRow, 2).Value = r
                outWs.Cells(outRow, 3).Value = DBMerge_CellByHeader(ws, r, headerMap, CStr(plan("keyHeader")))

                For c = 1 To dataHeaders.Count
                    outWs.Cells(outRow, c + 3).Value = DBMerge_CellByHeader(ws, r, headerMap, CStr(dataHeaders(c)))
                Next c

                outRow = outRow + 1
            End If
        Next r
    Next i

    Dim lastCol As Long
    lastCol = dataHeaders.Count + 3
    outWs.Rows(1).Font.Color = RGB(120, 120, 120)
    outWs.Rows(1).Interior.Color = RGB(235, 235, 235)
    outWs.Rows(2).Font.Bold = True
    outWs.Rows(2).Interior.Color = RGB(220, 235, 250)
    With outWs.Range(outWs.Cells(1, 1), outWs.Cells(Application.Max(2, outRow - 1), lastCol))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With

    DBMerge_WriteMergedSheet = outRow - 3
End Function

Private Function DBMerge_DataHeaders(ByVal headers As Collection, ByVal keyHeader As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim keyNorm As String
    keyNorm = DBMerge_NormalizeHeader(keyHeader)

    Dim i As Long
    For i = 1 To headers.Count
        If DBMerge_NormalizeHeader(CStr(headers(i))) <> keyNorm Then result.Add CStr(headers(i))
    Next i
    Set DBMerge_DataHeaders = result
End Function

Private Function DBMerge_RowHasSelectedData(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerMap As Object, ByVal headers As Collection) As Boolean
    Dim i As Long
    For i = 1 To headers.Count
        Dim norm As String
        norm = DBMerge_NormalizeHeader(CStr(headers(i)))
        If headerMap.Exists(norm) Then
            If Len(Trim$(CStr(ws.Cells(rowNo, CLng(headerMap(norm))).Value))) > 0 Then
                DBMerge_RowHasSelectedData = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function DBMerge_CellByHeader(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerMap As Object, ByVal headerText As String) As Variant
    Dim norm As String
    norm = DBMerge_NormalizeHeader(headerText)
    If headerMap.Exists(norm) Then DBMerge_CellByHeader = ws.Cells(rowNo, CLng(headerMap(norm))).Value
End Function

Private Function DBMerge_PrepareOutputSheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal outputPosition As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        If MsgBox(sheetName & " 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?", vbQuestion + vbYesNo, "시트 데이터 통합") = vbNo Then Exit Function
        ws.Cells.Clear
        Set DBMerge_PrepareOutputSheet = ws
        Exit Function
    End If

    Select Case outputPosition
        Case "맨 앞"
            Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        Case "현재 시트 뒤"
            Set ws = wb.Worksheets.Add(After:=ActiveSheet)
        Case Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End Select
    ws.Name = sheetName
    Set DBMerge_PrepareOutputSheet = ws
End Function

Private Function DBMerge_NameIsSelected(ByVal selectedSheetNames As String, ByVal sheetName As String) As Boolean
    Dim lines As Variant
    lines = Split(Replace(selectedSheetNames, vbCr, vbLf), vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If StrComp(Trim$(CStr(lines(i))), sheetName, vbTextCompare) = 0 Then
            DBMerge_NameIsSelected = True
            Exit Function
        End If
    Next i
End Function

Private Function DBMerge_FindManagementColumn(ByVal ws As Worksheet, ByVal codeText As String, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), codeText, vbTextCompare) = 0 Then
            DBMerge_FindManagementColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function DBMerge_UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim cleanBase As String
    cleanBase = DBMerge_CleanSheetName(baseName, "분리시트")
    If Len(cleanBase) > 25 Then cleanBase = Left$(cleanBase, 25)

    Dim candidate As String
    candidate = cleanBase
    Dim idx As Long
    idx = 1
    Do While DBMerge_SheetExists(wb, candidate)
        idx = idx + 1
        candidate = Left$(cleanBase, 25) & "_" & CStr(idx)
    Loop
    DBMerge_UniqueSheetName = candidate
End Function

Private Function DBMerge_SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    DBMerge_SheetExists = Not ws Is Nothing
End Function

Private Function DBMerge_CleanSheetName(ByVal sheetName As String, ByVal defaultName As String) As String
    Dim result As String
    result = Trim$(sheetName)
    If Len(result) = 0 Then result = defaultName

    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    result = Replace(result, ":", "_")

    If Len(result) > 31 Then result = Left$(result, 31)
    If Len(result) = 0 Then result = defaultName
    DBMerge_CleanSheetName = result
End Function

Private Function DBMerge_NormalizeHeader(ByVal headerText As String) As String
    Dim result As String
    result = LCase$(Trim$(CStr(headerText)))
    result = Replace(result, " ", "")
    result = Replace(result, vbTab, "")
    result = Replace(result, ChrW(160), "")
    result = Replace(result, "_", "")
    result = Replace(result, "-", "")
    result = Replace(result, ".", "")
    DBMerge_NormalizeHeader = result
End Function

Private Function DBMerge_ColumnLetter(ByVal colNo As Long) As String
    Dim n As Long
    Dim result As String
    n = colNo

    Do While n > 0
        result = Chr$(((n - 1) Mod 26) + 65) & result
        n = (n - 1) \ 26
    Loop
    DBMerge_ColumnLetter = result
End Function

Private Function DBMerge_LastUsedRow(ByVal ws As Worksheet) As Long
    On Error GoTo SafeExit
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not foundCell Is Nothing Then DBMerge_LastUsedRow = foundCell.Row
SafeExit:
End Function

Private Function DBMerge_LastUsedCol(ByVal ws As Worksheet) As Long
    On Error GoTo SafeExit
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not foundCell Is Nothing Then DBMerge_LastUsedCol = foundCell.Column
SafeExit:
End Function

' ============================================================
' 통합 실행 런처
' 설  명  : 리본 버튼 수가 많아지는 문제를 줄이기 위해 유사 기능을
'           하나의 팝업 선택 창에서 실행한다.
' ============================================================
Public Sub Toolkit_LaunchDataCleanup()
    Dim choice As String
    choice = Toolkit_LauncherPrompt( _
        "값 정리", _
        "실행할 값 정리 기능을 선택하세요." & vbCrLf & vbCrLf & _
        "1. 텍스트 숫자 변환 - 범위 선택" & vbCrLf & _
        "2. 텍스트 숫자 변환 - 현재 시트 전체 데이터 영역" & vbCrLf & _
        "3. 전화번호 하이픈 정리" & vbCrLf & _
        "4. 전화번호 숫자만 남기기" & vbCrLf & _
        "5. 사업자등록번호 숫자만 남기기" & vbCrLf & _
        "6. 사업자등록번호 3열 병합" & vbCrLf & _
        "7. 기업체명/사업체명 정리", _
        "1")

    Select Case choice
        Case "1": ConvertTextNumbersWithForm
        Case "2": ConvertTextNumbersInCurrentSheetDataRange
        Case "3": NormalizePhonesWithForm
        Case "4": RemovePhoneDashesToCopiedColumns
        Case "5": RemoveBizNoDashesToCopiedColumns
        Case "6": BizNo_MergeThreeColumns
        Case "7": CompanyName_CleanWithOptions
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "값 정리", "1~7 중 하나를 입력하세요."
    End Select
End Sub

Public Sub Toolkit_LaunchAddressSearch()
    Dim choice As String
    choice = Toolkit_LauncherPrompt( _
        "주소/검색", _
        "실행할 주소/검색 기능을 선택하세요." & vbCrLf & vbCrLf & _
        "1. 주소 유형 판별" & vbCrLf & _
        "2. 주소 체계 분리" & vbCrLf & _
        "3. 주소 통합" & vbCrLf & _
        "4. 행정구역 분해" & vbCrLf & _
        "5. 네이버 주소 검증" & vbCrLf & _
        "6. 네이버 신규 리스트 생성" & vbCrLf & _
        "7. 네이버 인증키 설정" & vbCrLf & _
        "8. 행정동 마스터 양식 생성" & vbCrLf & _
        "9. 법정동 마스터 양식 생성", _
        "1")

    Select Case choice
        Case "1": Address_ClassifyAddressType
        Case "2": Address_SplitDongAndRoad
        Case "3": Address_CombineDongAndRoad
        Case "4": Address_SplitAdministrativeParts
        Case "5": NaverLocal_VerifyExistingData
        Case "6": NaverLocal_BuildListByDong
        Case "7": NaverLocal_SetApiCredentials
        Case "8": NaverLocal_CreateDongMasterTemplate
        Case "9": NaverLocal_CreateLegalDongMasterTemplate
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "주소/검색", "1~9 중 하나를 입력하세요."
    End Select
End Sub

Public Sub Toolkit_LaunchRosterQuality()
    Dim choice As String
    choice = Toolkit_LauncherPrompt( _
        "명부 품질", _
        "실행할 명부 품질 기능을 선택하세요." & vbCrLf & vbCrLf & _
        "1. 품질 진단" & vbCrLf & _
        "2. 중복 탐지" & vbCrLf & _
        "3. 업체명 매칭" & vbCrLf & _
        "4. 전화번호 검사" & vbCrLf & _
        "5. 사업자등록 상태조회", _
        "1")

    Select Case choice
        Case "1": Roster_QualityReport
        Case "2": Roster_DetectDuplicates
        Case "3": Roster_MatchCompanyNames
        Case "4": Phone_ValidateSelection
        Case "5": BizNo_QueryStatusForSelection
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "명부 품질", "1~5 중 하나를 입력하세요."
    End Select
End Sub

Public Sub Toolkit_LaunchSpssTools()
    Dim choice As String
    choice = Toolkit_LauncherPrompt( _
        "SPSS 표 도구", _
        "실행할 SPSS 표 정리 기능을 선택하세요." & vbCrLf & vbCrLf & _
        "1. 목차 생성" & vbCrLf & _
        "2. 표/배너 재구성" & vbCrLf & _
        "3. 세로 배너 삭제" & vbCrLf & _
        "4. 표 메타 추출", _
        "2")

    Select Case choice
        Case "1": SPSS_CreateIndexForActiveSheet
        Case "2": SPSS_RebuildTablesFromIndex
        Case "3": SPSS_DeleteVerticalBannersFromIndex
        Case "4": SPSS_ExtractTableMetadata
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "SPSS 표 도구", "1~4 중 하나를 입력하세요."
    End Select
End Sub

Public Sub Toolkit_LaunchSampleDesign()
    Dim choice As String
    choice = Toolkit_LauncherPrompt( _
        "표본설계", _
        "실행할 표본설계 기능을 선택하세요." & vbCrLf & vbCrLf & _
        "1. 설계 시트 생성" & vbCrLf & _
        "2. 할당 계산" & vbCrLf & _
        "3. 진행률 관리표 생성" & vbCrLf & _
        "4. 쿼터 점검" & vbCrLf & _
        "5. 모집단 구성표 리스트화", _
        "1")

    Select Case choice
        Case "1": SampleDesign_CreateTemplate
        Case "2": SampleDesign_RunAllocation
        Case "3": SampleDesign_CreateProgressTracker
        Case "4": SampleDesign_ReviewQuota
        Case "5": SampleDesign_ListifyPopulationTable
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "표본설계", "1~5 중 하나를 입력하세요."
    End Select
End Sub

Private Function Toolkit_LauncherPrompt(ByVal title As String, ByVal promptText As String, ByVal defaultChoice As String) As String
    On Error GoTo Fallback

    Dim optionsText As String
    optionsText = Toolkit_OptionsFromNumberedPrompt(promptText)
    If Len(optionsText) > 0 Then
        Toolkit_LauncherPrompt = Toolkit_ButtonMenu(title, Toolkit_FirstPromptLine(promptText), optionsText, defaultChoice)
    Else
        Toolkit_LauncherPrompt = ""
    End If
    Exit Function

Fallback:
    Toolkit_LauncherPrompt = ""
End Function

Private Sub Toolkit_ShowLauncherError(ByVal title As String, ByVal message As String)
    MsgBox message, vbExclamation, title
End Sub

Private Function Toolkit_ButtonMenu(ByVal title As String, ByVal guideText As String, _
                                    ByVal optionsText As String, ByVal defaultChoice As String) As String
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmToolkitButtonMenu")
    form.Configure title, guideText, optionsText, defaultChoice
    form.Show
    Toolkit_ButtonMenu = CStr(form.SelectedKey)
    Unload form
    Exit Function

Fallback:
    Toolkit_ButtonMenu = ""
End Function

Private Function Toolkit_ButtonMenuFallback(ByVal title As String, ByVal guideText As String, _
                                            ByVal optionsText As String, ByVal defaultChoice As String) As String
    Toolkit_ButtonMenuFallback = ""
End Function

Private Function Toolkit_OptionsTextToPrompt(ByVal optionsText As String) As String
    Dim lines As Variant
    lines = Split(optionsText, vbLf)

    Dim result As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(CStr(lines(i)))
        If Len(lineText) > 0 Then
            Dim parts As Variant
            parts = Split(lineText, vbTab)
            If UBound(parts) >= 1 Then
                result = result & CStr(parts(0)) & ". " & CStr(parts(1)) & vbCrLf
            End If
        End If
    Next i
    Toolkit_OptionsTextToPrompt = result
End Function

Private Function Toolkit_OptionsFromNumberedPrompt(ByVal promptText As String) As String
    On Error GoTo SafeExit

    Dim normalizedText As String
    Dim lines As Variant
    Dim result As String
    Dim i As Long

    normalizedText = Replace(CStr(promptText), vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)

    If Len(normalizedText) = 0 Then Exit Function

    lines = Split(normalizedText, vbLf)

    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        Dim sepPos As Long
        Dim keyText As String
        Dim labelText As String

        lineText = Trim$(CStr(lines(i)))
        If Len(lineText) > 0 Then
            sepPos = InStr(1, lineText, ". ", vbTextCompare)
            If sepPos = 0 Then sepPos = InStr(1, lineText, ") ", vbTextCompare)

            If sepPos > 1 Then
                keyText = Trim$(Left$(lineText, sepPos - 1))
                labelText = Trim$(Mid$(lineText, sepPos + 2))
                If IsNumeric(keyText) And Len(labelText) > 0 Then
                    If Len(result) > 0 Then result = result & vbLf
                    result = result & keyText & vbTab & labelText
                End If
            End If
        End If
    Next i

SafeExit:
    Toolkit_OptionsFromNumberedPrompt = result
End Function

Private Function Toolkit_FirstPromptLine(ByVal promptText As String) As String
    On Error GoTo Fallback

    Dim normalizedText As String
    Dim lines As Variant

    normalizedText = Replace(CStr(promptText), vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    lines = Split(normalizedText, vbLf)

    If UBound(lines) >= LBound(lines) Then
        Toolkit_FirstPromptLine = Trim$(CStr(lines(LBound(lines))))
    End If
    If Len(Toolkit_FirstPromptLine) = 0 Then Toolkit_FirstPromptLine = "실행할 기능을 선택하세요."
    Exit Function

Fallback:
    Toolkit_FirstPromptLine = "실행할 기능을 선택하세요."
End Function

Private Function Toolkit_CheckListMenu(ByVal title As String, ByVal guideText As String, _
                                       ByVal optionsText As String, ByVal defaultCsv As String) As String
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmToolkitCheckListMenu")
    form.Configure title, guideText, optionsText, defaultCsv
    form.Show
    Toolkit_CheckListMenu = CStr(form.SelectedKeys)
    Unload form
    Exit Function

Fallback:
    Toolkit_CheckListMenu = ""
End Function

' ============================================================
' Ribbon callbacks
' 설  명  : .xlam 리본 탭 버튼에서 기존 매크로를 호출하기 위한 콜백
' ============================================================
Public Sub Ribbon_LaunchDataCleanup(control As IRibbonControl)
    Toolkit_LaunchDataCleanup
End Sub

Public Sub Ribbon_LaunchAddressSearch(control As IRibbonControl)
    Toolkit_LaunchAddressSearch
End Sub

Public Sub Ribbon_LaunchRosterQuality(control As IRibbonControl)
    Toolkit_LaunchRosterQuality
End Sub

Public Sub Ribbon_LaunchSpssTools(control As IRibbonControl)
    Toolkit_LaunchSpssTools
End Sub

Public Sub Ribbon_LaunchSampleDesign(control As IRibbonControl)
    Toolkit_LaunchSampleDesign
End Sub

Public Sub Ribbon_ToolkitAbout(control As IRibbonControl)
    Toolkit_About
End Sub

Public Sub Ribbon_ResetAllSheetsToTop(control As IRibbonControl)
    Toolkit_ResetAllSheetsToTop
End Sub

Public Sub Ribbon_SelectA1OnAllSheets(control As IRibbonControl)
    Toolkit_SelectA1OnAllSheets
End Sub

Public Sub Ribbon_ConvertTextNumbersInSelection(control As IRibbonControl)
    ConvertTextNumbersWithForm
End Sub

Public Sub Ribbon_ConvertTextNumbersInUsedRange(control As IRibbonControl)
    ConvertTextNumbersInCurrentSheetDataRange
End Sub

Public Sub Ribbon_ConvertTextNumbersInColumns(control As IRibbonControl)
    ConvertTextNumbersWithForm
End Sub

Public Sub Ribbon_NormalizePhonesInSelection(control As IRibbonControl)
    NormalizePhonesWithForm
End Sub

Public Sub Ribbon_NormalizePhonesInColumns(control As IRibbonControl)
    NormalizePhonesWithForm
End Sub

Public Sub Ribbon_RemovePhoneDashes(control As IRibbonControl)
    RemovePhoneDashesToCopiedColumns
End Sub

Public Sub Ribbon_CompanyNameClean(control As IRibbonControl)
    CompanyName_CleanWithOptions
End Sub

Public Sub Ribbon_AddressConvertOldToRoad(control As IRibbonControl)
    Address_ClassifyAddressType
End Sub

Public Sub Ribbon_AddressSplitDongAndRoad(control As IRibbonControl)
    Address_SplitDongAndRoad
End Sub

Public Sub Ribbon_AddressCombineDongAndRoad(control As IRibbonControl)
    Address_CombineDongAndRoad
End Sub

Public Sub Ribbon_BizNoSetApiKey(control As IRibbonControl)
    BizNo_SetApiKey
End Sub

Public Sub Ribbon_RemoveBizNoDashes(control As IRibbonControl)
    RemoveBizNoDashesToCopiedColumns
End Sub

Public Sub Ribbon_BizNoMergeThreeColumns(control As IRibbonControl)
    BizNo_MergeThreeColumns
End Sub

Public Sub Ribbon_BizNoQueryStatus(control As IRibbonControl)
    BizNo_QueryStatusForSelection
End Sub

Public Sub Ribbon_NaverLocalSetApiCredentials(control As IRibbonControl)
    NaverLocal_SetApiCredentials
End Sub

Public Sub Ribbon_NaverLocalBuildListByDong(control As IRibbonControl)
    NaverLocal_BuildListByDong
End Sub

Public Sub Ribbon_NaverLocalVerifyExistingData(control As IRibbonControl)
    NaverLocal_VerifyExistingData
End Sub

Public Sub Ribbon_NaverLocalCreateDongMasterTemplate(control As IRibbonControl)
    NaverLocal_CreateDongMasterTemplate
End Sub

Public Sub Ribbon_NaverLocalCreateLegalDongMasterTemplate(control As IRibbonControl)
    NaverLocal_CreateLegalDongMasterTemplate
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

Public Sub Ribbon_SPSSCreateIndexForActiveSheet(control As IRibbonControl)
    SPSS_CreateIndexForActiveSheet
End Sub

Public Sub Ribbon_SPSSRebuildTablesFromIndex(control As IRibbonControl)
    SPSS_RebuildTablesFromIndex
End Sub

Public Sub Ribbon_SPSSDeleteVerticalBanners(control As IRibbonControl)
    SPSS_DeleteVerticalBannersFromIndex
End Sub

Public Sub Ribbon_RosterDetectDuplicates(control As IRibbonControl)
    Roster_DetectDuplicates
End Sub

Public Sub Ribbon_RosterQualityReport(control As IRibbonControl)
    Roster_QualityReport
End Sub

Public Sub Ribbon_RosterMatchCompanyNames(control As IRibbonControl)
    Roster_MatchCompanyNames
End Sub

Public Sub Ribbon_AddressSplitAdministrativeParts(control As IRibbonControl)
    Address_SplitAdministrativeParts
End Sub

Public Sub Ribbon_PhoneValidateSelection(control As IRibbonControl)
    Phone_ValidateSelection
End Sub

Public Sub Ribbon_SPSSExtractTableMetadata(control As IRibbonControl)
    SPSS_ExtractTableMetadata
End Sub

Public Sub Ribbon_SampleDesignCreateTemplate(control As IRibbonControl)
    SampleDesign_CreateTemplate
End Sub

Public Sub Ribbon_SampleDesignListifyPopulationTable(control As IRibbonControl)
    SampleDesign_ListifyPopulationTable
End Sub

Public Sub Ribbon_SampleDesignRunAllocation(control As IRibbonControl)
    SampleDesign_RunAllocation
End Sub

Public Sub Ribbon_SampleDesignCreateProgressTracker(control As IRibbonControl)
    SampleDesign_CreateProgressTracker
End Sub

Public Sub Ribbon_DatabaseMergeSheets(control As IRibbonControl)
    Database_MergeSheetsByHeaders
End Sub

Public Sub Ribbon_DatabaseSplitMergedSheet(control As IRibbonControl)
    Database_SplitMergedSheetBySource
End Sub
