Attribute VB_Name = "MacroToolkit"
' ============================================================
' 모듈명  : MacroToolkit
' 설  명  : 엑셀 데이터 편집을 위한 매크로 도구 모음
'           - 텍스트 형식 숫자를 실제 숫자로 일괄 변환
'           - 혼용된 한국 전화번호 형식을 표준 형식으로 정규화
' ============================================================
Option Explicit

' 현재 툴킷 버전 (외부 모듈에서도 참조 가능하도록 Public 선언)
Public Const TOOLKIT_VERSION As String = "1.1.40"

' 버전 정보를 기록할 숨김 시트 이름 (언더스코어 접두어로 내부 시트임을 표시)
Private Const VERSION_SHEET As String = "_MacroToolkit"
Private Const LOG_SHEET As String = "_MacroToolkit_Log"
Private Const BIZ_API_BASE_URL As String = "https://api.odcloud.kr/api/nts-businessman/v1/status"
Private Const BIZ_APP_NAME As String = "DataEditingMacroToolkit"
Private Const BIZ_APP_SECTION As String = "BusinessApi"
Private Const BIZ_APP_KEY As String = "ServiceKey"
Private Const NAVER_API_BASE_URL As String = "https://openapi.naver.com/v1/search/local.json"
Private Const NAVER_WEB_API_BASE_URL As String = "https://openapi.naver.com/v1/search/webkr.json"
Private Const NAVER_APP_NAME As String = "DataEditingMacroToolkit"
Private Const NAVER_APP_SECTION As String = "NaverLocalSearch"
Private Const NAVER_CLIENT_ID_KEY As String = "ClientId"
Private Const NAVER_CLIENT_SECRET_KEY As String = "ClientSecret"
Private Const NAVER_BUILD_INDUSTRY_KEY As String = "BuildIndustry"
Private Const NAVER_BUILD_DISPLAY_KEY As String = "BuildDisplayCount"
Private Const NAVER_BUILD_DELAY_KEY As String = "BuildDelaySeconds"
Private Const NAVER_BUILD_SORT_KEY As String = "BuildSortMode"
Private Const NAVER_BUILD_BASIS_KEY As String = "BuildDongBasis"
Private Const NAVER_BUILD_SIDO_KEY As String = "BuildSidoFilter"
Private Const NAVER_BUILD_SIGUNGU_KEY As String = "BuildSigunguFilter"
Private Const NAVER_MALL_DISPLAY_KEY As String = "OfficialMallDisplayCount"
Private Const NAVER_MALL_DELAY_KEY As String = "OfficialMallDelaySeconds"
Private Const NAVER_DONG_MASTER_SHEET As String = "_NaverLocalDongMaster"
Private Const NAVER_USER_DONG_MASTER_SHEET As String = "행정동_마스터"
Private Const NAVER_LEGAL_DONG_MASTER_SHEET As String = "_NaverLocalLegalDongMaster"
Private Const NAVER_USER_LEGAL_DONG_MASTER_SHEET As String = "법정동_마스터"
Private Const API_RESOLVE_TIMEOUT_MS As Long = 5000
Private Const API_CONNECT_TIMEOUT_MS As Long = 10000
Private Const API_SEND_TIMEOUT_MS As Long = 15000
Private Const API_RECEIVE_TIMEOUT_MS As Long = 30000
Private Const TOOLKIT_PROGRESS_ROW_THRESHOLD As Long = 1000
Private Const TOOLKIT_PROGRESS_BATCH_ROWS As Long = 500
Private Const TOOLKIT_PROGRESS_PAIR_BATCH As Long = 10000
Private Const TOOLKIT_CANCEL_ERROR_NUMBER As Long = vbObjectError + 9130

Private mToolkitOperationDepth As Long
Private mToolkitPrevScreenUpdating As Boolean
Private mToolkitPrevEnableEvents As Boolean
Private mToolkitPrevDisplayAlerts As Boolean
Private mToolkitPrevCalculation As XlCalculation
Private mToolkitPrevStatusBar As Variant

Private gProgressForm As Object
Private mToolkitCancelRequested As Boolean

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
           "- 네이버 API 기반 업체 리스트 생성/주소 검증/공식몰 링크 검색" & vbCrLf & _
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
    ws.Range("A44").Value = "Database_CreateMergeMarkerRows"
    ws.Range("A45").Value = "Database_SplitMergedSheetBySource"
    ws.Range("A46").Value = "NaverLocal_CreateLegalDongMasterTemplate"
    ws.Range("A47").Value = "Address_CombineDongAndRoad"
    ws.Range("A48").Value = "SampleDesign_PivotPopulationList"
    ws.Range("A49").Value = "SampleDesign_ShowGuide"
    ws.Range("A50").Value = "SampleDesign_CreateBandColumn"
    ws.Range("A51").Value = "Roster_CreateValidationWorkspace"
    ws.Range("A52").Value = "NaverOfficialMall_FindLinksForSelection"

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
    mToolkitCancelRequested = False
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
    Dim cancelRequested As Boolean

    On Error Resume Next
    If Not gProgressForm Is Nothing Then
        gProgressForm.UpdateProgress currentCount, totalCount, message
        cancelRequested = CBool(gProgressForm.IsCancelled)
    End If

    Dim percentValue As Double
    If totalCount > 0 Then percentValue = currentCount / totalCount
    Application.StatusBar = Format(percentValue, "0%") & " - " & message & " (" & currentCount & "/" & totalCount & ")"
    DoEvents
    If Not gProgressForm Is Nothing Then cancelRequested = cancelRequested Or CBool(gProgressForm.IsCancelled)
    On Error GoTo 0
    If cancelRequested Then mToolkitCancelRequested = True
    Toolkit_RaiseIfCancelled
End Sub

Private Sub Toolkit_ProgressClose()
    On Error Resume Next
    If Not gProgressForm Is Nothing Then
        Unload gProgressForm
        Set gProgressForm = Nothing
    End If
    Application.StatusBar = False
    DoEvents
    mToolkitCancelRequested = False
    On Error GoTo 0
End Sub

Private Function Toolkit_ShouldShowRowProgress(ByVal totalRows As Long) As Boolean
    Toolkit_ShouldShowRowProgress = (totalRows > 0)
End Function

Private Sub Toolkit_RowProgressStart(ByVal title As String, ByVal totalRows As Long, ByVal message As String)
    If Toolkit_ShouldShowRowProgress(totalRows) Then Toolkit_ProgressShow title, message, totalRows
End Sub

Private Sub Toolkit_RowProgressUpdate(ByVal processedRows As Long, ByVal totalRows As Long, ByVal message As String)
    If Not Toolkit_ShouldShowRowProgress(totalRows) Then Exit Sub
    Toolkit_ProgressUpdate processedRows, totalRows, message
End Sub

Private Sub Toolkit_RowProgressClose(ByVal totalRows As Long)
    If Toolkit_ShouldShowRowProgress(totalRows) Then Toolkit_ProgressClose
End Sub

Private Function Toolkit_ProgressCancelRequested() As Boolean
    Dim formCancelled As Boolean

    On Error Resume Next
    If Not gProgressForm Is Nothing Then formCancelled = CBool(gProgressForm.IsCancelled)
    On Error GoTo 0

    If formCancelled Then mToolkitCancelRequested = True
    Toolkit_ProgressCancelRequested = mToolkitCancelRequested
End Function

Private Function Toolkit_IsCancelError(ByVal errorNumber As Long) As Boolean
    Toolkit_IsCancelError = (errorNumber = TOOLKIT_CANCEL_ERROR_NUMBER)
End Function

Private Function Toolkit_ShowCancelMessage(ByVal title As String) As Boolean
    If Not Toolkit_IsCancelError(Err.Number) Then Exit Function
    MsgBox "작업이 취소되었습니다.", vbInformation, title
    Toolkit_ShowCancelMessage = True
End Function

Private Sub Toolkit_RaiseIfCancelled()
    If Toolkit_ProgressCancelRequested() Then
        Err.Raise TOOLKIT_CANCEL_ERROR_NUMBER, "DataEditingMacroToolkit", "사용자가 작업을 취소했습니다."
    End If
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
    Dim processedCells As Long
    Dim totalCells As Long
    totalCells = target.Cells.CountLarge

    On Error GoTo ErrHandler
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
    Toolkit_RowProgressClose totalCells
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("텍스트 숫자 변환") Then Exit Sub
    MsgBox "오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, "텍스트 숫자 변환"
End Sub

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
                  "mobile_010" & vbTab & "8자리 값을 휴대폰 개인전화 010-XXXX-XXXX 형식으로 정리"

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

    NormalizePhonesInRange outputRange, "전화번호 정리 복사본", options
End Sub

' ============================================================
' 프로시저 : NormalizePhonesInRange (내부 전용)
' 설  명  : 주어진 Range 내 각 셀을 순회하며
'           한국 전화번호를 표준 형식으로 정규화한다.
'           셀 서식을 텍스트(@)로 변경해 앞자리 0이 유지되도록 한다.
'           수식 셀은 건너뛰고 완료 후 통계를 표시한다.
' 파라미터 : target - 정규화 대상 Range
' ============================================================
Private Sub NormalizePhonesInRange(ByVal target As Range, Optional ByVal resultTitle As String = "전화번호 정규화", Optional ByVal normalizeOptions As Variant)
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
    MsgBox changedCount & "개 전화번호가 정규화되었습니다." & vbCrLf & _
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
    options("removeBranchMarkers") = CBool(form.RemoveBranchMarkers)
    options("removeLegalEntityMarkers") = CBool(form.RemoveLegalEntityMarkers)
    options("removeParentheticalText") = CBool(form.RemoveParentheticalText)
    options("removeSpecialChars") = CBool(form.RemoveSpecialChars)
    options("pronounceEnglishWords") = CBool(form.PronounceEnglishWords)
    options("spellAlphabetKorean") = CBool(form.SpellAlphabetKorean)
    options("removeBilingualDuplicates") = CBool(form.RemoveBilingualDuplicates)

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
    options("cutBeforePossessive") = False
    options("removeBranchMarkers") = False
    options("removeLegalEntityMarkers") = False
    options("removeParentheticalText") = False
    options("removeSpecialChars") = False
    options("pronounceEnglishWords") = False
    options("spellAlphabetKorean") = True
    options("removeBilingualDuplicates") = True
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

    Dim sourceCell As Range
    Dim outputCell As Range
    Dim changedCount As Long
    Dim skippedFormulaCount As Long
    Dim blankPreventedCount As Long
    Dim cleaned As String
    Dim originalText As String
    Dim totalRows As Long
    totalRows = outputRange.Rows.Count

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "기업체명 정리", totalRows, "정리 준비 중"
    Dim relRow As Long, relCol As Long
    For relRow = 1 To outputRange.Rows.Count
        For relCol = 1 To outputRange.Columns.Count
            Set sourceCell = workRange.Cells(relRow, relCol)
            Set outputCell = outputRange.Cells(relRow, relCol)
            If sourceCell.HasFormula Then
                skippedFormulaCount = skippedFormulaCount + 1
            Else
                originalText = CStr(sourceCell.Value)
                cleaned = CompanyName_CleanOne(originalText, options)
                If Len(Trim$(originalText)) > 0 And Len(Trim$(cleaned)) = 0 Then
                    cleaned = originalText
                    blankPreventedCount = blankPreventedCount + 1
                End If
                If originalText <> cleaned Then
                    outputCell.NumberFormat = "@"
                    outputCell.Value = cleaned
                    changedCount = changedCount + 1
                End If
            End If
        Next relCol
        Toolkit_RowProgressUpdate relRow, totalRows, "기업체명 정리 중: " & CStr(relRow) & "/" & CStr(totalRows) & "행"
    Next relRow

    outputRange.Columns.AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox changedCount & "개 기업체명/사업체명을 정리했습니다." & vbCrLf & _
           blankPreventedCount & "개 비공란 원본은 공란화를 방지해 원본값을 유지했습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다." & vbCrLf & _
           "원본 범위는 변경하지 않았습니다.", vbInformation, "기업체명 정리"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("기업체명 정리") Then Exit Sub
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
    If CBool(options("removeParentheticalText")) Then textValue = CompanyName_RemoveParentheticalText(textValue)
    If CBool(options("removeBilingualDuplicates")) Then textValue = CompanyName_RemoveBilingualDuplicates(textValue)
    If CBool(options("removeSpecialChars")) And Not CBool(options("removeLegalEntityMarkers")) Then textValue = CompanyName_ReplaceBusinessSymbols(textValue)
    If CBool(options("pronounceEnglishWords")) Then textValue = CompanyName_PronounceEnglishWords(textValue)
    If CBool(options("removeSpaces")) And Not CBool(options("pronounceEnglishWords")) Then textValue = CompanyName_RemoveSpaces(textValue)
    If CBool(options("normalizeCorpSymbol")) Then textValue = CompanyName_NormalizeCorpSymbol(textValue)
    If CBool(options("normalizeCorpText")) Then textValue = Replace(textValue, "주식회사", "(주)")
    If CBool(options("moveCorpToFront")) Then textValue = CompanyName_MoveCorpMarker(textValue, True)
    If CBool(options("moveCorpToEnd")) Then textValue = CompanyName_MoveCorpMarker(textValue, False)
    If CBool(options("removeCorpMarkers")) Then textValue = CompanyName_RemoveCorpMarkers(textValue)
    If CBool(options("removeBranchMarkers")) Then textValue = CompanyName_RemoveBranchMarkers(textValue)
    If CBool(options("removeLegalEntityMarkers")) Then textValue = CompanyName_RemoveLegalEntityMarkers(textValue)
    If CBool(options("removeSpecialChars")) Then textValue = CompanyName_RemoveSpecialChars(textValue)
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
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "[\(（]\s*주\s*[\)）]|㈜|주식회사"
    CompanyName_HasCorpMarker = re.Test(textValue)
End Function

Private Function CompanyName_RemoveCorpMarkers(ByVal textValue As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\(（]\s*주\s*[\)）]|㈜|주식회사"
    CompanyName_RemoveCorpMarkers = Trim$(re.Replace(textValue, ""))
End Function

Private Function CompanyName_RemoveParentheticalText(ByVal textValue As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\(（][^\)）]*[\)）]"
    CompanyName_RemoveParentheticalText = Trim$(re.Replace(textValue, ""))
End Function

Private Function CompanyName_RemoveBranchMarkers(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\(（]\s*(본점|지점|[A-Za-z0-9가-힣]{1,4}점)\s*[\)）]"
    result = Trim$(re.Replace(result, ""))

    result = CompanyName_RemoveLeadingBranchMarker(result)
    result = CompanyName_RemoveTrailingBranchMarker(result)

    CompanyName_RemoveBranchMarkers = Trim$(result)
End Function

Private Function CompanyName_RemoveLegalEntityMarkers(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)
    If Len(result) = 0 Then Exit Function

    result = CompanyName_RemoveKoreanLegalEntityMarkers(result)
    result = CompanyName_RemoveForeignLegalEntityMarkers(result)
    result = CompanyName_CleanTrailingSeparators(result)

    CompanyName_RemoveLegalEntityMarkers = Trim$(result)
End Function

Private Function CompanyName_RemoveKoreanLegalEntityMarkers(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\(（]\s*(유|유한|유한회사|유한책임회사|합자|합명)\s*[\)）]"
    result = Trim$(re.Replace(result, ""))

    Dim markers As Variant
    markers = Array("유한책임회사", "유한회사", "합자회사", "합명회사", "사단법인", "재단법인", _
                    "학교법인", "의료법인", "농업회사법인", "영농조합법인", "사회복지법인", _
                    "법무법인", "세무법인", "회계법인", "특허법인")

    Dim i As Long
    For i = LBound(markers) To UBound(markers)
        result = CompanyName_RemoveLegalMarkerAtEdges(result, CStr(markers(i)))
    Next i

    CompanyName_RemoveKoreanLegalEntityMarkers = Trim$(result)
End Function

Private Function CompanyName_RemoveLegalMarkerAtEdges(ByVal textValue As String, ByVal markerText As String) As String
    Dim result As String
    result = Trim$(textValue)

    Do While Len(result) > Len(markerText) And Left$(result, Len(markerText)) = markerText
        result = Trim$(Mid$(result, Len(markerText) + 1))
        result = CompanyName_CleanLeadingSeparators(result)
    Loop

    Do While Len(result) > Len(markerText) And Right$(result, Len(markerText)) = markerText
        result = Trim$(Left$(result, Len(result) - Len(markerText)))
        result = CompanyName_CleanTrailingSeparators(result)
    Loop

    CompanyName_RemoveLegalMarkerAtEdges = Trim$(result)
End Function

Private Function CompanyName_RemoveForeignLegalEntityMarkers(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\s,\.·ㆍ・\-_/\(\)（）]*(co\.?\s*,?\s*ltd\.?|company\s+limited|pvt\.?\s*ltd\.?|pte\.?\s*ltd\.?|ltd\.?|limited|inc\.?|incorporated|corp\.?|corporation|llc|l\.l\.c\.?|llp|l\.l\.p\.?|plc|gmbh)[\s,\.·ㆍ・\-_/\(\)（）]*$"

    Do While re.Test(result)
        result = Trim$(re.Replace(result, ""))
    Loop

    re.Pattern = "[\s,\.·ㆍ・\-_/\(\)（）]+(s\.?a\.?|ag)[\s,\.·ㆍ・\-_/\(\)（）]*$"
    Do While re.Test(result)
        result = Trim$(re.Replace(result, ""))
    Loop

    CompanyName_RemoveForeignLegalEntityMarkers = Trim$(result)
End Function

Private Function CompanyName_CleanLeadingSeparators(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)
    Do While Len(result) > 0 And CompanyName_IsLegalSeparator(Left$(result, 1))
        result = Trim$(Mid$(result, 2))
    Loop
    CompanyName_CleanLeadingSeparators = result
End Function

Private Function CompanyName_CleanTrailingSeparators(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)
    Do While Len(result) > 0 And CompanyName_IsLegalSeparator(Right$(result, 1))
        result = Trim$(Left$(result, Len(result) - 1))
    Loop
    CompanyName_CleanTrailingSeparators = result
End Function

Private Function CompanyName_IsLegalSeparator(ByVal ch As String) As Boolean
    CompanyName_IsLegalSeparator = (ch = " " Or ch = "," Or ch = "." Or ch = "-" Or ch = "_" Or _
                                   ch = "/" Or ch = "·" Or ch = "ㆍ" Or ch = "・" Or _
                                   ch = "(" Or ch = ")" Or ch = "（" Or ch = "）")
End Function

Private Function CompanyName_RemoveLeadingBranchMarker(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "^(본점|지점)[\s_\-]+"
    result = re.Replace(result, "")

    CompanyName_RemoveLeadingBranchMarker = Trim$(result)
End Function

Private Function CompanyName_RemoveTrailingBranchMarker(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)
    If Len(result) = 0 Then Exit Function

    Dim knownResult As String
    If CompanyName_RemoveKnownTrailingBranch(result, "본점", knownResult) Then
        CompanyName_RemoveTrailingBranchMarker = knownResult
        Exit Function
    End If
    If CompanyName_RemoveKnownTrailingBranch(result, "지점", knownResult) Then
        CompanyName_RemoveTrailingBranchMarker = knownResult
        Exit Function
    End If

    If Right$(result, 1) = "점" Then
        Dim branchLen As Long
        For branchLen = 4 To 1 Step -1
            Dim suffixLen As Long
            suffixLen = branchLen + 1
            If Len(result) > suffixLen Then
                Dim baseName As String
                Dim suffixText As String
                baseName = Left$(result, Len(result) - suffixLen)
                suffixText = Right$(result, suffixLen)
                If CompanyName_VisibleLength(baseName) >= 2 And CompanyName_IsBranchSuffix(suffixText) Then
                    CompanyName_RemoveTrailingBranchMarker = Trim$(baseName)
                    Exit Function
                End If
            End If
        Next branchLen
    End If

    CompanyName_RemoveTrailingBranchMarker = result
End Function

Private Function CompanyName_RemoveKnownTrailingBranch(ByVal textValue As String, ByVal suffixText As String, ByRef resultOut As String) As Boolean
    If Len(textValue) <= Len(suffixText) Then Exit Function
    If Right$(textValue, Len(suffixText)) <> suffixText Then Exit Function

    Dim baseName As String
    baseName = Trim$(Left$(textValue, Len(textValue) - Len(suffixText)))
    If CompanyName_VisibleLength(baseName) < 2 Then Exit Function

    resultOut = baseName
    CompanyName_RemoveKnownTrailingBranch = True
End Function

Private Function CompanyName_IsBranchSuffix(ByVal suffixText As String) As Boolean
    If Len(suffixText) < 2 Or Len(suffixText) > 5 Then Exit Function
    If Right$(suffixText, 1) <> "점" Then Exit Function

    Dim stem As String
    stem = Left$(suffixText, Len(suffixText) - 1)
    CompanyName_IsBranchSuffix = (CompanyName_VisibleLength(stem) >= 1 And CompanyName_VisibleLength(stem) <= 4)
End Function

Private Function CompanyName_VisibleLength(ByVal textValue As String) As Long
    Dim i As Long
    For i = 1 To Len(textValue)
        Dim ch As String
        ch = Mid$(textValue, i, 1)
        If CompanyName_IsNameCharacter(ch) And ch <> " " Then CompanyName_VisibleLength = CompanyName_VisibleLength + 1
    Next i
End Function

Private Function CompanyName_NormalizeCorpSymbol(ByVal textValue As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\(（]\s*주\s*[\)）]"
    textValue = re.Replace(textValue, "(주)")
    textValue = Replace(textValue, "㈜", "(주)")
    CompanyName_NormalizeCorpSymbol = textValue
End Function

Private Function CompanyName_ReplaceBusinessSymbols(ByVal textValue As String) As String
    Dim result As String
    result = CStr(textValue)

    result = Replace(result, "&", "앤")
    result = Replace(result, "＆", "앤")

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(^|[^A-Za-z])co\.(?=$|[^A-Za-z])"
    result = re.Replace(result, "$1컴패니")

    re.IgnoreCase = False
    re.Pattern = "(Co|CO)\.$"
    result = re.Replace(result, "컴패니")

    CompanyName_ReplaceBusinessSymbols = result
End Function

Private Function CompanyName_RemoveSpecialChars(ByVal textValue As String) As String
    textValue = CompanyName_ReplaceBusinessSymbols(textValue)

    Dim result As String
    Dim i As Long
    For i = 1 To Len(textValue)
        Dim ch As String
        ch = Mid$(textValue, i, 1)
        If CompanyName_IsNameCharacter(ch) Then result = result & ch
    Next i
    CompanyName_RemoveSpecialChars = Trim$(result)
End Function

Private Function CompanyName_IsNameCharacter(ByVal ch As String) As Boolean
    Dim codeValue As Long
    codeValue = AscW(ch)
    If codeValue < 0 Then codeValue = codeValue + 65536

    CompanyName_IsNameCharacter = _
        ((codeValue >= 48 And codeValue <= 57) Or _
         (codeValue >= 65 And codeValue <= 90) Or _
         (codeValue >= 97 And codeValue <= 122) Or _
         (codeValue >= &HAC00& And codeValue <= &HD7A3&) Or _
         (codeValue >= &H1100& And codeValue <= &H11FF&) Or _
         (codeValue >= &H3130& And codeValue <= &H318F&) Or _
         (codeValue >= &H4E00& And codeValue <= &H9FFF&) Or _
         ch = " ")
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

Private Function CompanyName_RemoveBilingualDuplicates(ByVal textValue As String) As String
    Dim result As String
    result = CStr(textValue)

    Dim changed As Boolean
    Do
        changed = False

        Dim i As Long
        i = 1
        Do While i <= Len(result)
            Dim ch As String
            ch = Mid$(result, i, 1)

            If CompanyName_IsHangulLetter(ch) Then
                Dim koreanStart As Long, koreanEnd As Long
                koreanStart = i
                koreanEnd = CompanyName_ReadKoreanBlockEnd(result, koreanStart)

                Dim englishStart As Long
                englishStart = CompanyName_SkipBilingualSeparators(result, koreanEnd + 1)
                If englishStart <= Len(result) And CompanyName_IsAsciiLetterOrDigit(Mid$(result, englishStart, 1)) Then
                    Dim englishEnd As Long
                    englishEnd = CompanyName_ReadEnglishBlockEnd(result, englishStart)
                    If CompanyName_BilingualBlocksMatch(Mid$(result, koreanStart, koreanEnd - koreanStart + 1), _
                                                        Mid$(result, englishStart, englishEnd - englishStart + 1)) Then
                        result = Left$(result, koreanEnd) & Mid$(result, CompanyName_ReadTrailingBilingualSeparators(result, englishEnd) + 1)
                        changed = True
                        Exit Do
                    End If
                End If

                i = koreanEnd + 1
            ElseIf CompanyName_IsAsciiLetterOrDigit(ch) Then
                Dim engStart As Long, engEnd As Long
                engStart = i
                engEnd = CompanyName_ReadEnglishBlockEnd(result, engStart)

                Dim korStart As Long
                korStart = CompanyName_SkipBilingualSeparators(result, engEnd + 1)
                If korStart <= Len(result) And CompanyName_IsHangulLetter(Mid$(result, korStart, 1)) Then
                    Dim korEnd As Long
                    korEnd = CompanyName_ReadKoreanBlockEnd(result, korStart)
                    If CompanyName_BilingualBlocksMatch(Mid$(result, korStart, korEnd - korStart + 1), _
                                                        Mid$(result, engStart, engEnd - engStart + 1)) Then
                        result = Left$(result, engStart - 1) & Mid$(result, korStart)
                        changed = True
                        Exit Do
                    End If
                End If

                i = engEnd + 1
            Else
                i = i + 1
            End If
        Loop
    Loop While changed

    CompanyName_RemoveBilingualDuplicates = Trim$(result)
End Function

Private Function CompanyName_BilingualBlocksMatch(ByVal koreanBlock As String, ByVal englishBlock As String) As Boolean
    Dim koreanText As String
    Dim englishKoreanText As String
    koreanText = CompanyName_KoreanComparable(koreanBlock)
    englishKoreanText = CompanyName_EnglishBlockKoreanComparable(englishBlock)

    If Len(koreanText) < 2 Or Len(englishKoreanText) < 2 Then Exit Function
    CompanyName_BilingualBlocksMatch = (StrComp(koreanText, englishKoreanText, vbTextCompare) = 0)
End Function

Private Function CompanyName_KoreanComparable(ByVal textValue As String) As String
    Dim result As String
    Dim i As Long
    For i = 1 To Len(textValue)
        Dim ch As String
        ch = Mid$(textValue, i, 1)
        If CompanyName_IsHangulLetter(ch) Then result = result & ch
    Next i
    CompanyName_KoreanComparable = result
End Function

Private Function CompanyName_EnglishBlockKoreanComparable(ByVal textValue As String) As String
    Dim result As String
    Dim token As String
    Dim i As Long

    For i = 1 To Len(textValue)
        Dim ch As String
        ch = Mid$(textValue, i, 1)
        If CompanyName_IsAsciiLetter(ch) Then
            token = token & ch
        ElseIf CompanyName_IsAsciiDigit(ch) Then
            If Len(token) > 0 Then
                result = result & CompanyName_SpellAlphabetKorean(CompanyName_PronounceEnglishToken(token))
                token = ""
            End If
            result = result & CompanyName_DigitKorean(ch)
        ElseIf Len(token) > 0 Then
            result = result & CompanyName_SpellAlphabetKorean(CompanyName_PronounceEnglishToken(token))
            token = ""
        End If
    Next i

    If Len(token) > 0 Then result = result & CompanyName_SpellAlphabetKorean(CompanyName_PronounceEnglishToken(token))
    CompanyName_EnglishBlockKoreanComparable = CompanyName_KoreanComparable(result)
End Function

Private Function CompanyName_ReadKoreanBlockEnd(ByVal textValue As String, ByVal startPos As Long) As Long
    Dim pos As Long
    pos = startPos
    Do While pos <= Len(textValue)
        If CompanyName_IsHangulLetter(Mid$(textValue, pos, 1)) Then
            CompanyName_ReadKoreanBlockEnd = pos
            pos = pos + 1
        ElseIf CompanyName_IsBilingualSeparator(Mid$(textValue, pos, 1)) And _
               CompanyName_NextNonBilingualSeparatorPos(textValue, pos + 1) <= Len(textValue) And _
               CompanyName_IsHangulLetter(Mid$(textValue, CompanyName_NextNonBilingualSeparatorPos(textValue, pos + 1), 1)) Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Function

Private Function CompanyName_ReadEnglishBlockEnd(ByVal textValue As String, ByVal startPos As Long) As Long
    Dim pos As Long
    pos = startPos
    Do While pos <= Len(textValue)
        If CompanyName_IsAsciiLetterOrDigit(Mid$(textValue, pos, 1)) Then
            CompanyName_ReadEnglishBlockEnd = pos
            pos = pos + 1
        ElseIf CompanyName_IsBilingualSeparator(Mid$(textValue, pos, 1)) And _
               CompanyName_NextNonBilingualSeparatorPos(textValue, pos + 1) <= Len(textValue) And _
               CompanyName_IsAsciiLetterOrDigit(Mid$(textValue, CompanyName_NextNonBilingualSeparatorPos(textValue, pos + 1), 1)) Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Function

Private Function CompanyName_SkipBilingualSeparators(ByVal textValue As String, ByVal startPos As Long) As Long
    Dim pos As Long
    pos = startPos
    Do While pos <= Len(textValue) And CompanyName_IsBilingualSeparator(Mid$(textValue, pos, 1))
        pos = pos + 1
    Loop
    CompanyName_SkipBilingualSeparators = pos
End Function

Private Function CompanyName_NextNonBilingualSeparatorPos(ByVal textValue As String, ByVal startPos As Long) As Long
    CompanyName_NextNonBilingualSeparatorPos = CompanyName_SkipBilingualSeparators(textValue, startPos)
End Function

Private Function CompanyName_ReadTrailingBilingualSeparators(ByVal textValue As String, ByVal endPos As Long) As Long
    Dim pos As Long
    pos = endPos + 1
    Do While pos <= Len(textValue) And CompanyName_IsBilingualSeparator(Mid$(textValue, pos, 1))
        pos = pos + 1
    Loop
    CompanyName_ReadTrailingBilingualSeparators = pos - 1
End Function

Private Function CompanyName_IsBilingualSeparator(ByVal ch As String) As Boolean
    CompanyName_IsBilingualSeparator = (ch = " " Or ch = vbTab Or ch = "-" Or ch = "_" Or ch = "." Or _
                                       ch = "/" Or ch = "(" Or ch = ")" Or ch = "（" Or ch = "）")
End Function

Private Function CompanyName_IsHangulLetter(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Dim codeValue As Long
    codeValue = AscW(ch)
    If codeValue < 0 Then codeValue = codeValue + 65536
    CompanyName_IsHangulLetter = ((codeValue >= &HAC00& And codeValue <= &HD7A3&) Or _
                                  (codeValue >= &H1100& And codeValue <= &H11FF&) Or _
                                  (codeValue >= &H3130& And codeValue <= &H318F&))
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
                result = result & CompanyName_PronounceEnglishToken(token)
                token = ""
            End If
            result = result & ch
        End If
    Next i

    If Len(token) > 0 Then result = result & CompanyName_PronounceEnglishToken(token)
    CompanyName_PronounceEnglishWords = result
End Function

Private Function CompanyName_PronounceEnglishToken(ByVal tokenText As String) As String
    Dim directKorean As String
    If CompanyName_TryEnglishWordKorean(tokenText, directKorean) Then
        CompanyName_PronounceEnglishToken = directKorean
        Exit Function
    End If

    Dim result As String
    Dim pos As Long
    pos = 1
    Do While pos <= Len(tokenText)
        Dim matchLen As Long
        Dim matchKorean As String
        If CompanyName_LongestEnglishWordKorean(tokenText, pos, matchLen, matchKorean) Then
            result = result & matchKorean
            pos = pos + matchLen
        Else
            result = result & Mid$(tokenText, pos, 1)
            pos = pos + 1
        End If
    Loop

    CompanyName_PronounceEnglishToken = result
End Function

Private Function CompanyName_LongestEnglishWordKorean(ByVal tokenText As String, ByVal startPos As Long, ByRef matchLen As Long, ByRef koreanOut As String) As Boolean
    Dim candidateLen As Long
    For candidateLen = Len(tokenText) - startPos + 1 To 2 Step -1
        Dim candidate As String
        candidate = Mid$(tokenText, startPos, candidateLen)
        If CompanyName_TryEnglishWordKorean(candidate, koreanOut) Then
            matchLen = candidateLen
            CompanyName_LongestEnglishWordKorean = True
            Exit Function
        End If
    Next candidateLen
End Function

Private Function CompanyName_TryEnglishWordKorean(ByVal wordText As String, ByRef koreanOut As String) As Boolean
    koreanOut = CompanyName_EnglishWordKorean(wordText)
    CompanyName_TryEnglishWordKorean = (StrComp(koreanOut, wordText, vbBinaryCompare) <> 0)
End Function

Private Function CompanyName_IsAsciiLetter(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Dim codeValue As Long
    codeValue = AscW(ch)
    CompanyName_IsAsciiLetter = ((codeValue >= 65 And codeValue <= 90) Or (codeValue >= 97 And codeValue <= 122))
End Function

Private Function CompanyName_IsAsciiDigit(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Dim codeValue As Long
    codeValue = AscW(ch)
    CompanyName_IsAsciiDigit = (codeValue >= 48 And codeValue <= 57)
End Function

Private Function CompanyName_IsAsciiLetterOrDigit(ByVal ch As String) As Boolean
    CompanyName_IsAsciiLetterOrDigit = (CompanyName_IsAsciiLetter(ch) Or CompanyName_IsAsciiDigit(ch))
End Function

Private Function CompanyName_DigitKorean(ByVal ch As String) As String
    Select Case ch
        Case "0": CompanyName_DigitKorean = "제로"
        Case "1": CompanyName_DigitKorean = "원"
        Case "2": CompanyName_DigitKorean = "투"
        Case "3": CompanyName_DigitKorean = "쓰리"
        Case "4": CompanyName_DigitKorean = "포"
        Case "5": CompanyName_DigitKorean = "파이브"
        Case "6": CompanyName_DigitKorean = "식스"
        Case "7": CompanyName_DigitKorean = "세븐"
        Case "8": CompanyName_DigitKorean = "에이트"
        Case "9": CompanyName_DigitKorean = "나인"
        Case Else: CompanyName_DigitKorean = ch
    End Select
End Function

Private Function CompanyName_EnglishWordKorean(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "ai": CompanyName_EnglishWordKorean = "에이아이"
        Case "academy": CompanyName_EnglishWordKorean = "아카데미"
        Case "accounting": CompanyName_EnglishWordKorean = "어카운팅"
        Case "ad", "ads": CompanyName_EnglishWordKorean = "애드"
        Case "advertising": CompanyName_EnglishWordKorean = "애드버타이징"
        Case "agency": CompanyName_EnglishWordKorean = "에이전시"
        Case "agri": CompanyName_EnglishWordKorean = "애그리"
        Case "agriculture": CompanyName_EnglishWordKorean = "애그리컬처"
        Case "agro": CompanyName_EnglishWordKorean = "아그로"
        Case "air": CompanyName_EnglishWordKorean = "에어"
        Case "asset", "assets": CompanyName_EnglishWordKorean = "애셋"
        Case "app": CompanyName_EnglishWordKorean = "앱"
        Case "architecture": CompanyName_EnglishWordKorean = "아키텍처"
        Case "art", "arts": CompanyName_EnglishWordKorean = "아트"
        Case "auto": CompanyName_EnglishWordKorean = "오토"
        Case "bank": CompanyName_EnglishWordKorean = "뱅크"
        Case "bakery": CompanyName_EnglishWordKorean = "베이커리"
        Case "beauty": CompanyName_EnglishWordKorean = "뷰티"
        Case "bio": CompanyName_EnglishWordKorean = "바이오"
        Case "block", "blocks": CompanyName_EnglishWordKorean = "블록"
        Case "blue": CompanyName_EnglishWordKorean = "블루"
        Case "book": CompanyName_EnglishWordKorean = "북"
        Case "books": CompanyName_EnglishWordKorean = "북스"
        Case "bridge": CompanyName_EnglishWordKorean = "브릿지"
        Case "build": CompanyName_EnglishWordKorean = "빌드"
        Case "building": CompanyName_EnglishWordKorean = "빌딩"
        Case "business": CompanyName_EnglishWordKorean = "비즈니스"
        Case "cafe": CompanyName_EnglishWordKorean = "카페"
        Case "capital": CompanyName_EnglishWordKorean = "캐피탈"
        Case "cargo": CompanyName_EnglishWordKorean = "카고"
        Case "care": CompanyName_EnglishWordKorean = "케어"
        Case "center", "centre": CompanyName_EnglishWordKorean = "센터"
        Case "chemical", "chemicals": CompanyName_EnglishWordKorean = "케미컬"
        Case "city": CompanyName_EnglishWordKorean = "시티"
        Case "clinic": CompanyName_EnglishWordKorean = "클리닉"
        Case "cloud": CompanyName_EnglishWordKorean = "클라우드"
        Case "code": CompanyName_EnglishWordKorean = "코드"
        Case "coffee": CompanyName_EnglishWordKorean = "커피"
        Case "commerce": CompanyName_EnglishWordKorean = "커머스"
        Case "communication", "communications": CompanyName_EnglishWordKorean = "커뮤니케이션"
        Case "company": CompanyName_EnglishWordKorean = "컴퍼니"
        Case "construction": CompanyName_EnglishWordKorean = "컨스트럭션"
        Case "consulting": CompanyName_EnglishWordKorean = "컨설팅"
        Case "contents", "content": CompanyName_EnglishWordKorean = "콘텐츠"
        Case "creative": CompanyName_EnglishWordKorean = "크리에이티브"
        Case "credit": CompanyName_EnglishWordKorean = "크레딧"
        Case "culture": CompanyName_EnglishWordKorean = "컬처"
        Case "data": CompanyName_EnglishWordKorean = "데이터"
        Case "dental": CompanyName_EnglishWordKorean = "덴탈"
        Case "department": CompanyName_EnglishWordKorean = "디파트먼트"
        Case "delivery": CompanyName_EnglishWordKorean = "딜리버리"
        Case "develop", "developer", "developers": CompanyName_EnglishWordKorean = "디벨로퍼"
        Case "development": CompanyName_EnglishWordKorean = "디벨롭먼트"
        Case "dev": CompanyName_EnglishWordKorean = "데브"
        Case "design": CompanyName_EnglishWordKorean = "디자인"
        Case "digital": CompanyName_EnglishWordKorean = "디지털"
        Case "distribution": CompanyName_EnglishWordKorean = "디스트리뷰션"
        Case "edu", "education": CompanyName_EnglishWordKorean = "에듀"
        Case "electric": CompanyName_EnglishWordKorean = "일렉트릭"
        Case "electronics": CompanyName_EnglishWordKorean = "일렉트로닉스"
        Case "energy": CompanyName_EnglishWordKorean = "에너지"
        Case "engineering": CompanyName_EnglishWordKorean = "엔지니어링"
        Case "entertainment": CompanyName_EnglishWordKorean = "엔터테인먼트"
        Case "environment": CompanyName_EnglishWordKorean = "인바이런먼트"
        Case "estate": CompanyName_EnglishWordKorean = "에스테이트"
        Case "event", "events": CompanyName_EnglishWordKorean = "이벤트"
        Case "express": CompanyName_EnglishWordKorean = "익스프레스"
        Case "factory": CompanyName_EnglishWordKorean = "팩토리"
        Case "farm", "farms": CompanyName_EnglishWordKorean = "팜"
        Case "fashion": CompanyName_EnglishWordKorean = "패션"
        Case "finance", "financial": CompanyName_EnglishWordKorean = "파이낸스"
        Case "food": CompanyName_EnglishWordKorean = "푸드"
        Case "forward": CompanyName_EnglishWordKorean = "포워드"
        Case "fresh": CompanyName_EnglishWordKorean = "프레시"
        Case "furniture": CompanyName_EnglishWordKorean = "퍼니처"
        Case "game", "games", "gaming": CompanyName_EnglishWordKorean = "게임"
        Case "global": CompanyName_EnglishWordKorean = "글로벌"
        Case "green": CompanyName_EnglishWordKorean = "그린"
        Case "group": CompanyName_EnglishWordKorean = "그룹"
        Case "health", "healthcare": CompanyName_EnglishWordKorean = "헬스케어"
        Case "home": CompanyName_EnglishWordKorean = "홈"
        Case "holdings", "holding": CompanyName_EnglishWordKorean = "홀딩스"
        Case "hospital": CompanyName_EnglishWordKorean = "호스피탈"
        Case "house": CompanyName_EnglishWordKorean = "하우스"
        Case "hotel": CompanyName_EnglishWordKorean = "호텔"
        Case "inc": CompanyName_EnglishWordKorean = "인크"
        Case "insurance": CompanyName_EnglishWordKorean = "인슈어런스"
        Case "internet": CompanyName_EnglishWordKorean = "인터넷"
        Case "industry", "industries": CompanyName_EnglishWordKorean = "인더스트리"
        Case "innovation": CompanyName_EnglishWordKorean = "이노베이션"
        Case "international": CompanyName_EnglishWordKorean = "인터내셔널"
        Case "investment": CompanyName_EnglishWordKorean = "인베스트먼트"
        Case "it": CompanyName_EnglishWordKorean = "아이티"
        Case "kakao": CompanyName_EnglishWordKorean = "카카오"
        Case "korea": CompanyName_EnglishWordKorean = "코리아"
        Case "lab", "labs": CompanyName_EnglishWordKorean = "랩"
        Case "law": CompanyName_EnglishWordKorean = "로"
        Case "legal": CompanyName_EnglishWordKorean = "리걸"
        Case "leisure": CompanyName_EnglishWordKorean = "레저"
        Case "life": CompanyName_EnglishWordKorean = "라이프"
        Case "logistics": CompanyName_EnglishWordKorean = "로지스틱스"
        Case "machine", "machines": CompanyName_EnglishWordKorean = "머신"
        Case "machinery": CompanyName_EnglishWordKorean = "머시너리"
        Case "mall": CompanyName_EnglishWordKorean = "몰"
        Case "management": CompanyName_EnglishWordKorean = "매니지먼트"
        Case "manufacturing": CompanyName_EnglishWordKorean = "매뉴팩처링"
        Case "marine": CompanyName_EnglishWordKorean = "마린"
        Case "marketing": CompanyName_EnglishWordKorean = "마케팅"
        Case "market": CompanyName_EnglishWordKorean = "마켓"
        Case "materials": CompanyName_EnglishWordKorean = "머티리얼즈"
        Case "media": CompanyName_EnglishWordKorean = "미디어"
        Case "medical": CompanyName_EnglishWordKorean = "메디컬"
        Case "metal", "metals": CompanyName_EnglishWordKorean = "메탈"
        Case "mobility": CompanyName_EnglishWordKorean = "모빌리티"
        Case "mobile": CompanyName_EnglishWordKorean = "모바일"
        Case "motor", "motors": CompanyName_EnglishWordKorean = "모터"
        Case "movie", "movies": CompanyName_EnglishWordKorean = "무비"
        Case "music": CompanyName_EnglishWordKorean = "뮤직"
        Case "network", "networks": CompanyName_EnglishWordKorean = "네트워크"
        Case "office": CompanyName_EnglishWordKorean = "오피스"
        Case "online": CompanyName_EnglishWordKorean = "온라인"
        Case "partners": CompanyName_EnglishWordKorean = "파트너스"
        Case "pharma": CompanyName_EnglishWordKorean = "파마"
        Case "pharmaceutical", "pharmaceuticals": CompanyName_EnglishWordKorean = "파마슈티컬"
        Case "plastic", "plastics": CompanyName_EnglishWordKorean = "플라스틱"
        Case "platform": CompanyName_EnglishWordKorean = "플랫폼"
        Case "plus": CompanyName_EnglishWordKorean = "플러스"
        Case "port": CompanyName_EnglishWordKorean = "포트"
        Case "power": CompanyName_EnglishWordKorean = "파워"
        Case "print", "printing": CompanyName_EnglishWordKorean = "프린팅"
        Case "production": CompanyName_EnglishWordKorean = "프로덕션"
        Case "project": CompanyName_EnglishWordKorean = "프로젝트"
        Case "property": CompanyName_EnglishWordKorean = "프로퍼티"
        Case "realty": CompanyName_EnglishWordKorean = "리얼티"
        Case "research": CompanyName_EnglishWordKorean = "리서치"
        Case "resort", "resorts": CompanyName_EnglishWordKorean = "리조트"
        Case "restaurant", "restaurants": CompanyName_EnglishWordKorean = "레스토랑"
        Case "royal": CompanyName_EnglishWordKorean = "로얄"
        Case "school": CompanyName_EnglishWordKorean = "스쿨"
        Case "science": CompanyName_EnglishWordKorean = "사이언스"
        Case "security": CompanyName_EnglishWordKorean = "시큐리티"
        Case "securities": CompanyName_EnglishWordKorean = "시큐리티즈"
        Case "server": CompanyName_EnglishWordKorean = "서버"
        Case "service", "services": CompanyName_EnglishWordKorean = "서비스"
        Case "shipping": CompanyName_EnglishWordKorean = "쉬핑"
        Case "shop": CompanyName_EnglishWordKorean = "샵"
        Case "smart": CompanyName_EnglishWordKorean = "스마트"
        Case "soft": CompanyName_EnglishWordKorean = "소프트"
        Case "software": CompanyName_EnglishWordKorean = "소프트웨어"
        Case "solution", "solutions": CompanyName_EnglishWordKorean = "솔루션"
        Case "space": CompanyName_EnglishWordKorean = "스페이스"
        Case "sports": CompanyName_EnglishWordKorean = "스포츠"
        Case "steel": CompanyName_EnglishWordKorean = "스틸"
        Case "store", "stores": CompanyName_EnglishWordKorean = "스토어"
        Case "studio": CompanyName_EnglishWordKorean = "스튜디오"
        Case "system", "systems": CompanyName_EnglishWordKorean = "시스템"
        Case "tax": CompanyName_EnglishWordKorean = "택스"
        Case "tech": CompanyName_EnglishWordKorean = "테크"
        Case "technology", "technologies": CompanyName_EnglishWordKorean = "테크놀로지"
        Case "telecom": CompanyName_EnglishWordKorean = "텔레콤"
        Case "textile", "textiles": CompanyName_EnglishWordKorean = "텍스타일"
        Case "tour": CompanyName_EnglishWordKorean = "투어"
        Case "trade", "trading": CompanyName_EnglishWordKorean = "트레이딩"
        Case "transport": CompanyName_EnglishWordKorean = "트랜스포트"
        Case "travel": CompanyName_EnglishWordKorean = "트래블"
        Case "tva": CompanyName_EnglishWordKorean = "티바"
        Case "tvasoft": CompanyName_EnglishWordKorean = "티바소프트"
        Case "venture", "ventures": CompanyName_EnglishWordKorean = "벤처스"
        Case "vision": CompanyName_EnglishWordKorean = "비전"
        Case "web": CompanyName_EnglishWordKorean = "웹"
        Case "works": CompanyName_EnglishWordKorean = "웍스"
        Case "world": CompanyName_EnglishWordKorean = "월드"
        Case Else: CompanyName_EnglishWordKorean = CompanyName_KRDictLoanwordKorean(wordText)
    End Select
End Function

' BEGIN KRDICT LOANWORD DICTIONARY
' Source: National Institute of Korean Language Korean Basic Dictionary XML.
' Generated by scripts\update_krdict_loanword_dictionary.ps1.
Private Function CompanyName_KRDictLoanwordKorean(ByVal wordText As String) As String
    Select Case Left$(LCase$(wordText), 1)
        Case "a": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_A(wordText)
        Case "b": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_B(wordText)
        Case "c": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_C(wordText)
        Case "d": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_D(wordText)
        Case "e": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_E(wordText)
        Case "f": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_F(wordText)
        Case "g": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_G(wordText)
        Case "h": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_H(wordText)
        Case "i": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_I(wordText)
        Case "j": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_J(wordText)
        Case "k": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_K(wordText)
        Case "l": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_L(wordText)
        Case "m": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_M(wordText)
        Case "n": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_N(wordText)
        Case "o": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_O(wordText)
        Case "p": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_P(wordText)
        Case "q": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_Q(wordText)
        Case "r": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_R(wordText)
        Case "s": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_S(wordText)
        Case "t": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_T(wordText)
        Case "u": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_U(wordText)
        Case "v": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_V(wordText)
        Case "w": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_W(wordText)
        Case "x": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_X(wordText)
        Case "y": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_Y(wordText)
        Case "z": CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_Z(wordText)
        Case Else: CompanyName_KRDictLoanwordKorean = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_A(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "acacia": CompanyName_KRDictLoanwordKorean_A = "아카시아"
        Case "academy": CompanyName_KRDictLoanwordKorean_A = "아카데미"
        Case "accent": CompanyName_KRDictLoanwordKorean_A = "악센트"
        Case "accordion": CompanyName_KRDictLoanwordKorean_A = "아코디언"
        Case "ace": CompanyName_KRDictLoanwordKorean_A = "에이스"
        Case "acetone": CompanyName_KRDictLoanwordKorean_A = "아세톤"
        Case "acre": CompanyName_KRDictLoanwordKorean_A = "에이커"
        Case "action": CompanyName_KRDictLoanwordKorean_A = "액션"
        Case "aerobic": CompanyName_KRDictLoanwordKorean_A = "에어로빅"
        Case "afghanistan": CompanyName_KRDictLoanwordKorean_A = "아프가니스탄"
        Case "africa": CompanyName_KRDictLoanwordKorean_A = "아프리카"
        Case "afterservice": CompanyName_KRDictLoanwordKorean_A = "애프터서비스"
        Case "agent": CompanyName_KRDictLoanwordKorean_A = "에이전트"
        Case "aids": CompanyName_KRDictLoanwordKorean_A = "에이즈"
        Case "airbag": CompanyName_KRDictLoanwordKorean_A = "에어백"
        Case "airconditioner": CompanyName_KRDictLoanwordKorean_A = "에어컨디셔너"
        Case "airfryer": CompanyName_KRDictLoanwordKorean_A = "에어프라이어"
        Case "alarm": CompanyName_KRDictLoanwordKorean_A = "알람"
        Case "albania": CompanyName_KRDictLoanwordKorean_A = "알바니아"
        Case "album": CompanyName_KRDictLoanwordKorean_A = "앨범"
        Case "alcohol": CompanyName_KRDictLoanwordKorean_A = "알코올"
        Case "algorithm": CompanyName_KRDictLoanwordKorean_A = "알고리듬"
        Case "alibi": CompanyName_KRDictLoanwordKorean_A = "알리바이"
        Case "alkali": CompanyName_KRDictLoanwordKorean_A = "알칼리"
        Case "allah": CompanyName_KRDictLoanwordKorean_A = "알라"
        Case "almond": CompanyName_KRDictLoanwordKorean_A = "아몬드"
        Case "aloe": CompanyName_KRDictLoanwordKorean_A = "알로에"
        Case "alphabet": CompanyName_KRDictLoanwordKorean_A = "알파벳"
        Case "alto": CompanyName_KRDictLoanwordKorean_A = "알토"
        Case "aluminium": CompanyName_KRDictLoanwordKorean_A = "알루미늄"
        Case "amateur": CompanyName_KRDictLoanwordKorean_A = "아마추어"
        Case "ambulance": CompanyName_KRDictLoanwordKorean_A = "앰뷸런스"
        Case "amen": CompanyName_KRDictLoanwordKorean_A = "아멘"
        Case "america": CompanyName_KRDictLoanwordKorean_A = "아메리카"
        Case "americano": CompanyName_KRDictLoanwordKorean_A = "아메리카노"
        Case "ammonia": CompanyName_KRDictLoanwordKorean_A = "암모니아"
        Case "amoeba": CompanyName_KRDictLoanwordKorean_A = "아메바"
        Case "analogue": CompanyName_KRDictLoanwordKorean_A = "아날로그"
        Case "anchor": CompanyName_KRDictLoanwordKorean_A = "앵커"
        Case "andante": CompanyName_KRDictLoanwordKorean_A = "안단테"
        Case "andantino": CompanyName_KRDictLoanwordKorean_A = "안단티노"
        Case "angola": CompanyName_KRDictLoanwordKorean_A = "앙골라"
        Case "angora": CompanyName_KRDictLoanwordKorean_A = "앙고라"
        Case "animation": CompanyName_KRDictLoanwordKorean_A = "애니메이션"
        Case "announcer": CompanyName_KRDictLoanwordKorean_A = "아나운서"
        Case "antenna": CompanyName_KRDictLoanwordKorean_A = "안테나"
        Case "app": CompanyName_KRDictLoanwordKorean_A = "앱"
        Case "appeal": CompanyName_KRDictLoanwordKorean_A = "어필"
        Case "application": CompanyName_KRDictLoanwordKorean_A = "애플리케이션"
        Case "aquarium": CompanyName_KRDictLoanwordKorean_A = "아쿠아리움"
        Case "arab": CompanyName_KRDictLoanwordKorean_A = "아랍"
        Case "arabemirates": CompanyName_KRDictLoanwordKorean_A = "아랍에미리트"
        Case "arch": CompanyName_KRDictLoanwordKorean_A = "아치"
        Case "argentina": CompanyName_KRDictLoanwordKorean_A = "아르헨티나"
        Case "aria": CompanyName_KRDictLoanwordKorean_A = "아리아"
        Case "aroma": CompanyName_KRDictLoanwordKorean_A = "아로마"
        Case "art": CompanyName_KRDictLoanwordKorean_A = "아트"
        Case "arthall": CompanyName_KRDictLoanwordKorean_A = "아트홀"
        Case "artist": CompanyName_KRDictLoanwordKorean_A = "아티스트"
        Case "asia": CompanyName_KRDictLoanwordKorean_A = "아시아"
        Case "asphalt": CompanyName_KRDictLoanwordKorean_A = "아스팔트"
        Case "aspirin": CompanyName_KRDictLoanwordKorean_A = "아스피린"
        Case "audio": CompanyName_KRDictLoanwordKorean_A = "오디오"
        Case "audition": CompanyName_KRDictLoanwordKorean_A = "오디션"
        Case "aurora": CompanyName_KRDictLoanwordKorean_A = "오로라"
        Case "australia": CompanyName_KRDictLoanwordKorean_A = "오스트레일리아"
        Case "austria": CompanyName_KRDictLoanwordKorean_A = "오스트리아"
        Case "avocado": CompanyName_KRDictLoanwordKorean_A = "아보카도"
        Case Else: CompanyName_KRDictLoanwordKorean_A = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_B(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "babysitter": CompanyName_KRDictLoanwordKorean_B = "베이비시터"
        Case "backpack": CompanyName_KRDictLoanwordKorean_B = "백팩"
        Case "backup": CompanyName_KRDictLoanwordKorean_B = "백업"
        Case "bacon": CompanyName_KRDictLoanwordKorean_B = "베이컨"
        Case "bacteria": CompanyName_KRDictLoanwordKorean_B = "박테리아"
        Case "badge": CompanyName_KRDictLoanwordKorean_B = "배지"
        Case "badminton": CompanyName_KRDictLoanwordKorean_B = "배드민턴"
        Case "bag": CompanyName_KRDictLoanwordKorean_B = "백"
        Case "bagel": CompanyName_KRDictLoanwordKorean_B = "베이글"
        Case "baguette": CompanyName_KRDictLoanwordKorean_B = "바게트"
        Case "bahrain": CompanyName_KRDictLoanwordKorean_B = "바레인"
        Case "bakery": CompanyName_KRDictLoanwordKorean_B = "베이커리"
        Case "bakingpowder": CompanyName_KRDictLoanwordKorean_B = "베이킹파우더"
        Case "balcony": CompanyName_KRDictLoanwordKorean_B = "발코니"
        Case "ball": CompanyName_KRDictLoanwordKorean_B = "볼"
        Case "ballerina": CompanyName_KRDictLoanwordKorean_B = "발레리나"
        Case "ballet": CompanyName_KRDictLoanwordKorean_B = "발레"
        Case "banana": CompanyName_KRDictLoanwordKorean_B = "바나나"
        Case "band": CompanyName_KRDictLoanwordKorean_B = "밴드"
        Case "bangkok": CompanyName_KRDictLoanwordKorean_B = "방콕"
        Case "bangladesh": CompanyName_KRDictLoanwordKorean_B = "방글라데시"
        Case "banking": CompanyName_KRDictLoanwordKorean_B = "뱅킹"
        Case "bar": CompanyName_KRDictLoanwordKorean_B = "바"
        Case "barbell": CompanyName_KRDictLoanwordKorean_B = "바벨"
        Case "barcelona": CompanyName_KRDictLoanwordKorean_B = "바르셀로나"
        Case "barcode": CompanyName_KRDictLoanwordKorean_B = "바코드"
        Case "bargainsale": CompanyName_KRDictLoanwordKorean_B = "바겐세일"
        Case "barista": CompanyName_KRDictLoanwordKorean_B = "바리스타"
        Case "baritone": CompanyName_KRDictLoanwordKorean_B = "바리톤"
        Case "barometer": CompanyName_KRDictLoanwordKorean_B = "바로미터"
        Case "baroque": CompanyName_KRDictLoanwordKorean_B = "바로크"
        Case "barrel": CompanyName_KRDictLoanwordKorean_B = "배럴"
        Case "barricade": CompanyName_KRDictLoanwordKorean_B = "바리케이드"
        Case "bartender": CompanyName_KRDictLoanwordKorean_B = "바텐더"
        Case "base": CompanyName_KRDictLoanwordKorean_B = "베이스"
        Case "basket": CompanyName_KRDictLoanwordKorean_B = "바스켓"
        Case "bass": CompanyName_KRDictLoanwordKorean_B = "베이스"
        Case "bassoon": CompanyName_KRDictLoanwordKorean_B = "바순"
        Case "bat": CompanyName_KRDictLoanwordKorean_B = "배트"
        Case "baton": CompanyName_KRDictLoanwordKorean_B = "배턴"
        Case "battery": CompanyName_KRDictLoanwordKorean_B = "배터리"
        Case "bboy": CompanyName_KRDictLoanwordKorean_B = "비보이"
        Case "beaker": CompanyName_KRDictLoanwordKorean_B = "비커"
        Case "beefsteak": CompanyName_KRDictLoanwordKorean_B = "비프스테이크"
        Case "beige": CompanyName_KRDictLoanwordKorean_B = "베이지"
        Case "beijing": CompanyName_KRDictLoanwordKorean_B = "베이징"
        Case "bell": CompanyName_KRDictLoanwordKorean_B = "벨"
        Case "belt": CompanyName_KRDictLoanwordKorean_B = "벨트"
        Case "bench": CompanyName_KRDictLoanwordKorean_B = "벤치"
        Case "benchmarking": CompanyName_KRDictLoanwordKorean_B = "벤치마킹"
        Case "bestseller": CompanyName_KRDictLoanwordKorean_B = "베스트셀러"
        Case "bhutan": CompanyName_KRDictLoanwordKorean_B = "부탄"
        Case "bible": CompanyName_KRDictLoanwordKorean_B = "바이블"
        Case "bidet": CompanyName_KRDictLoanwordKorean_B = "비데"
        Case "biennale": CompanyName_KRDictLoanwordKorean_B = "비엔날레"
        Case "big": CompanyName_KRDictLoanwordKorean_B = "빅"
        Case "bigdata": CompanyName_KRDictLoanwordKorean_B = "빅데이터"
        Case "bikini": CompanyName_KRDictLoanwordKorean_B = "비키니"
        Case "binder": CompanyName_KRDictLoanwordKorean_B = "바인더"
        Case "biscuit": CompanyName_KRDictLoanwordKorean_B = "비스킷"
        Case "blackbox": CompanyName_KRDictLoanwordKorean_B = "블랙박스"
        Case "blackhole": CompanyName_KRDictLoanwordKorean_B = "블랙홀"
        Case "blacklist": CompanyName_KRDictLoanwordKorean_B = "블랙리스트"
        Case "blind": CompanyName_KRDictLoanwordKorean_B = "블라인드"
        Case "blockbuster": CompanyName_KRDictLoanwordKorean_B = "블록버스터"
        Case "blog": CompanyName_KRDictLoanwordKorean_B = "블로그"
        Case "blogger": CompanyName_KRDictLoanwordKorean_B = "블로거"
        Case "blouse": CompanyName_KRDictLoanwordKorean_B = "블라우스"
        Case "blueberry": CompanyName_KRDictLoanwordKorean_B = "블루베리"
        Case "bluecollar": CompanyName_KRDictLoanwordKorean_B = "블루칼라"
        Case "bluetooth": CompanyName_KRDictLoanwordKorean_B = "블루투스"
        Case "boat": CompanyName_KRDictLoanwordKorean_B = "보트"
        Case "body": CompanyName_KRDictLoanwordKorean_B = "보디"
        Case "bodybuilding": CompanyName_KRDictLoanwordKorean_B = "보디빌딩"
        Case "bodywash": CompanyName_KRDictLoanwordKorean_B = "보디워시"
        Case "boiler": CompanyName_KRDictLoanwordKorean_B = "보일러"
        Case "bolero": CompanyName_KRDictLoanwordKorean_B = "볼레로"
        Case "bolivia": CompanyName_KRDictLoanwordKorean_B = "볼리비아"
        Case "bolt": CompanyName_KRDictLoanwordKorean_B = "볼트"
        Case "bonus": CompanyName_KRDictLoanwordKorean_B = "보너스"
        Case "boom": CompanyName_KRDictLoanwordKorean_B = "붐"
        Case "boomerang": CompanyName_KRDictLoanwordKorean_B = "부메랑"
        Case "booth": CompanyName_KRDictLoanwordKorean_B = "부스"
        Case "booting": CompanyName_KRDictLoanwordKorean_B = "부팅"
        Case "boss": CompanyName_KRDictLoanwordKorean_B = "보스"
        Case "bouquet": CompanyName_KRDictLoanwordKorean_B = "부케"
        Case "bourgeois": CompanyName_KRDictLoanwordKorean_B = "부르주아"
        Case "bowling": CompanyName_KRDictLoanwordKorean_B = "볼링"
        Case "box": CompanyName_KRDictLoanwordKorean_B = "박스"
        Case "boxing": CompanyName_KRDictLoanwordKorean_B = "복싱"
        Case "boy": CompanyName_KRDictLoanwordKorean_B = "보이"
        Case "boycott": CompanyName_KRDictLoanwordKorean_B = "보이콧"
        Case "brake": CompanyName_KRDictLoanwordKorean_B = "브레이크"
        Case "brand": CompanyName_KRDictLoanwordKorean_B = "브랜드"
        Case "bravo": CompanyName_KRDictLoanwordKorean_B = "브라보"
        Case "brazil": CompanyName_KRDictLoanwordKorean_B = "브라질"
        Case "briefing": CompanyName_KRDictLoanwordKorean_B = "브리핑"
        Case "broccoli": CompanyName_KRDictLoanwordKorean_B = "브로콜리"
        Case "broker": CompanyName_KRDictLoanwordKorean_B = "브로커"
        Case "brooch": CompanyName_KRDictLoanwordKorean_B = "브로치"
        Case "browser": CompanyName_KRDictLoanwordKorean_B = "브라우저"
        Case "brunch": CompanyName_KRDictLoanwordKorean_B = "브런치"
        Case "brush": CompanyName_KRDictLoanwordKorean_B = "브러시"
        Case "buffet": CompanyName_KRDictLoanwordKorean_B = "뷔페"
        Case "bug": CompanyName_KRDictLoanwordKorean_B = "버그"
        Case "building": CompanyName_KRDictLoanwordKorean_B = "빌딩"
        Case "builtin": CompanyName_KRDictLoanwordKorean_B = "빌트인"
        Case "bulgaria": CompanyName_KRDictLoanwordKorean_B = "불가리아"
        Case "bulldog": CompanyName_KRDictLoanwordKorean_B = "불도그"
        Case "bulldozer": CompanyName_KRDictLoanwordKorean_B = "불도저"
        Case "bumper": CompanyName_KRDictLoanwordKorean_B = "범퍼"
        Case "bungalow": CompanyName_KRDictLoanwordKorean_B = "방갈로"
        Case "bungee": CompanyName_KRDictLoanwordKorean_B = "번지"
        Case "bungeejump": CompanyName_KRDictLoanwordKorean_B = "번지점프"
        Case "bunker": CompanyName_KRDictLoanwordKorean_B = "벙커"
        Case "bunt": CompanyName_KRDictLoanwordKorean_B = "번트"
        Case "burger": CompanyName_KRDictLoanwordKorean_B = "버거"
        Case "burner": CompanyName_KRDictLoanwordKorean_B = "버너"
        Case "bus": CompanyName_KRDictLoanwordKorean_B = "버스"
        Case "business": CompanyName_KRDictLoanwordKorean_B = "비즈니스"
        Case "butanegas": CompanyName_KRDictLoanwordKorean_B = "부탄가스"
        Case "butter": CompanyName_KRDictLoanwordKorean_B = "버터"
        Case "button": CompanyName_KRDictLoanwordKorean_B = "버튼"
        Case "buyer": CompanyName_KRDictLoanwordKorean_B = "바이어"
        Case "buzzer": CompanyName_KRDictLoanwordKorean_B = "버저"
        Case "byte": CompanyName_KRDictLoanwordKorean_B = "바이트"
        Case Else: CompanyName_KRDictLoanwordKorean_B = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_C(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "cabaret": CompanyName_KRDictLoanwordKorean_C = "카바레"
        Case "cabinet": CompanyName_KRDictLoanwordKorean_C = "캐비닛"
        Case "cable": CompanyName_KRDictLoanwordKorean_C = "케이블"
        Case "cablecar": CompanyName_KRDictLoanwordKorean_C = "케이블카"
        Case "cafeteria": CompanyName_KRDictLoanwordKorean_C = "카페테리아"
        Case "caffeine": CompanyName_KRDictLoanwordKorean_C = "카페인"
        Case "cake": CompanyName_KRDictLoanwordKorean_C = "케이크"
        Case "calcium": CompanyName_KRDictLoanwordKorean_C = "칼슘"
        Case "callcenter": CompanyName_KRDictLoanwordKorean_C = "콜센터"
        Case "callgirl": CompanyName_KRDictLoanwordKorean_C = "콜걸"
        Case "calltaxi": CompanyName_KRDictLoanwordKorean_C = "콜택시"
        Case "calorie": CompanyName_KRDictLoanwordKorean_C = "칼로리"
        Case "cambodia": CompanyName_KRDictLoanwordKorean_C = "캄보디아"
        Case "camera": CompanyName_KRDictLoanwordKorean_C = "카메라"
        Case "cameroon": CompanyName_KRDictLoanwordKorean_C = "카메룬"
        Case "camisole": CompanyName_KRDictLoanwordKorean_C = "캐미솔"
        Case "camp": CompanyName_KRDictLoanwordKorean_C = "캠프"
        Case "campaign": CompanyName_KRDictLoanwordKorean_C = "캠페인"
        Case "campfire": CompanyName_KRDictLoanwordKorean_C = "캠프파이어"
        Case "camping": CompanyName_KRDictLoanwordKorean_C = "캠핑"
        Case "campus": CompanyName_KRDictLoanwordKorean_C = "캠퍼스"
        Case "can": CompanyName_KRDictLoanwordKorean_C = "캔"
        Case "canada": CompanyName_KRDictLoanwordKorean_C = "캐나다"
        Case "candy": CompanyName_KRDictLoanwordKorean_C = "캔디"
        Case "canoe": CompanyName_KRDictLoanwordKorean_C = "카누"
        Case "canvas": CompanyName_KRDictLoanwordKorean_C = "캔버스"
        Case "cap": CompanyName_KRDictLoanwordKorean_C = "캡"
        Case "cappuccino": CompanyName_KRDictLoanwordKorean_C = "카푸치노"
        Case "capsule": CompanyName_KRDictLoanwordKorean_C = "캡슐"
        Case "capture": CompanyName_KRDictLoanwordKorean_C = "캡처"
        Case "caramel": CompanyName_KRDictLoanwordKorean_C = "캐러멜"
        Case "carat": CompanyName_KRDictLoanwordKorean_C = "캐럿"
        Case "card": CompanyName_KRDictLoanwordKorean_C = "카드"
        Case "cardigan": CompanyName_KRDictLoanwordKorean_C = "카디건"
        Case "carnation": CompanyName_KRDictLoanwordKorean_C = "카네이션"
        Case "carol": CompanyName_KRDictLoanwordKorean_C = "캐럴"
        Case "carpet": CompanyName_KRDictLoanwordKorean_C = "카펫"
        Case "carphone": CompanyName_KRDictLoanwordKorean_C = "카폰"
        Case "carpool": CompanyName_KRDictLoanwordKorean_C = "카풀"
        Case "carracer": CompanyName_KRDictLoanwordKorean_C = "카레이서"
        Case "cart": CompanyName_KRDictLoanwordKorean_C = "카트"
        Case "case": CompanyName_KRDictLoanwordKorean_C = "케이스"
        Case "cashback": CompanyName_KRDictLoanwordKorean_C = "캐시백"
        Case "casino": CompanyName_KRDictLoanwordKorean_C = "카지노"
        Case "cassette": CompanyName_KRDictLoanwordKorean_C = "카세트"
        Case "cassettetape": CompanyName_KRDictLoanwordKorean_C = "카세트테이프"
        Case "castanets": CompanyName_KRDictLoanwordKorean_C = "캐스터네츠"
        Case "caster": CompanyName_KRDictLoanwordKorean_C = "캐스터"
        Case "casting": CompanyName_KRDictLoanwordKorean_C = "캐스팅"
        Case "catalog": CompanyName_KRDictLoanwordKorean_C = "카탈로그"
        Case "catharsis": CompanyName_KRDictLoanwordKorean_C = "카타르시스"
        Case "catholic": CompanyName_KRDictLoanwordKorean_C = "가톨릭"
        Case "cd": CompanyName_KRDictLoanwordKorean_C = "시디"
        Case "cdplayer": CompanyName_KRDictLoanwordKorean_C = "시디플레이어"
        Case "cdrom": CompanyName_KRDictLoanwordKorean_C = "시디롬"
        Case "celery": CompanyName_KRDictLoanwordKorean_C = "셀러리"
        Case "cellist": CompanyName_KRDictLoanwordKorean_C = "첼리스트"
        Case "cello": CompanyName_KRDictLoanwordKorean_C = "첼로"
        Case "cellophane": CompanyName_KRDictLoanwordKorean_C = "셀로판"
        Case "cement": CompanyName_KRDictLoanwordKorean_C = "시멘트"
        Case "cent": CompanyName_KRDictLoanwordKorean_C = "센트"
        Case "center": CompanyName_KRDictLoanwordKorean_C = "센터"
        Case "centimeter": CompanyName_KRDictLoanwordKorean_C = "센티미터"
        Case "ceramics": CompanyName_KRDictLoanwordKorean_C = "세라믹"
        Case "cereal": CompanyName_KRDictLoanwordKorean_C = "시리얼"
        Case "chain": CompanyName_KRDictLoanwordKorean_C = "체인"
        Case "chameleon": CompanyName_KRDictLoanwordKorean_C = "카멜레온"
        Case "champagne": CompanyName_KRDictLoanwordKorean_C = "샴페인"
        Case "champion": CompanyName_KRDictLoanwordKorean_C = "챔피언"
        Case "chandelier": CompanyName_KRDictLoanwordKorean_C = "샹들리에"
        Case "channel": CompanyName_KRDictLoanwordKorean_C = "채널"
        Case "chanson": CompanyName_KRDictLoanwordKorean_C = "샹송"
        Case "chaos": CompanyName_KRDictLoanwordKorean_C = "카오스"
        Case "chapel": CompanyName_KRDictLoanwordKorean_C = "채플"
        Case "character": CompanyName_KRDictLoanwordKorean_C = "캐릭터"
        Case "charisma": CompanyName_KRDictLoanwordKorean_C = "카리스마"
        Case "chart": CompanyName_KRDictLoanwordKorean_C = "차트"
        Case "chauvinism": CompanyName_KRDictLoanwordKorean_C = "쇼비니즘"
        Case "check": CompanyName_KRDictLoanwordKorean_C = "체크"
        Case "checkcard": CompanyName_KRDictLoanwordKorean_C = "체크카드"
        Case "checkin": CompanyName_KRDictLoanwordKorean_C = "체크인"
        Case "checkout": CompanyName_KRDictLoanwordKorean_C = "체크아웃"
        Case "cheerleader": CompanyName_KRDictLoanwordKorean_C = "치어리더"
        Case "cheese": CompanyName_KRDictLoanwordKorean_C = "치즈"
        Case "cheetah": CompanyName_KRDictLoanwordKorean_C = "치타"
        Case "chef": CompanyName_KRDictLoanwordKorean_C = "셰프"
        Case "chess": CompanyName_KRDictLoanwordKorean_C = "체스"
        Case "chicken": CompanyName_KRDictLoanwordKorean_C = "치킨"
        Case "chile": CompanyName_KRDictLoanwordKorean_C = "칠레"
        Case "chimpanzee": CompanyName_KRDictLoanwordKorean_C = "침팬지"
        Case "chip": CompanyName_KRDictLoanwordKorean_C = "칩"
        Case "chlorella": CompanyName_KRDictLoanwordKorean_C = "클로렐라"
        Case "chocolate": CompanyName_KRDictLoanwordKorean_C = "초콜릿"
        Case "cholera": CompanyName_KRDictLoanwordKorean_C = "콜레라"
        Case "cholesterol": CompanyName_KRDictLoanwordKorean_C = "콜레스테롤"
        Case "chorus": CompanyName_KRDictLoanwordKorean_C = "코러스"
        Case "christian": CompanyName_KRDictLoanwordKorean_C = "크리스천"
        Case "christmas": CompanyName_KRDictLoanwordKorean_C = "크리스마스"
        Case "christmascarol": CompanyName_KRDictLoanwordKorean_C = "크리스마스캐럴"
        Case "christmaseve": CompanyName_KRDictLoanwordKorean_C = "크리스마스이브"
        Case "christmastree": CompanyName_KRDictLoanwordKorean_C = "크리스마스트리"
        Case "cigar": CompanyName_KRDictLoanwordKorean_C = "시가"
        Case "cinderella": CompanyName_KRDictLoanwordKorean_C = "신데렐라"
        Case "cinnamon": CompanyName_KRDictLoanwordKorean_C = "시나몬"
        Case "circus": CompanyName_KRDictLoanwordKorean_C = "서커스"
        Case "clarinet": CompanyName_KRDictLoanwordKorean_C = "클라리넷"
        Case "classic": CompanyName_KRDictLoanwordKorean_C = "클래식"
        Case "click": CompanyName_KRDictLoanwordKorean_C = "클릭"
        Case "clip": CompanyName_KRDictLoanwordKorean_C = "클립"
        Case "closeup": CompanyName_KRDictLoanwordKorean_C = "클로즈업"
        Case "cloud": CompanyName_KRDictLoanwordKorean_C = "클라우드"
        Case "clover": CompanyName_KRDictLoanwordKorean_C = "클로버"
        Case "club": CompanyName_KRDictLoanwordKorean_C = "클럽"
        Case "clutch": CompanyName_KRDictLoanwordKorean_C = "클러치"
        Case "coach": CompanyName_KRDictLoanwordKorean_C = "코치"
        Case "coat": CompanyName_KRDictLoanwordKorean_C = "코트"
        Case "coating": CompanyName_KRDictLoanwordKorean_C = "코팅"
        Case "cobra": CompanyName_KRDictLoanwordKorean_C = "코브라"
        Case "cocaine": CompanyName_KRDictLoanwordKorean_C = "코카인"
        Case "cocktail": CompanyName_KRDictLoanwordKorean_C = "칵테일"
        Case "cocoa": CompanyName_KRDictLoanwordKorean_C = "코코아"
        Case "coconut": CompanyName_KRDictLoanwordKorean_C = "코코넛"
        Case "code": CompanyName_KRDictLoanwordKorean_C = "코드"
        Case "coding": CompanyName_KRDictLoanwordKorean_C = "코딩"
        Case "coffee": CompanyName_KRDictLoanwordKorean_C = "커피"
        Case "coffeepot": CompanyName_KRDictLoanwordKorean_C = "커피포트"
        Case "coffeeshop": CompanyName_KRDictLoanwordKorean_C = "커피숍"
        Case "cola": CompanyName_KRDictLoanwordKorean_C = "콜라"
        Case "coldcream": CompanyName_KRDictLoanwordKorean_C = "콜드크림"
        Case "collagen": CompanyName_KRDictLoanwordKorean_C = "콜라겐"
        Case "collar": CompanyName_KRDictLoanwordKorean_C = "칼라"
        Case "collection": CompanyName_KRDictLoanwordKorean_C = "컬렉션"
        Case "colombia": CompanyName_KRDictLoanwordKorean_C = "콜롬비아"
        Case "color": CompanyName_KRDictLoanwordKorean_C = "컬러"
        Case "column": CompanyName_KRDictLoanwordKorean_C = "칼럼"
        Case "columnist": CompanyName_KRDictLoanwordKorean_C = "칼럼니스트"
        Case "comeback": CompanyName_KRDictLoanwordKorean_C = "컴백"
        Case "comedian": CompanyName_KRDictLoanwordKorean_C = "코미디언"
        Case "comedy": CompanyName_KRDictLoanwordKorean_C = "코미디"
        Case "comment": CompanyName_KRDictLoanwordKorean_C = "코멘트"
        Case "commission": CompanyName_KRDictLoanwordKorean_C = "커미션"
        Case "communication": CompanyName_KRDictLoanwordKorean_C = "커뮤니케이션"
        Case "community": CompanyName_KRDictLoanwordKorean_C = "커뮤니티"
        Case "compact": CompanyName_KRDictLoanwordKorean_C = "콤팩트"
        Case "compass": CompanyName_KRDictLoanwordKorean_C = "컴퍼스"
        Case "complex": CompanyName_KRDictLoanwordKorean_C = "콤플렉스"
        Case "computer": CompanyName_KRDictLoanwordKorean_C = "컴퓨터"
        Case "computergame": CompanyName_KRDictLoanwordKorean_C = "컴퓨터게임"
        Case "computervirus": CompanyName_KRDictLoanwordKorean_C = "컴퓨터바이러스"
        Case "concert": CompanyName_KRDictLoanwordKorean_C = "콘서트"
        Case "concours": CompanyName_KRDictLoanwordKorean_C = "콩쿠르"
        Case "concrete": CompanyName_KRDictLoanwordKorean_C = "콘크리트"
        Case "condition": CompanyName_KRDictLoanwordKorean_C = "컨디션"
        Case "condom": CompanyName_KRDictLoanwordKorean_C = "콘돔"
        Case "cone": CompanyName_KRDictLoanwordKorean_C = "콘"
        Case "conference": CompanyName_KRDictLoanwordKorean_C = "콘퍼런스"
        Case "consortium": CompanyName_KRDictLoanwordKorean_C = "컨소시엄"
        Case "consultant": CompanyName_KRDictLoanwordKorean_C = "컨설턴트"
        Case "consulting": CompanyName_KRDictLoanwordKorean_C = "컨설팅"
        Case "contactlens": CompanyName_KRDictLoanwordKorean_C = "콘택트렌즈"
        Case "container": CompanyName_KRDictLoanwordKorean_C = "컨테이너"
        Case "conte": CompanyName_KRDictLoanwordKorean_C = "콩트"
        Case "contest": CompanyName_KRDictLoanwordKorean_C = "콘테스트"
        Case "contrabass": CompanyName_KRDictLoanwordKorean_C = "콘트라베이스"
        Case "control": CompanyName_KRDictLoanwordKorean_C = "컨트롤"
        Case "cookie": CompanyName_KRDictLoanwordKorean_C = "쿠키"
        Case "copy": CompanyName_KRDictLoanwordKorean_C = "카피"
        Case "copywriter": CompanyName_KRDictLoanwordKorean_C = "카피라이터"
        Case "cork": CompanyName_KRDictLoanwordKorean_C = "코르크"
        Case "corner": CompanyName_KRDictLoanwordKorean_C = "코너"
        Case "cosmos": CompanyName_KRDictLoanwordKorean_C = "코스모스"
        Case "counseling": CompanyName_KRDictLoanwordKorean_C = "카운슬링"
        Case "counselor": CompanyName_KRDictLoanwordKorean_C = "카운슬러"
        Case "counter": CompanyName_KRDictLoanwordKorean_C = "카운터"
        Case "couple": CompanyName_KRDictLoanwordKorean_C = "커플"
        Case "coupon": CompanyName_KRDictLoanwordKorean_C = "쿠폰"
        Case "course": CompanyName_KRDictLoanwordKorean_C = "코스"
        Case "court": CompanyName_KRDictLoanwordKorean_C = "코트"
        Case "cover": CompanyName_KRDictLoanwordKorean_C = "커버"
        Case "cowboy": CompanyName_KRDictLoanwordKorean_C = "카우보이"
        Case "cracker": CompanyName_KRDictLoanwordKorean_C = "크래커"
        Case "crane": CompanyName_KRDictLoanwordKorean_C = "크레인"
        Case "crayon": CompanyName_KRDictLoanwordKorean_C = "크레용"
        Case "cream": CompanyName_KRDictLoanwordKorean_C = "크림"
        Case "credit": CompanyName_KRDictLoanwordKorean_C = "크레디트"
        Case "creditcard": CompanyName_KRDictLoanwordKorean_C = "크레디트카드"
        Case "croatia": CompanyName_KRDictLoanwordKorean_C = "크로아티아"
        Case "cruise": CompanyName_KRDictLoanwordKorean_C = "크루즈"
        Case "ct": CompanyName_KRDictLoanwordKorean_C = "시티"
        Case "cuba": CompanyName_KRDictLoanwordKorean_C = "쿠바"
        Case "cup": CompanyName_KRDictLoanwordKorean_C = "컵"
        Case "cupid": CompanyName_KRDictLoanwordKorean_C = "큐피드"
        Case "curator": CompanyName_KRDictLoanwordKorean_C = "큐레이터"
        Case "curriculum": CompanyName_KRDictLoanwordKorean_C = "커리큘럼"
        Case "cursor": CompanyName_KRDictLoanwordKorean_C = "커서"
        Case "curtain": CompanyName_KRDictLoanwordKorean_C = "커튼"
        Case "curve": CompanyName_KRDictLoanwordKorean_C = "커브"
        Case "cushion": CompanyName_KRDictLoanwordKorean_C = "쿠션"
        Case "cut": CompanyName_KRDictLoanwordKorean_C = "컷"
        Case "cyber": CompanyName_KRDictLoanwordKorean_C = "사이버"
        Case "cyberterror": CompanyName_KRDictLoanwordKorean_C = "사이버테러"
        Case "cycle": CompanyName_KRDictLoanwordKorean_C = "사이클"
        Case Else: CompanyName_KRDictLoanwordKorean_C = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_D(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "dam": CompanyName_KRDictLoanwordKorean_D = "댐"
        Case "dance": CompanyName_KRDictLoanwordKorean_D = "댄스"
        Case "dancer": CompanyName_KRDictLoanwordKorean_D = "댄서"
        Case "darkcircle": CompanyName_KRDictLoanwordKorean_D = "다크서클"
        Case "darkhorse": CompanyName_KRDictLoanwordKorean_D = "다크호스"
        Case "dash": CompanyName_KRDictLoanwordKorean_D = "대시"
        Case "data": CompanyName_KRDictLoanwordKorean_D = "데이터"
        Case "database": CompanyName_KRDictLoanwordKorean_D = "데이터베이스"
        Case "date": CompanyName_KRDictLoanwordKorean_D = "데이트"
        Case "dealer": CompanyName_KRDictLoanwordKorean_D = "딜러"
        Case "decibel": CompanyName_KRDictLoanwordKorean_D = "데시벨"
        Case "deciliter": CompanyName_KRDictLoanwordKorean_D = "데시리터"
        Case "demo": CompanyName_KRDictLoanwordKorean_D = "데모"
        Case "denmark": CompanyName_KRDictLoanwordKorean_D = "덴마크"
        Case "design": CompanyName_KRDictLoanwordKorean_D = "디자인"
        Case "designer": CompanyName_KRDictLoanwordKorean_D = "디자이너"
        Case "desk": CompanyName_KRDictLoanwordKorean_D = "데스크"
        Case "desktop": CompanyName_KRDictLoanwordKorean_D = "데스크톱"
        Case "desktopcomputer": CompanyName_KRDictLoanwordKorean_D = "데스크톱컴퓨터"
        Case "dessert": CompanyName_KRDictLoanwordKorean_D = "디저트"
        Case "dial": CompanyName_KRDictLoanwordKorean_D = "다이얼"
        Case "diamond": CompanyName_KRDictLoanwordKorean_D = "다이아몬드"
        Case "diary": CompanyName_KRDictLoanwordKorean_D = "다이어리"
        Case "diet": CompanyName_KRDictLoanwordKorean_D = "다이어트"
        Case "digital": CompanyName_KRDictLoanwordKorean_D = "디지털"
        Case "digitalcamera": CompanyName_KRDictLoanwordKorean_D = "디지털카메라"
        Case "dilemma": CompanyName_KRDictLoanwordKorean_D = "딜레마"
        Case "disco": CompanyName_KRDictLoanwordKorean_D = "디스코"
        Case "disk": CompanyName_KRDictLoanwordKorean_D = "디스크"
        Case "diskdrive": CompanyName_KRDictLoanwordKorean_D = "디스크드라이브"
        Case "diskette": CompanyName_KRDictLoanwordKorean_D = "디스켓"
        Case "display": CompanyName_KRDictLoanwordKorean_D = "디스플레이"
        Case "diving": CompanyName_KRDictLoanwordKorean_D = "다이빙"
        Case "dj": CompanyName_KRDictLoanwordKorean_D = "디제이"
        Case "dna": CompanyName_KRDictLoanwordKorean_D = "디엔에이"
        Case "do": CompanyName_KRDictLoanwordKorean_D = "도"
        Case "documentary": CompanyName_KRDictLoanwordKorean_D = "다큐멘터리"
        Case "dogma": CompanyName_KRDictLoanwordKorean_D = "도그마"
        Case "dollar": CompanyName_KRDictLoanwordKorean_D = "달러"
        Case "domain": CompanyName_KRDictLoanwordKorean_D = "도메인"
        Case "dome": CompanyName_KRDictLoanwordKorean_D = "돔"
        Case "domino": CompanyName_KRDictLoanwordKorean_D = "도미노"
        Case "door": CompanyName_KRDictLoanwordKorean_D = "도어"
        Case "double": CompanyName_KRDictLoanwordKorean_D = "더블"
        Case "doublebass": CompanyName_KRDictLoanwordKorean_D = "더블베이스"
        Case "doughnut": CompanyName_KRDictLoanwordKorean_D = "도넛"
        Case "down": CompanyName_KRDictLoanwordKorean_D = "다운"
        Case "download": CompanyName_KRDictLoanwordKorean_D = "다운로드"
        Case "dracula": CompanyName_KRDictLoanwordKorean_D = "드라큘라"
        Case "drama": CompanyName_KRDictLoanwordKorean_D = "드라마"
        Case "dress": CompanyName_KRDictLoanwordKorean_D = "드레스"
        Case "dressing": CompanyName_KRDictLoanwordKorean_D = "드레싱"
        Case "drill": CompanyName_KRDictLoanwordKorean_D = "드릴"
        Case "drink": CompanyName_KRDictLoanwordKorean_D = "드링크"
        Case "drive": CompanyName_KRDictLoanwordKorean_D = "드라이브"
        Case "driver": CompanyName_KRDictLoanwordKorean_D = "드라이버"
        Case "drone": CompanyName_KRDictLoanwordKorean_D = "드론"
        Case "drum": CompanyName_KRDictLoanwordKorean_D = "드럼"
        Case "dry": CompanyName_KRDictLoanwordKorean_D = "드라이"
        Case "drycleaning": CompanyName_KRDictLoanwordKorean_D = "드라이클리닝"
        Case "dryer": CompanyName_KRDictLoanwordKorean_D = "드라이어"
        Case "dubbing": CompanyName_KRDictLoanwordKorean_D = "더빙"
        Case "duet": CompanyName_KRDictLoanwordKorean_D = "듀엣"
        Case "dumping": CompanyName_KRDictLoanwordKorean_D = "덤핑"
        Case "dumptruck": CompanyName_KRDictLoanwordKorean_D = "덤프트럭"
        Case "dvd": CompanyName_KRDictLoanwordKorean_D = "디브이디"
        Case "dynamite": CompanyName_KRDictLoanwordKorean_D = "다이너마이트"
        Case Else: CompanyName_KRDictLoanwordKorean_D = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_E(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "earphone": CompanyName_KRDictLoanwordKorean_E = "이어폰"
        Case "ecuador": CompanyName_KRDictLoanwordKorean_E = "에콰도르"
        Case "egoism": CompanyName_KRDictLoanwordKorean_E = "에고이즘"
        Case "egypt": CompanyName_KRDictLoanwordKorean_E = "이집트"
        Case "elevator": CompanyName_KRDictLoanwordKorean_E = "엘리베이터"
        Case "elsalvador": CompanyName_KRDictLoanwordKorean_E = "엘살바도르"
        Case "email": CompanyName_KRDictLoanwordKorean_E = "이메일"
        Case "emirates": CompanyName_KRDictLoanwordKorean_E = "에미리트"
        Case "emoticon": CompanyName_KRDictLoanwordKorean_E = "이모티콘"
        Case "enamel": CompanyName_KRDictLoanwordKorean_E = "에나멜"
        Case "encore": CompanyName_KRDictLoanwordKorean_E = "앙코르"
        Case "endorphin": CompanyName_KRDictLoanwordKorean_E = "엔도르핀"
        Case "energy": CompanyName_KRDictLoanwordKorean_E = "에너지"
        Case "engine": CompanyName_KRDictLoanwordKorean_E = "엔진"
        Case "engineer": CompanyName_KRDictLoanwordKorean_E = "엔지니어"
        Case "engineering": CompanyName_KRDictLoanwordKorean_E = "엔지니어링"
        Case "england": CompanyName_KRDictLoanwordKorean_E = "잉글랜드"
        Case "ensemble": CompanyName_KRDictLoanwordKorean_E = "앙상블"
        Case "entertainment": CompanyName_KRDictLoanwordKorean_E = "엔터테인먼트"
        Case "entry": CompanyName_KRDictLoanwordKorean_E = "엔트리"
        Case "episode": CompanyName_KRDictLoanwordKorean_E = "에피소드"
        Case "eros": CompanyName_KRDictLoanwordKorean_E = "에로스"
        Case "eroticism": CompanyName_KRDictLoanwordKorean_E = "에로티시즘"
        Case "error": CompanyName_KRDictLoanwordKorean_E = "에러"
        Case "escalator": CompanyName_KRDictLoanwordKorean_E = "에스컬레이터"
        Case "escort": CompanyName_KRDictLoanwordKorean_E = "에스코트"
        Case "espresso": CompanyName_KRDictLoanwordKorean_E = "에스프레소"
        Case "essay": CompanyName_KRDictLoanwordKorean_E = "에세이"
        Case "ethiopia": CompanyName_KRDictLoanwordKorean_E = "에티오피아"
        Case "eurasia": CompanyName_KRDictLoanwordKorean_E = "유라시아"
        Case "euro": CompanyName_KRDictLoanwordKorean_E = "유로"
        Case "europe": CompanyName_KRDictLoanwordKorean_E = "유럽"
        Case "eve": CompanyName_KRDictLoanwordKorean_E = "이브"
        Case "event": CompanyName_KRDictLoanwordKorean_E = "이벤트"
        Case "expo": CompanyName_KRDictLoanwordKorean_E = "엑스포"
        Case "extra": CompanyName_KRDictLoanwordKorean_E = "엑스트라"
        Case "eyeliner": CompanyName_KRDictLoanwordKorean_E = "아이라이너"
        Case "eyeshadow": CompanyName_KRDictLoanwordKorean_E = "아이섀도"
        Case Else: CompanyName_KRDictLoanwordKorean_E = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_F(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "fa": CompanyName_KRDictLoanwordKorean_F = "파"
        Case "facsimile": CompanyName_KRDictLoanwordKorean_F = "팩시밀리"
        Case "fan": CompanyName_KRDictLoanwordKorean_F = "팬"
        Case "fandom": CompanyName_KRDictLoanwordKorean_F = "팬덤"
        Case "fanfare": CompanyName_KRDictLoanwordKorean_F = "팡파르"
        Case "fantasy": CompanyName_KRDictLoanwordKorean_F = "판타지"
        Case "fascism": CompanyName_KRDictLoanwordKorean_F = "파시즘"
        Case "fascist": CompanyName_KRDictLoanwordKorean_F = "파시스트"
        Case "fashion": CompanyName_KRDictLoanwordKorean_F = "패션"
        Case "fashionshow": CompanyName_KRDictLoanwordKorean_F = "패션쇼"
        Case "fastfood": CompanyName_KRDictLoanwordKorean_F = "패스트푸드"
        Case "fax": CompanyName_KRDictLoanwordKorean_F = "팩스"
        Case "feedback": CompanyName_KRDictLoanwordKorean_F = "피드백"
        Case "feet": CompanyName_KRDictLoanwordKorean_F = "피트"
        Case "feminism": CompanyName_KRDictLoanwordKorean_F = "페미니즘"
        Case "feminist": CompanyName_KRDictLoanwordKorean_F = "페미니스트"
        Case "fencing": CompanyName_KRDictLoanwordKorean_F = "펜싱"
        Case "festival": CompanyName_KRDictLoanwordKorean_F = "페스티벌"
        Case "fiction": CompanyName_KRDictLoanwordKorean_F = "픽션"
        Case "field": CompanyName_KRDictLoanwordKorean_F = "필드"
        Case "fieldhockey": CompanyName_KRDictLoanwordKorean_F = "필드하키"
        Case "figure": CompanyName_KRDictLoanwordKorean_F = "피겨"
        Case "figureskating": CompanyName_KRDictLoanwordKorean_F = "피겨스케이팅"
        Case "fiji": CompanyName_KRDictLoanwordKorean_F = "피지"
        Case "file": CompanyName_KRDictLoanwordKorean_F = "파일"
        Case "film": CompanyName_KRDictLoanwordKorean_F = "필름"
        Case "filter": CompanyName_KRDictLoanwordKorean_F = "필터"
        Case "finale": CompanyName_KRDictLoanwordKorean_F = "피날레"
        Case "finland": CompanyName_KRDictLoanwordKorean_F = "핀란드"
        Case "flash": CompanyName_KRDictLoanwordKorean_F = "플래시"
        Case "floppy": CompanyName_KRDictLoanwordKorean_F = "플로피"
        Case "floppydisk": CompanyName_KRDictLoanwordKorean_F = "플로피디스크"
        Case "flute": CompanyName_KRDictLoanwordKorean_F = "플루트"
        Case "folder": CompanyName_KRDictLoanwordKorean_F = "폴더"
        Case "food": CompanyName_KRDictLoanwordKorean_F = "푸드"
        Case "foodcourt": CompanyName_KRDictLoanwordKorean_F = "푸드코트"
        Case "foodtruck": CompanyName_KRDictLoanwordKorean_F = "푸드트럭"
        Case "football": CompanyName_KRDictLoanwordKorean_F = "풋볼"
        Case "fork": CompanyName_KRDictLoanwordKorean_F = "포크"
        Case "form": CompanyName_KRDictLoanwordKorean_F = "폼"
        Case "format": CompanyName_KRDictLoanwordKorean_F = "포맷"
        Case "forum": CompanyName_KRDictLoanwordKorean_F = "포럼"
        Case "foul": CompanyName_KRDictLoanwordKorean_F = "파울"
        Case "foundation": CompanyName_KRDictLoanwordKorean_F = "파운데이션"
        Case "france": CompanyName_KRDictLoanwordKorean_F = "프랑스"
        Case "franchise": CompanyName_KRDictLoanwordKorean_F = "프랜차이즈"
        Case "freelancer": CompanyName_KRDictLoanwordKorean_F = "프리랜서"
        Case "frenchtoast": CompanyName_KRDictLoanwordKorean_F = "프렌치토스트"
        Case "friedchicken": CompanyName_KRDictLoanwordKorean_F = "프라이드치킨"
        Case "front": CompanyName_KRDictLoanwordKorean_F = "프런트"
        Case "fry": CompanyName_KRDictLoanwordKorean_F = "프라이"
        Case "fulltime": CompanyName_KRDictLoanwordKorean_F = "풀타임"
        Case "fund": CompanyName_KRDictLoanwordKorean_F = "펀드"
        Case "fuse": CompanyName_KRDictLoanwordKorean_F = "퓨즈"
        Case "fusion": CompanyName_KRDictLoanwordKorean_F = "퓨전"
        Case Else: CompanyName_KRDictLoanwordKorean_F = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_G(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "gabon": CompanyName_KRDictLoanwordKorean_G = "가봉"
        Case "gallery": CompanyName_KRDictLoanwordKorean_G = "갤러리"
        Case "gallon": CompanyName_KRDictLoanwordKorean_G = "갤런"
        Case "game": CompanyName_KRDictLoanwordKorean_G = "게임"
        Case "gamer": CompanyName_KRDictLoanwordKorean_G = "게이머"
        Case "gang": CompanyName_KRDictLoanwordKorean_G = "갱"
        Case "gap": CompanyName_KRDictLoanwordKorean_G = "갭"
        Case "gas": CompanyName_KRDictLoanwordKorean_G = "가스"
        Case "gasboiler": CompanyName_KRDictLoanwordKorean_G = "가스보일러"
        Case "gasoline": CompanyName_KRDictLoanwordKorean_G = "가솔린"
        Case "gasrange": CompanyName_KRDictLoanwordKorean_G = "가스레인지"
        Case "gate": CompanyName_KRDictLoanwordKorean_G = "게이트"
        Case "gauge": CompanyName_KRDictLoanwordKorean_G = "게이지"
        Case "gauze": CompanyName_KRDictLoanwordKorean_G = "거즈"
        Case "gear": CompanyName_KRDictLoanwordKorean_G = "기어"
        Case "gelatin": CompanyName_KRDictLoanwordKorean_G = "젤라틴"
        Case "gender": CompanyName_KRDictLoanwordKorean_G = "젠더"
        Case "genre": CompanyName_KRDictLoanwordKorean_G = "장르"
        Case "gesture": CompanyName_KRDictLoanwordKorean_G = "제스처"
        Case "ghana": CompanyName_KRDictLoanwordKorean_G = "가나"
        Case "giga": CompanyName_KRDictLoanwordKorean_G = "기가"
        Case "gigabyte": CompanyName_KRDictLoanwordKorean_G = "기가바이트"
        Case "girdle": CompanyName_KRDictLoanwordKorean_G = "거들"
        Case "glass": CompanyName_KRDictLoanwordKorean_G = "글라스"
        Case "glider": CompanyName_KRDictLoanwordKorean_G = "글라이더"
        Case "global": CompanyName_KRDictLoanwordKorean_G = "글로벌"
        Case "glove": CompanyName_KRDictLoanwordKorean_G = "글러브"
        Case "goal": CompanyName_KRDictLoanwordKorean_G = "골"
        Case "goalkeeper": CompanyName_KRDictLoanwordKorean_G = "골키퍼"
        Case "goggles": CompanyName_KRDictLoanwordKorean_G = "고글"
        Case "golf": CompanyName_KRDictLoanwordKorean_G = "골프"
        Case "golfer": CompanyName_KRDictLoanwordKorean_G = "골퍼"
        Case "gondola": CompanyName_KRDictLoanwordKorean_G = "곤돌라"
        Case "goods": CompanyName_KRDictLoanwordKorean_G = "굿즈"
        Case "gorilla": CompanyName_KRDictLoanwordKorean_G = "고릴라"
        Case "gossip": CompanyName_KRDictLoanwordKorean_G = "가십"
        Case "gothic": CompanyName_KRDictLoanwordKorean_G = "고딕"
        Case "gown": CompanyName_KRDictLoanwordKorean_G = "가운"
        Case "gram": CompanyName_KRDictLoanwordKorean_G = "그램"
        Case "grandprix": CompanyName_KRDictLoanwordKorean_G = "그랑프리"
        Case "graph": CompanyName_KRDictLoanwordKorean_G = "그래프"
        Case "graphic": CompanyName_KRDictLoanwordKorean_G = "그래픽"
        Case "greece": CompanyName_KRDictLoanwordKorean_G = "그리스"
        Case "greenbelt": CompanyName_KRDictLoanwordKorean_G = "그린벨트"
        Case "grill": CompanyName_KRDictLoanwordKorean_G = "그릴"
        Case "ground": CompanyName_KRDictLoanwordKorean_G = "그라운드"
        Case "group": CompanyName_KRDictLoanwordKorean_G = "그룹"
        Case "guard": CompanyName_KRDictLoanwordKorean_G = "가드"
        Case "guardrail": CompanyName_KRDictLoanwordKorean_G = "가드레일"
        Case "guatemala": CompanyName_KRDictLoanwordKorean_G = "과테말라"
        Case "guest": CompanyName_KRDictLoanwordKorean_G = "게스트"
        Case "guide": CompanyName_KRDictLoanwordKorean_G = "가이드"
        Case "guidebook": CompanyName_KRDictLoanwordKorean_G = "가이드북"
        Case "guideline": CompanyName_KRDictLoanwordKorean_G = "가이드라인"
        Case "guinnessbook": CompanyName_KRDictLoanwordKorean_G = "기네스북"
        Case "guitar": CompanyName_KRDictLoanwordKorean_G = "기타"
        Case "gypsy": CompanyName_KRDictLoanwordKorean_G = "집시"
        Case Else: CompanyName_KRDictLoanwordKorean_G = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_H(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "hacker": CompanyName_KRDictLoanwordKorean_H = "해커"
        Case "hacking": CompanyName_KRDictLoanwordKorean_H = "해킹"
        Case "hairdryer": CompanyName_KRDictLoanwordKorean_H = "헤어드라이어"
        Case "hairstyle": CompanyName_KRDictLoanwordKorean_H = "헤어스타일"
        Case "hall": CompanyName_KRDictLoanwordKorean_H = "홀"
        Case "hallelujah": CompanyName_KRDictLoanwordKorean_H = "할렐루야"
        Case "ham": CompanyName_KRDictLoanwordKorean_H = "햄"
        Case "hamburger": CompanyName_KRDictLoanwordKorean_H = "햄버거"
        Case "hammer": CompanyName_KRDictLoanwordKorean_H = "해머"
        Case "hand": CompanyName_KRDictLoanwordKorean_H = "핸드"
        Case "handbag": CompanyName_KRDictLoanwordKorean_H = "핸드백"
        Case "handball": CompanyName_KRDictLoanwordKorean_H = "핸드볼"
        Case "handbrake": CompanyName_KRDictLoanwordKorean_H = "핸드브레이크"
        Case "handicap": CompanyName_KRDictLoanwordKorean_H = "핸디캡"
        Case "handle": CompanyName_KRDictLoanwordKorean_H = "핸들"
        Case "hangglider": CompanyName_KRDictLoanwordKorean_H = "행글라이더"
        Case "happening": CompanyName_KRDictLoanwordKorean_H = "해프닝"
        Case "hard": CompanyName_KRDictLoanwordKorean_H = "하드"
        Case "harddisk": CompanyName_KRDictLoanwordKorean_H = "하드디스크"
        Case "hardware": CompanyName_KRDictLoanwordKorean_H = "하드웨어"
        Case "harmonica": CompanyName_KRDictLoanwordKorean_H = "하모니카"
        Case "harmony": CompanyName_KRDictLoanwordKorean_H = "하모니"
        Case "harp": CompanyName_KRDictLoanwordKorean_H = "하프"
        Case "heading": CompanyName_KRDictLoanwordKorean_H = "헤딩"
        Case "headlight": CompanyName_KRDictLoanwordKorean_H = "헤드라이트"
        Case "headline": CompanyName_KRDictLoanwordKorean_H = "헤드라인"
        Case "headphone": CompanyName_KRDictLoanwordKorean_H = "헤드폰"
        Case "health": CompanyName_KRDictLoanwordKorean_H = "헬스"
        Case "healthcenter": CompanyName_KRDictLoanwordKorean_H = "헬스센터"
        Case "healthclub": CompanyName_KRDictLoanwordKorean_H = "헬스클럽"
        Case "heart": CompanyName_KRDictLoanwordKorean_H = "하트"
        Case "heater": CompanyName_KRDictLoanwordKorean_H = "히터"
        Case "heel": CompanyName_KRDictLoanwordKorean_H = "힐"
        Case "helicopter": CompanyName_KRDictLoanwordKorean_H = "헬리콥터"
        Case "helium": CompanyName_KRDictLoanwordKorean_H = "헬륨"
        Case "helmet": CompanyName_KRDictLoanwordKorean_H = "헬멧"
        Case "hertz": CompanyName_KRDictLoanwordKorean_H = "헤르츠"
        Case "highlight": CompanyName_KRDictLoanwordKorean_H = "하이라이트"
        Case "hightech": CompanyName_KRDictLoanwordKorean_H = "하이테크"
        Case "hiking": CompanyName_KRDictLoanwordKorean_H = "하이킹"
        Case "hint": CompanyName_KRDictLoanwordKorean_H = "힌트"
        Case "hip": CompanyName_KRDictLoanwordKorean_H = "히프"
        Case "hiphop": CompanyName_KRDictLoanwordKorean_H = "힙합"
        Case "hit": CompanyName_KRDictLoanwordKorean_H = "히트"
        Case "hockey": CompanyName_KRDictLoanwordKorean_H = "하키"
        Case "holiday": CompanyName_KRDictLoanwordKorean_H = "홀리데이"
        Case "home": CompanyName_KRDictLoanwordKorean_H = "홈"
        Case "homebanking": CompanyName_KRDictLoanwordKorean_H = "홈뱅킹"
        Case "homepage": CompanyName_KRDictLoanwordKorean_H = "홈페이지"
        Case "homerun": CompanyName_KRDictLoanwordKorean_H = "홈런"
        Case "homeshopping": CompanyName_KRDictLoanwordKorean_H = "홈쇼핑"
        Case "homo": CompanyName_KRDictLoanwordKorean_H = "호모"
        Case "hongkong": CompanyName_KRDictLoanwordKorean_H = "홍콩"
        Case "hop": CompanyName_KRDictLoanwordKorean_H = "홉"
        Case "hormone": CompanyName_KRDictLoanwordKorean_H = "호르몬"
        Case "horn": CompanyName_KRDictLoanwordKorean_H = "호른"
        Case "hose": CompanyName_KRDictLoanwordKorean_H = "호스"
        Case "hospice": CompanyName_KRDictLoanwordKorean_H = "호스피스"
        Case "hostel": CompanyName_KRDictLoanwordKorean_H = "호스텔"
        Case "hostess": CompanyName_KRDictLoanwordKorean_H = "호스티스"
        Case "hotdog": CompanyName_KRDictLoanwordKorean_H = "핫도그"
        Case "hotel": CompanyName_KRDictLoanwordKorean_H = "호텔"
        Case "hotpants": CompanyName_KRDictLoanwordKorean_H = "핫팬츠"
        Case "hub": CompanyName_KRDictLoanwordKorean_H = "허브"
        Case "humanist": CompanyName_KRDictLoanwordKorean_H = "휴머니스트"
        Case "humor": CompanyName_KRDictLoanwordKorean_H = "유머"
        Case "hungary": CompanyName_KRDictLoanwordKorean_H = "헝가리"
        Case "hurdle": CompanyName_KRDictLoanwordKorean_H = "허들"
        Case "hyena": CompanyName_KRDictLoanwordKorean_H = "하이에나"
        Case "hyphen": CompanyName_KRDictLoanwordKorean_H = "하이픈"
        Case Else: CompanyName_KRDictLoanwordKorean_H = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_I(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "icebox": CompanyName_KRDictLoanwordKorean_I = "아이스박스"
        Case "icecream": CompanyName_KRDictLoanwordKorean_I = "아이스크림"
        Case "icehockey": CompanyName_KRDictLoanwordKorean_I = "아이스하키"
        Case "iceland": CompanyName_KRDictLoanwordKorean_I = "아이슬란드"
        Case "icetea": CompanyName_KRDictLoanwordKorean_I = "아이스티"
        Case "icon": CompanyName_KRDictLoanwordKorean_I = "아이콘"
        Case "id": CompanyName_KRDictLoanwordKorean_I = "아이디"
        Case "idea": CompanyName_KRDictLoanwordKorean_I = "이데아"
        Case "idol": CompanyName_KRDictLoanwordKorean_I = "아이돌"
        Case "idolgroup": CompanyName_KRDictLoanwordKorean_I = "아이돌그룹"
        Case "iguana": CompanyName_KRDictLoanwordKorean_I = "이구아나"
        Case "illustration": CompanyName_KRDictLoanwordKorean_I = "일러스트레이션"
        Case "image": CompanyName_KRDictLoanwordKorean_I = "이미지"
        Case "implant": CompanyName_KRDictLoanwordKorean_I = "임플란트"
        Case "incentive": CompanyName_KRDictLoanwordKorean_I = "인센티브"
        Case "inch": CompanyName_KRDictLoanwordKorean_I = "인치"
        Case "indonesia": CompanyName_KRDictLoanwordKorean_I = "인도네시아"
        Case "induction": CompanyName_KRDictLoanwordKorean_I = "인덕션"
        Case "inflation": CompanyName_KRDictLoanwordKorean_I = "인플레이션"
        Case "influenza": CompanyName_KRDictLoanwordKorean_I = "인플루엔자"
        Case "initial": CompanyName_KRDictLoanwordKorean_I = "이니셜"
        Case "ink": CompanyName_KRDictLoanwordKorean_I = "잉크"
        Case "inning": CompanyName_KRDictLoanwordKorean_I = "이닝"
        Case "instant": CompanyName_KRDictLoanwordKorean_I = "인스턴트"
        Case "insulin": CompanyName_KRDictLoanwordKorean_I = "인슐린"
        Case "interchange": CompanyName_KRDictLoanwordKorean_I = "인터체인지"
        Case "interior": CompanyName_KRDictLoanwordKorean_I = "인테리어"
        Case "intern": CompanyName_KRDictLoanwordKorean_I = "인턴"
        Case "internet": CompanyName_KRDictLoanwordKorean_I = "인터넷"
        Case "internetbanking": CompanyName_KRDictLoanwordKorean_I = "인터넷뱅킹"
        Case "internetshopping": CompanyName_KRDictLoanwordKorean_I = "인터넷쇼핑"
        Case "internetsite": CompanyName_KRDictLoanwordKorean_I = "인터넷사이트"
        Case "internship": CompanyName_KRDictLoanwordKorean_I = "인턴십"
        Case "interphone": CompanyName_KRDictLoanwordKorean_I = "인터폰"
        Case "interview": CompanyName_KRDictLoanwordKorean_I = "인터뷰"
        Case "ion": CompanyName_KRDictLoanwordKorean_I = "이온"
        Case "iq": CompanyName_KRDictLoanwordKorean_I = "아이큐"
        Case "iran": CompanyName_KRDictLoanwordKorean_I = "이란"
        Case "iraq": CompanyName_KRDictLoanwordKorean_I = "이라크"
        Case "ireland": CompanyName_KRDictLoanwordKorean_I = "아일랜드"
        Case "irony": CompanyName_KRDictLoanwordKorean_I = "아이러니"
        Case "islam": CompanyName_KRDictLoanwordKorean_I = "이슬람"
        Case "israel": CompanyName_KRDictLoanwordKorean_I = "이스라엘"
        Case "issue": CompanyName_KRDictLoanwordKorean_I = "이슈"
        Case "it": CompanyName_KRDictLoanwordKorean_I = "아이티"
        Case "item": CompanyName_KRDictLoanwordKorean_I = "아이템"
        Case "ivory": CompanyName_KRDictLoanwordKorean_I = "아이보리"
        Case Else: CompanyName_KRDictLoanwordKorean_I = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_J(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "jacket": CompanyName_KRDictLoanwordKorean_J = "재킷"
        Case "jakarta": CompanyName_KRDictLoanwordKorean_J = "자카르타"
        Case "jam": CompanyName_KRDictLoanwordKorean_J = "잼"
        Case "jamaica": CompanyName_KRDictLoanwordKorean_J = "자메이카"
        Case "jazz": CompanyName_KRDictLoanwordKorean_J = "재즈"
        Case "jeep": CompanyName_KRDictLoanwordKorean_J = "지프"
        Case "jelly": CompanyName_KRDictLoanwordKorean_J = "젤리"
        Case "jinx": CompanyName_KRDictLoanwordKorean_J = "징크스"
        Case "jogging": CompanyName_KRDictLoanwordKorean_J = "조깅"
        Case "joke": CompanyName_KRDictLoanwordKorean_J = "조크"
        Case "jordan": CompanyName_KRDictLoanwordKorean_J = "요르단"
        Case "journalism": CompanyName_KRDictLoanwordKorean_J = "저널리즘"
        Case "journalist": CompanyName_KRDictLoanwordKorean_J = "저널리스트"
        Case "juice": CompanyName_KRDictLoanwordKorean_J = "주스"
        Case "jump": CompanyName_KRDictLoanwordKorean_J = "점프"
        Case "jumper": CompanyName_KRDictLoanwordKorean_J = "점퍼"
        Case "jungle": CompanyName_KRDictLoanwordKorean_J = "정글"
        Case Else: CompanyName_KRDictLoanwordKorean_J = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_K(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "kalium": CompanyName_KRDictLoanwordKorean_K = "칼륨"
        Case "kangaroo": CompanyName_KRDictLoanwordKorean_K = "캥거루"
        Case "karaoke": CompanyName_KRDictLoanwordKorean_K = "가라오케"
        Case "katusa": CompanyName_KRDictLoanwordKorean_K = "카투사"
        Case "kazakhstan": CompanyName_KRDictLoanwordKorean_K = "카자흐스탄"
        Case "kenya": CompanyName_KRDictLoanwordKorean_K = "케냐"
        Case "ketchup": CompanyName_KRDictLoanwordKorean_K = "케첩"
        Case "key": CompanyName_KRDictLoanwordKorean_K = "키"
        Case "keyboard": CompanyName_KRDictLoanwordKorean_K = "키보드"
        Case "keyword": CompanyName_KRDictLoanwordKorean_K = "키워드"
        Case "kickboard": CompanyName_KRDictLoanwordKorean_K = "킥보드"
        Case "kilo": CompanyName_KRDictLoanwordKorean_K = "킬로"
        Case "kilobyte": CompanyName_KRDictLoanwordKorean_K = "킬로바이트"
        Case "kilogram": CompanyName_KRDictLoanwordKorean_K = "킬로그램"
        Case "kilometer": CompanyName_KRDictLoanwordKorean_K = "킬로미터"
        Case "kilowatt": CompanyName_KRDictLoanwordKorean_K = "킬로와트"
        Case "kiosk": CompanyName_KRDictLoanwordKorean_K = "키오스크"
        Case "kiss": CompanyName_KRDictLoanwordKorean_K = "키스"
        Case "kit": CompanyName_KRDictLoanwordKorean_K = "키트"
        Case "kitchentowel": CompanyName_KRDictLoanwordKorean_K = "키친타월"
        Case "kiwi": CompanyName_KRDictLoanwordKorean_K = "키위"
        Case "knife": CompanyName_KRDictLoanwordKorean_K = "나이프"
        Case "knit": CompanyName_KRDictLoanwordKorean_K = "니트"
        Case "knock": CompanyName_KRDictLoanwordKorean_K = "노크"
        Case "knot": CompanyName_KRDictLoanwordKorean_K = "노트"
        Case "knowhow": CompanyName_KRDictLoanwordKorean_K = "노하우"
        Case "koala": CompanyName_KRDictLoanwordKorean_K = "코알라"
        Case "koran": CompanyName_KRDictLoanwordKorean_K = "코란"
        Case "kosdaq": CompanyName_KRDictLoanwordKorean_K = "코스닥"
        Case "kospi": CompanyName_KRDictLoanwordKorean_K = "코스피"
        Case "kpop": CompanyName_KRDictLoanwordKorean_K = "케이팝"
        Case "ks": CompanyName_KRDictLoanwordKorean_K = "케이에스"
        Case "kuwait": CompanyName_KRDictLoanwordKorean_K = "쿠웨이트"
        Case Else: CompanyName_KRDictLoanwordKorean_K = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_L(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "la": CompanyName_KRDictLoanwordKorean_L = "라"
        Case "label": CompanyName_KRDictLoanwordKorean_L = "라벨"
        Case "lace": CompanyName_KRDictLoanwordKorean_L = "레이스"
        Case "lameduck": CompanyName_KRDictLoanwordKorean_L = "레임덕"
        Case "lamp": CompanyName_KRDictLoanwordKorean_L = "램프"
        Case "lane": CompanyName_KRDictLoanwordKorean_L = "레인"
        Case "laos": CompanyName_KRDictLoanwordKorean_L = "라오스"
        Case "laptop": CompanyName_KRDictLoanwordKorean_L = "랩톱"
        Case "laptopcomputer": CompanyName_KRDictLoanwordKorean_L = "랩톱컴퓨터"
        Case "laser": CompanyName_KRDictLoanwordKorean_L = "레이저"
        Case "lasik": CompanyName_KRDictLoanwordKorean_L = "라식"
        Case "latte": CompanyName_KRDictLoanwordKorean_L = "라테"
        Case "lavender": CompanyName_KRDictLoanwordKorean_L = "라벤더"
        Case "lawfirm": CompanyName_KRDictLoanwordKorean_L = "로펌"
        Case "lead": CompanyName_KRDictLoanwordKorean_L = "리드"
        Case "leader": CompanyName_KRDictLoanwordKorean_L = "리더"
        Case "leadership": CompanyName_KRDictLoanwordKorean_L = "리더십"
        Case "league": CompanyName_KRDictLoanwordKorean_L = "리그"
        Case "lease": CompanyName_KRDictLoanwordKorean_L = "리스"
        Case "lebanon": CompanyName_KRDictLoanwordKorean_L = "레바논"
        Case "leggings": CompanyName_KRDictLoanwordKorean_L = "레깅스"
        Case "leisure": CompanyName_KRDictLoanwordKorean_L = "레저"
        Case "leisuresports": CompanyName_KRDictLoanwordKorean_L = "레포츠"
        Case "lemon": CompanyName_KRDictLoanwordKorean_L = "레몬"
        Case "lemonade": CompanyName_KRDictLoanwordKorean_L = "레모네이드"
        Case "lens": CompanyName_KRDictLoanwordKorean_L = "렌즈"
        Case "lesson": CompanyName_KRDictLoanwordKorean_L = "레슨"
        Case "level": CompanyName_KRDictLoanwordKorean_L = "레벨"
        Case "libya": CompanyName_KRDictLoanwordKorean_L = "리비아"
        Case "license": CompanyName_KRDictLoanwordKorean_L = "라이선스"
        Case "lift": CompanyName_KRDictLoanwordKorean_L = "리프트"
        Case "lighter": CompanyName_KRDictLoanwordKorean_L = "라이터"
        Case "lilac": CompanyName_KRDictLoanwordKorean_L = "라일락"
        Case "limousine": CompanyName_KRDictLoanwordKorean_L = "리무진"
        Case "line": CompanyName_KRDictLoanwordKorean_L = "라인"
        Case "lingerie": CompanyName_KRDictLoanwordKorean_L = "란제리"
        Case "link": CompanyName_KRDictLoanwordKorean_L = "링크"
        Case "lipbalm": CompanyName_KRDictLoanwordKorean_L = "립밤"
        Case "lipgloss": CompanyName_KRDictLoanwordKorean_L = "립글로스"
        Case "lipstick": CompanyName_KRDictLoanwordKorean_L = "립스틱"
        Case "lipsync": CompanyName_KRDictLoanwordKorean_L = "립싱크"
        Case "list": CompanyName_KRDictLoanwordKorean_L = "리스트"
        Case "liter": CompanyName_KRDictLoanwordKorean_L = "리터"
        Case "live": CompanyName_KRDictLoanwordKorean_L = "라이브"
        Case "lobby": CompanyName_KRDictLoanwordKorean_L = "로비"
        Case "lobster": CompanyName_KRDictLoanwordKorean_L = "랍스터"
        Case "logo": CompanyName_KRDictLoanwordKorean_L = "로고"
        Case "london": CompanyName_KRDictLoanwordKorean_L = "런던"
        Case "lotion": CompanyName_KRDictLoanwordKorean_L = "로션"
        Case "lounge": CompanyName_KRDictLoanwordKorean_L = "라운지"
        Case "lymph": CompanyName_KRDictLoanwordKorean_L = "림프"
        Case "lynch": CompanyName_KRDictLoanwordKorean_L = "린치"
        Case Else: CompanyName_KRDictLoanwordKorean_L = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_M(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "macaroni": CompanyName_KRDictLoanwordKorean_M = "마카로니"
        Case "mafia": CompanyName_KRDictLoanwordKorean_M = "마피아"
        Case "magma": CompanyName_KRDictLoanwordKorean_M = "마그마"
        Case "magnesium": CompanyName_KRDictLoanwordKorean_M = "마그네슘"
        Case "mail": CompanyName_KRDictLoanwordKorean_M = "메일"
        Case "maker": CompanyName_KRDictLoanwordKorean_M = "메이커"
        Case "makeup": CompanyName_KRDictLoanwordKorean_M = "메이크업"
        Case "malaria": CompanyName_KRDictLoanwordKorean_M = "말라리아"
        Case "malaysia": CompanyName_KRDictLoanwordKorean_M = "말레이시아"
        Case "maldives": CompanyName_KRDictLoanwordKorean_M = "몰디브"
        Case "mammoth": CompanyName_KRDictLoanwordKorean_M = "매머드"
        Case "manager": CompanyName_KRDictLoanwordKorean_M = "매니저"
        Case "mandolin": CompanyName_KRDictLoanwordKorean_M = "만돌린"
        Case "mango": CompanyName_KRDictLoanwordKorean_M = "망고"
        Case "manicure": CompanyName_KRDictLoanwordKorean_M = "매니큐어"
        Case "mannequin": CompanyName_KRDictLoanwordKorean_M = "마네킹"
        Case "manner": CompanyName_KRDictLoanwordKorean_M = "매너"
        Case "mannerism": CompanyName_KRDictLoanwordKorean_M = "매너리즘"
        Case "mansion": CompanyName_KRDictLoanwordKorean_M = "맨션"
        Case "manual": CompanyName_KRDictLoanwordKorean_M = "매뉴얼"
        Case "marathon": CompanyName_KRDictLoanwordKorean_M = "마라톤"
        Case "margarine": CompanyName_KRDictLoanwordKorean_M = "마가린"
        Case "margin": CompanyName_KRDictLoanwordKorean_M = "마진"
        Case "mark": CompanyName_KRDictLoanwordKorean_M = "마크"
        Case "marketing": CompanyName_KRDictLoanwordKorean_M = "마케팅"
        Case "marshmallow": CompanyName_KRDictLoanwordKorean_M = "마시멜로"
        Case "mart": CompanyName_KRDictLoanwordKorean_M = "마트"
        Case "mascara": CompanyName_KRDictLoanwordKorean_M = "마스카라"
        Case "mascot": CompanyName_KRDictLoanwordKorean_M = "마스코트"
        Case "mask": CompanyName_KRDictLoanwordKorean_M = "마스크"
        Case "massage": CompanyName_KRDictLoanwordKorean_M = "마사지"
        Case "master": CompanyName_KRDictLoanwordKorean_M = "마스터"
        Case "mat": CompanyName_KRDictLoanwordKorean_M = "매트"
        Case "match": CompanyName_KRDictLoanwordKorean_M = "매치"
        Case "mattress": CompanyName_KRDictLoanwordKorean_M = "매트리스"
        Case "mayonnaise": CompanyName_KRDictLoanwordKorean_M = "마요네즈"
        Case "mc": CompanyName_KRDictLoanwordKorean_M = "엠시"
        Case "mechanism": CompanyName_KRDictLoanwordKorean_M = "메커니즘"
        Case "medal": CompanyName_KRDictLoanwordKorean_M = "메달"
        Case "medalist": CompanyName_KRDictLoanwordKorean_M = "메달리스트"
        Case "media": CompanyName_KRDictLoanwordKorean_M = "미디어"
        Case "mega": CompanyName_KRDictLoanwordKorean_M = "메가"
        Case "megabyte": CompanyName_KRDictLoanwordKorean_M = "메가바이트"
        Case "megahertz": CompanyName_KRDictLoanwordKorean_M = "메가헤르츠"
        Case "melanin": CompanyName_KRDictLoanwordKorean_M = "멜라닌"
        Case "melodrama": CompanyName_KRDictLoanwordKorean_M = "멜로드라마"
        Case "melody": CompanyName_KRDictLoanwordKorean_M = "멜로디"
        Case "melon": CompanyName_KRDictLoanwordKorean_M = "멜론"
        Case "member": CompanyName_KRDictLoanwordKorean_M = "멤버"
        Case "membership": CompanyName_KRDictLoanwordKorean_M = "멤버십"
        Case "memo": CompanyName_KRDictLoanwordKorean_M = "메모"
        Case "memory": CompanyName_KRDictLoanwordKorean_M = "메모리"
        Case "mentor": CompanyName_KRDictLoanwordKorean_M = "멘토"
        Case "mentoring": CompanyName_KRDictLoanwordKorean_M = "멘토링"
        Case "menu": CompanyName_KRDictLoanwordKorean_M = "메뉴"
        Case "message": CompanyName_KRDictLoanwordKorean_M = "메시지"
        Case "messenger": CompanyName_KRDictLoanwordKorean_M = "메신저"
        Case "messiah": CompanyName_KRDictLoanwordKorean_M = "메시아"
        Case "meter": CompanyName_KRDictLoanwordKorean_M = "미터"
        Case "metronome": CompanyName_KRDictLoanwordKorean_M = "메트로놈"
        Case "mexico": CompanyName_KRDictLoanwordKorean_M = "멕시코"
        Case "mezzosoprano": CompanyName_KRDictLoanwordKorean_M = "메조소프라노"
        Case "mi": CompanyName_KRDictLoanwordKorean_M = "미"
        Case "micro": CompanyName_KRDictLoanwordKorean_M = "마이크로"
        Case "mile": CompanyName_KRDictLoanwordKorean_M = "마일"
        Case "mileage": CompanyName_KRDictLoanwordKorean_M = "마일리지"
        Case "milkshake": CompanyName_KRDictLoanwordKorean_M = "밀크셰이크"
        Case "milligram": CompanyName_KRDictLoanwordKorean_M = "밀리그램"
        Case "milliliter": CompanyName_KRDictLoanwordKorean_M = "밀리리터"
        Case "millimeter": CompanyName_KRDictLoanwordKorean_M = "밀리미터"
        Case "mindmap": CompanyName_KRDictLoanwordKorean_M = "마인드맵"
        Case "mineral": CompanyName_KRDictLoanwordKorean_M = "미네랄"
        Case "mini": CompanyName_KRDictLoanwordKorean_M = "미니"
        Case "minicar": CompanyName_KRDictLoanwordKorean_M = "미니카"
        Case "miniskirt": CompanyName_KRDictLoanwordKorean_M = "미니스커트"
        Case "mink": CompanyName_KRDictLoanwordKorean_M = "밍크"
        Case "mint": CompanyName_KRDictLoanwordKorean_M = "민트"
        Case "minuet": CompanyName_KRDictLoanwordKorean_M = "미뉴에트"
        Case "minus": CompanyName_KRDictLoanwordKorean_M = "마이너스"
        Case "miss": CompanyName_KRDictLoanwordKorean_M = "미스"
        Case "missile": CompanyName_KRDictLoanwordKorean_M = "미사일"
        Case "mission": CompanyName_KRDictLoanwordKorean_M = "미션"
        Case "mistermr": CompanyName_KRDictLoanwordKorean_M = "미스터"
        Case "mobile": CompanyName_KRDictLoanwordKorean_M = "모빌"
        Case "mobilebanking": CompanyName_KRDictLoanwordKorean_M = "모바일뱅킹"
        Case "mode": CompanyName_KRDictLoanwordKorean_M = "모드"
        Case "model": CompanyName_KRDictLoanwordKorean_M = "모델"
        Case "modem": CompanyName_KRDictLoanwordKorean_M = "모뎀"
        Case "modernism": CompanyName_KRDictLoanwordKorean_M = "모더니즘"
        Case "modernist": CompanyName_KRDictLoanwordKorean_M = "모더니스트"
        Case "monaco": CompanyName_KRDictLoanwordKorean_M = "모나코"
        Case "monitor": CompanyName_KRDictLoanwordKorean_M = "모니터"
        Case "monodrama": CompanyName_KRDictLoanwordKorean_M = "모노드라마"
        Case "monsoon": CompanyName_KRDictLoanwordKorean_M = "몬순"
        Case "montage": CompanyName_KRDictLoanwordKorean_M = "몽타주"
        Case "mood": CompanyName_KRDictLoanwordKorean_M = "무드"
        Case "morocco": CompanyName_KRDictLoanwordKorean_M = "모로코"
        Case "morphine": CompanyName_KRDictLoanwordKorean_M = "모르핀"
        Case "mosaic": CompanyName_KRDictLoanwordKorean_M = "모자이크"
        Case "motel": CompanyName_KRDictLoanwordKorean_M = "모텔"
        Case "motion": CompanyName_KRDictLoanwordKorean_M = "모션"
        Case "motive": CompanyName_KRDictLoanwordKorean_M = "모티브"
        Case "motor": CompanyName_KRDictLoanwordKorean_M = "모터"
        Case "motorboat": CompanyName_KRDictLoanwordKorean_M = "모터보트"
        Case "motto": CompanyName_KRDictLoanwordKorean_M = "모토"
        Case "mound": CompanyName_KRDictLoanwordKorean_M = "마운드"
        Case "mouse": CompanyName_KRDictLoanwordKorean_M = "마우스"
        Case "mousse": CompanyName_KRDictLoanwordKorean_M = "무스"
        Case "mozambique": CompanyName_KRDictLoanwordKorean_M = "모잠비크"
        Case "mrs": CompanyName_KRDictLoanwordKorean_M = "미시즈"
        Case "muffin": CompanyName_KRDictLoanwordKorean_M = "머핀"
        Case "muffler": CompanyName_KRDictLoanwordKorean_M = "머플러"
        Case "mug": CompanyName_KRDictLoanwordKorean_M = "머그"
        Case "multimedia": CompanyName_KRDictLoanwordKorean_M = "멀티미디어"
        Case "musical": CompanyName_KRDictLoanwordKorean_M = "뮤지컬"
        Case "muslim": CompanyName_KRDictLoanwordKorean_M = "무슬림"
        Case "mustard": CompanyName_KRDictLoanwordKorean_M = "머스터드"
        Case "myanmar": CompanyName_KRDictLoanwordKorean_M = "미얀마"
        Case "mycin": CompanyName_KRDictLoanwordKorean_M = "마이신"
        Case "mystery": CompanyName_KRDictLoanwordKorean_M = "미스터리"
        Case Else: CompanyName_KRDictLoanwordKorean_M = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_N(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "nail": CompanyName_KRDictLoanwordKorean_N = "네일"
        Case "nailart": CompanyName_KRDictLoanwordKorean_N = "네일아트"
        Case "nano": CompanyName_KRDictLoanwordKorean_N = "나노"
        Case "nanometer": CompanyName_KRDictLoanwordKorean_N = "나노미터"
        Case "naphthalene": CompanyName_KRDictLoanwordKorean_N = "나프탈렌"
        Case "napkin": CompanyName_KRDictLoanwordKorean_N = "냅킨"
        Case "narration": CompanyName_KRDictLoanwordKorean_N = "내레이션"
        Case "narrator": CompanyName_KRDictLoanwordKorean_N = "내레이터"
        Case "natrium": CompanyName_KRDictLoanwordKorean_N = "나트륨"
        Case "navigation": CompanyName_KRDictLoanwordKorean_N = "내비게이션"
        Case "necktie": CompanyName_KRDictLoanwordKorean_N = "넥타이"
        Case "neon": CompanyName_KRDictLoanwordKorean_N = "네온"
        Case "neonsign": CompanyName_KRDictLoanwordKorean_N = "네온사인"
        Case "nepal": CompanyName_KRDictLoanwordKorean_N = "네팔"
        Case "net": CompanyName_KRDictLoanwordKorean_N = "네트"
        Case "netherlands": CompanyName_KRDictLoanwordKorean_N = "네덜란드"
        Case "netizen": CompanyName_KRDictLoanwordKorean_N = "네티즌"
        Case "network": CompanyName_KRDictLoanwordKorean_N = "네트워크"
        Case "news": CompanyName_KRDictLoanwordKorean_N = "뉴스"
        Case "newyork": CompanyName_KRDictLoanwordKorean_N = "뉴욕"
        Case "newzealand": CompanyName_KRDictLoanwordKorean_N = "뉴질랜드"
        Case "ng": CompanyName_KRDictLoanwordKorean_N = "엔지"
        Case "nickel": CompanyName_KRDictLoanwordKorean_N = "니켈"
        Case "nicotine": CompanyName_KRDictLoanwordKorean_N = "니코틴"
        Case "nigeria": CompanyName_KRDictLoanwordKorean_N = "나이지리아"
        Case "night": CompanyName_KRDictLoanwordKorean_N = "나이트"
        Case "nightclub": CompanyName_KRDictLoanwordKorean_N = "나이트클럽"
        Case "nightgown": CompanyName_KRDictLoanwordKorean_N = "나이트가운"
        Case "nonfiction": CompanyName_KRDictLoanwordKorean_N = "논픽션"
        Case "nonsense": CompanyName_KRDictLoanwordKorean_N = "난센스"
        Case "nonstop": CompanyName_KRDictLoanwordKorean_N = "논스톱"
        Case "norway": CompanyName_KRDictLoanwordKorean_N = "노르웨이"
        Case "nostalgia": CompanyName_KRDictLoanwordKorean_N = "노스탤지어"
        Case "note": CompanyName_KRDictLoanwordKorean_N = "노트"
        Case "notebook": CompanyName_KRDictLoanwordKorean_N = "노트북"
        Case "nuance": CompanyName_KRDictLoanwordKorean_N = "뉘앙스"
        Case "nude": CompanyName_KRDictLoanwordKorean_N = "누드"
        Case "number": CompanyName_KRDictLoanwordKorean_N = "넘버"
        Case "nut": CompanyName_KRDictLoanwordKorean_N = "너트"
        Case "nylon": CompanyName_KRDictLoanwordKorean_N = "나일론"
        Case Else: CompanyName_KRDictLoanwordKorean_N = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_O(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "oak": CompanyName_KRDictLoanwordKorean_O = "오크"
        Case "oasis": CompanyName_KRDictLoanwordKorean_O = "오아시스"
        Case "oboe": CompanyName_KRDictLoanwordKorean_O = "오보에"
        Case "oceania": CompanyName_KRDictLoanwordKorean_O = "오세아니아"
        Case "octave": CompanyName_KRDictLoanwordKorean_O = "옥타브"
        Case "offer": CompanyName_KRDictLoanwordKorean_O = "오퍼"
        Case "offline": CompanyName_KRDictLoanwordKorean_O = "오프라인"
        Case "olive": CompanyName_KRDictLoanwordKorean_O = "올리브"
        Case "omelet": CompanyName_KRDictLoanwordKorean_O = "오믈렛"
        Case "onepiece": CompanyName_KRDictLoanwordKorean_O = "원피스"
        Case "online": CompanyName_KRDictLoanwordKorean_O = "온라인"
        Case "opening": CompanyName_KRDictLoanwordKorean_O = "오프닝"
        Case "opera": CompanyName_KRDictLoanwordKorean_O = "오페라"
        Case "option": CompanyName_KRDictLoanwordKorean_O = "옵션"
        Case "orange": CompanyName_KRDictLoanwordKorean_O = "오렌지"
        Case "orangutan": CompanyName_KRDictLoanwordKorean_O = "오랑우탄"
        Case "orchestra": CompanyName_KRDictLoanwordKorean_O = "오케스트라"
        Case "order": CompanyName_KRDictLoanwordKorean_O = "오더"
        Case "organ": CompanyName_KRDictLoanwordKorean_O = "오르간"
        Case "orientation": CompanyName_KRDictLoanwordKorean_O = "오리엔테이션"
        Case "original": CompanyName_KRDictLoanwordKorean_O = "오리지널"
        Case "ounce": CompanyName_KRDictLoanwordKorean_O = "온스"
        Case "out": CompanyName_KRDictLoanwordKorean_O = "아웃"
        Case "outlet": CompanyName_KRDictLoanwordKorean_O = "아웃렛"
        Case "outsider": CompanyName_KRDictLoanwordKorean_O = "아웃사이더"
        Case "oven": CompanyName_KRDictLoanwordKorean_O = "오븐"
        Case "over": CompanyName_KRDictLoanwordKorean_O = "오버"
        Case "ozone": CompanyName_KRDictLoanwordKorean_O = "오존"
        Case Else: CompanyName_KRDictLoanwordKorean_O = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_P(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "pace": CompanyName_KRDictLoanwordKorean_P = "페이스"
        Case "pack": CompanyName_KRDictLoanwordKorean_P = "팩"
        Case "package": CompanyName_KRDictLoanwordKorean_P = "패키지"
        Case "packing": CompanyName_KRDictLoanwordKorean_P = "패킹"
        Case "pad": CompanyName_KRDictLoanwordKorean_P = "패드"
        Case "padding": CompanyName_KRDictLoanwordKorean_P = "패딩"
        Case "paddingjumper": CompanyName_KRDictLoanwordKorean_P = "패딩점퍼"
        Case "page": CompanyName_KRDictLoanwordKorean_P = "페이지"
        Case "paint": CompanyName_KRDictLoanwordKorean_P = "페인트"
        Case "pajamas": CompanyName_KRDictLoanwordKorean_P = "파자마"
        Case "pakistan": CompanyName_KRDictLoanwordKorean_P = "파키스탄"
        Case "palette": CompanyName_KRDictLoanwordKorean_P = "팔레트"
        Case "pamphlet": CompanyName_KRDictLoanwordKorean_P = "팸플릿"
        Case "pan": CompanyName_KRDictLoanwordKorean_P = "팬"
        Case "pancake": CompanyName_KRDictLoanwordKorean_P = "팬케이크"
        Case "panda": CompanyName_KRDictLoanwordKorean_P = "판다"
        Case "panel": CompanyName_KRDictLoanwordKorean_P = "패널"
        Case "panorama": CompanyName_KRDictLoanwordKorean_P = "파노라마"
        Case "pansy": CompanyName_KRDictLoanwordKorean_P = "팬지"
        Case "pantomime": CompanyName_KRDictLoanwordKorean_P = "팬터마임"
        Case "pants": CompanyName_KRDictLoanwordKorean_P = "팬츠"
        Case "papuanewguinea": CompanyName_KRDictLoanwordKorean_P = "파푸아뉴기니"
        Case "papyrus": CompanyName_KRDictLoanwordKorean_P = "파피루스"
        Case "parade": CompanyName_KRDictLoanwordKorean_P = "퍼레이드"
        Case "paradigm": CompanyName_KRDictLoanwordKorean_P = "패러다임"
        Case "paradise": CompanyName_KRDictLoanwordKorean_P = "파라다이스"
        Case "paradox": CompanyName_KRDictLoanwordKorean_P = "패러독스"
        Case "paraguay": CompanyName_KRDictLoanwordKorean_P = "파라과이"
        Case "parasol": CompanyName_KRDictLoanwordKorean_P = "파라솔"
        Case "parka": CompanyName_KRDictLoanwordKorean_P = "파카"
        Case "parody": CompanyName_KRDictLoanwordKorean_P = "패러디"
        Case "parsley": CompanyName_KRDictLoanwordKorean_P = "파슬리"
        Case "part": CompanyName_KRDictLoanwordKorean_P = "파트"
        Case "partition": CompanyName_KRDictLoanwordKorean_P = "파티션"
        Case "partner": CompanyName_KRDictLoanwordKorean_P = "파트너"
        Case "parttime": CompanyName_KRDictLoanwordKorean_P = "파트타임"
        Case "party": CompanyName_KRDictLoanwordKorean_P = "파티"
        Case "pass": CompanyName_KRDictLoanwordKorean_P = "패스"
        Case "passport": CompanyName_KRDictLoanwordKorean_P = "패스포트"
        Case "pasta": CompanyName_KRDictLoanwordKorean_P = "파스타"
        Case "pastel": CompanyName_KRDictLoanwordKorean_P = "파스텔"
        Case "patch": CompanyName_KRDictLoanwordKorean_P = "패치"
        Case "pattern": CompanyName_KRDictLoanwordKorean_P = "패턴"
        Case "pc": CompanyName_KRDictLoanwordKorean_P = "피시"
        Case "pd": CompanyName_KRDictLoanwordKorean_P = "피디"
        Case "peak": CompanyName_KRDictLoanwordKorean_P = "피크"
        Case "pedal": CompanyName_KRDictLoanwordKorean_P = "페달"
        Case "pen": CompanyName_KRDictLoanwordKorean_P = "펜"
        Case "penalty": CompanyName_KRDictLoanwordKorean_P = "페널티"
        Case "pendant": CompanyName_KRDictLoanwordKorean_P = "펜던트"
        Case "penguin": CompanyName_KRDictLoanwordKorean_P = "펭귄"
        Case "penpal": CompanyName_KRDictLoanwordKorean_P = "펜팔"
        Case "pension": CompanyName_KRDictLoanwordKorean_P = "펜션"
        Case "percent": CompanyName_KRDictLoanwordKorean_P = "퍼센트"
        Case "performance": CompanyName_KRDictLoanwordKorean_P = "퍼포먼스"
        Case "peru": CompanyName_KRDictLoanwordKorean_P = "페루"
        Case "pest": CompanyName_KRDictLoanwordKorean_P = "페스트"
        Case "petticoat": CompanyName_KRDictLoanwordKorean_P = "페티코트"
        Case "philippines": CompanyName_KRDictLoanwordKorean_P = "필리핀"
        Case "phishing": CompanyName_KRDictLoanwordKorean_P = "피싱"
        Case "phone": CompanyName_KRDictLoanwordKorean_P = "폰"
        Case "phonebanking": CompanyName_KRDictLoanwordKorean_P = "폰뱅킹"
        Case "pianist": CompanyName_KRDictLoanwordKorean_P = "피아니스트"
        Case "piano": CompanyName_KRDictLoanwordKorean_P = "피아노"
        Case "picket": CompanyName_KRDictLoanwordKorean_P = "피켓"
        Case "pickle": CompanyName_KRDictLoanwordKorean_P = "피클"
        Case "pie": CompanyName_KRDictLoanwordKorean_P = "파이"
        Case "pilot": CompanyName_KRDictLoanwordKorean_P = "파일럿"
        Case "pin": CompanyName_KRDictLoanwordKorean_P = "핀"
        Case "pineapple": CompanyName_KRDictLoanwordKorean_P = "파인애플"
        Case "pink": CompanyName_KRDictLoanwordKorean_P = "핑크"
        Case "pipe": CompanyName_KRDictLoanwordKorean_P = "파이프"
        Case "pipeorgan": CompanyName_KRDictLoanwordKorean_P = "파이프오르간"
        Case "piston": CompanyName_KRDictLoanwordKorean_P = "피스톤"
        Case "pitching": CompanyName_KRDictLoanwordKorean_P = "피칭"
        Case "pizza": CompanyName_KRDictLoanwordKorean_P = "피자"
        Case "placard": CompanyName_KRDictLoanwordKorean_P = "플래카드"
        Case "plankton": CompanyName_KRDictLoanwordKorean_P = "플랑크톤"
        Case "plastic": CompanyName_KRDictLoanwordKorean_P = "플라스틱"
        Case "platform": CompanyName_KRDictLoanwordKorean_P = "플랫폼"
        Case "play": CompanyName_KRDictLoanwordKorean_P = "플레이"
        Case "plot": CompanyName_KRDictLoanwordKorean_P = "플롯"
        Case "plug": CompanyName_KRDictLoanwordKorean_P = "플러그"
        Case "plus": CompanyName_KRDictLoanwordKorean_P = "플러스"
        Case "point": CompanyName_KRDictLoanwordKorean_P = "포인트"
        Case "poland": CompanyName_KRDictLoanwordKorean_P = "폴란드"
        Case "pool": CompanyName_KRDictLoanwordKorean_P = "풀"
        Case "pop": CompanyName_KRDictLoanwordKorean_P = "팝"
        Case "popcorn": CompanyName_KRDictLoanwordKorean_P = "팝콘"
        Case "porkcutlet": CompanyName_KRDictLoanwordKorean_P = "포크커틀릿"
        Case "portal": CompanyName_KRDictLoanwordKorean_P = "포털"
        Case "portalsite": CompanyName_KRDictLoanwordKorean_P = "포털사이트"
        Case "portfolio": CompanyName_KRDictLoanwordKorean_P = "포트폴리오"
        Case "portugal": CompanyName_KRDictLoanwordKorean_P = "포르투갈"
        Case "pose": CompanyName_KRDictLoanwordKorean_P = "포즈"
        Case "position": CompanyName_KRDictLoanwordKorean_P = "포지션"
        Case "poster": CompanyName_KRDictLoanwordKorean_P = "포스터"
        Case "pound": CompanyName_KRDictLoanwordKorean_P = "파운드"
        Case "powder": CompanyName_KRDictLoanwordKorean_P = "파우더"
        Case "power": CompanyName_KRDictLoanwordKorean_P = "파워"
        Case "premium": CompanyName_KRDictLoanwordKorean_P = "프리미엄"
        Case "presentation": CompanyName_KRDictLoanwordKorean_P = "프레젠테이션"
        Case "pride": CompanyName_KRDictLoanwordKorean_P = "프라이드"
        Case "print": CompanyName_KRDictLoanwordKorean_P = "프린트"
        Case "printer": CompanyName_KRDictLoanwordKorean_P = "프린터"
        Case "prism": CompanyName_KRDictLoanwordKorean_P = "프리즘"
        Case "privacy": CompanyName_KRDictLoanwordKorean_P = "프라이버시"
        Case "processor": CompanyName_KRDictLoanwordKorean_P = "프로세서"
        Case "producer": CompanyName_KRDictLoanwordKorean_P = "프로듀서"
        Case "production": CompanyName_KRDictLoanwordKorean_P = "프로덕션"
        Case "profile": CompanyName_KRDictLoanwordKorean_P = "프로필"
        Case "program": CompanyName_KRDictLoanwordKorean_P = "프로그램"
        Case "programmer": CompanyName_KRDictLoanwordKorean_P = "프로그래머"
        Case "programming": CompanyName_KRDictLoanwordKorean_P = "프로그래밍"
        Case "project": CompanyName_KRDictLoanwordKorean_P = "프로젝트"
        Case "projector": CompanyName_KRDictLoanwordKorean_P = "프로젝터"
        Case "propane": CompanyName_KRDictLoanwordKorean_P = "프로판"
        Case "propeller": CompanyName_KRDictLoanwordKorean_P = "프로펠러"
        Case "psycho": CompanyName_KRDictLoanwordKorean_P = "사이코"
        Case "pudding": CompanyName_KRDictLoanwordKorean_P = "푸딩"
        Case "puertorico": CompanyName_KRDictLoanwordKorean_P = "푸에르토리코"
        Case "pump": CompanyName_KRDictLoanwordKorean_P = "펌프"
        Case "punch": CompanyName_KRDictLoanwordKorean_P = "펀치"
        Case "puzzle": CompanyName_KRDictLoanwordKorean_P = "퍼즐"
        Case "px": CompanyName_KRDictLoanwordKorean_P = "피엑스"
        Case "pyramid": CompanyName_KRDictLoanwordKorean_P = "피라미드"
        Case Else: CompanyName_KRDictLoanwordKorean_P = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_Q(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "qatar": CompanyName_KRDictLoanwordKorean_Q = "카타르"
        Case "quarter": CompanyName_KRDictLoanwordKorean_Q = "쿼터"
        Case "quilt": CompanyName_KRDictLoanwordKorean_Q = "퀼트"
        Case "quiz": CompanyName_KRDictLoanwordKorean_Q = "퀴즈"
        Case "quota": CompanyName_KRDictLoanwordKorean_Q = "쿼터"
        Case Else: CompanyName_KRDictLoanwordKorean_Q = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_R(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "race": CompanyName_KRDictLoanwordKorean_R = "레이스"
        Case "racket": CompanyName_KRDictLoanwordKorean_R = "라켓"
        Case "radar": CompanyName_KRDictLoanwordKorean_R = "레이다"
        Case "radiator": CompanyName_KRDictLoanwordKorean_R = "라디에이터"
        Case "radio": CompanyName_KRDictLoanwordKorean_R = "라디오"
        Case "rafting": CompanyName_KRDictLoanwordKorean_R = "래프팅"
        Case "rail": CompanyName_KRDictLoanwordKorean_R = "레일"
        Case "ranking": CompanyName_KRDictLoanwordKorean_R = "랭킹"
        Case "rap": CompanyName_KRDictLoanwordKorean_R = "랩"
        Case "re": CompanyName_KRDictLoanwordKorean_R = "레"
        Case "realism": CompanyName_KRDictLoanwordKorean_R = "리얼리즘"
        Case "reality": CompanyName_KRDictLoanwordKorean_R = "리얼리티"
        Case "rebound": CompanyName_KRDictLoanwordKorean_R = "리바운드"
        Case "recall": CompanyName_KRDictLoanwordKorean_R = "리콜"
        Case "reception": CompanyName_KRDictLoanwordKorean_R = "리셉션"
        Case "recital": CompanyName_KRDictLoanwordKorean_R = "리사이틀"
        Case "record": CompanyName_KRDictLoanwordKorean_R = "레코드"
        Case "recorder": CompanyName_KRDictLoanwordKorean_R = "리코더"
        Case "recreation": CompanyName_KRDictLoanwordKorean_R = "레크리에이션"
        Case "redcard": CompanyName_KRDictLoanwordKorean_R = "레드카드"
        Case "refill": CompanyName_KRDictLoanwordKorean_R = "리필"
        Case "rehearsal": CompanyName_KRDictLoanwordKorean_R = "리허설"
        Case "relay": CompanyName_KRDictLoanwordKorean_R = "릴레이"
        Case "remake": CompanyName_KRDictLoanwordKorean_R = "리메이크"
        Case "remodeling": CompanyName_KRDictLoanwordKorean_R = "리모델링"
        Case "renaissance": CompanyName_KRDictLoanwordKorean_R = "르네상스"
        Case "rendezvous": CompanyName_KRDictLoanwordKorean_R = "랑데부"
        Case "renewal": CompanyName_KRDictLoanwordKorean_R = "리뉴얼"
        Case "report": CompanyName_KRDictLoanwordKorean_R = "리포트"
        Case "reporter": CompanyName_KRDictLoanwordKorean_R = "리포터"
        Case "research": CompanyName_KRDictLoanwordKorean_R = "리서치"
        Case "resident": CompanyName_KRDictLoanwordKorean_R = "레지던트"
        Case "resort": CompanyName_KRDictLoanwordKorean_R = "리조트"
        Case "restaurant": CompanyName_KRDictLoanwordKorean_R = "레스토랑"
        Case "review": CompanyName_KRDictLoanwordKorean_R = "리뷰"
        Case "revival": CompanyName_KRDictLoanwordKorean_R = "리바이벌"
        Case "rheumatism": CompanyName_KRDictLoanwordKorean_R = "류머티즘"
        Case "rhythm": CompanyName_KRDictLoanwordKorean_R = "리듬"
        Case "ribbon": CompanyName_KRDictLoanwordKorean_R = "리본"
        Case "ring": CompanyName_KRDictLoanwordKorean_R = "링"
        Case "rink": CompanyName_KRDictLoanwordKorean_R = "링크"
        Case "rinse": CompanyName_KRDictLoanwordKorean_R = "린스"
        Case "risk": CompanyName_KRDictLoanwordKorean_R = "리스크"
        Case "rival": CompanyName_KRDictLoanwordKorean_R = "라이벌"
        Case "roaming": CompanyName_KRDictLoanwordKorean_R = "로밍"
        Case "robot": CompanyName_KRDictLoanwordKorean_R = "로봇"
        Case "rock": CompanyName_KRDictLoanwordKorean_R = "록"
        Case "rocket": CompanyName_KRDictLoanwordKorean_R = "로켓"
        Case "roller": CompanyName_KRDictLoanwordKorean_R = "롤러"
        Case "rollercoaster": CompanyName_KRDictLoanwordKorean_R = "롤러코스터"
        Case "rollerskate": CompanyName_KRDictLoanwordKorean_R = "롤러스케이트"
        Case "romance": CompanyName_KRDictLoanwordKorean_R = "로맨스"
        Case "romania": CompanyName_KRDictLoanwordKorean_R = "루마니아"
        Case "roommate": CompanyName_KRDictLoanwordKorean_R = "룸메이트"
        Case "rope": CompanyName_KRDictLoanwordKorean_R = "로프"
        Case "round": CompanyName_KRDictLoanwordKorean_R = "라운드"
        Case "route": CompanyName_KRDictLoanwordKorean_R = "루트"
        Case "royalty": CompanyName_KRDictLoanwordKorean_R = "로열티"
        Case "ruby": CompanyName_KRDictLoanwordKorean_R = "루비"
        Case "rugby": CompanyName_KRDictLoanwordKorean_R = "럭비"
        Case "rugbyfootball": CompanyName_KRDictLoanwordKorean_R = "럭비풋볼"
        Case "rumor": CompanyName_KRDictLoanwordKorean_R = "루머"
        Case "rushhour": CompanyName_KRDictLoanwordKorean_R = "러시아워"
        Case "russia": CompanyName_KRDictLoanwordKorean_R = "러시아"
        Case "rwanda": CompanyName_KRDictLoanwordKorean_R = "르완다"
        Case Else: CompanyName_KRDictLoanwordKorean_R = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_S(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "safari": CompanyName_KRDictLoanwordKorean_S = "사파리"
        Case "safe": CompanyName_KRDictLoanwordKorean_S = "세이프"
        Case "salad": CompanyName_KRDictLoanwordKorean_S = "샐러드"
        Case "sale": CompanyName_KRDictLoanwordKorean_S = "세일"
        Case "salesman": CompanyName_KRDictLoanwordKorean_S = "세일즈맨"
        Case "salon": CompanyName_KRDictLoanwordKorean_S = "살롱"
        Case "salsa": CompanyName_KRDictLoanwordKorean_S = "살사"
        Case "sample": CompanyName_KRDictLoanwordKorean_S = "샘플"
        Case "sandal": CompanyName_KRDictLoanwordKorean_S = "샌들"
        Case "sandwich": CompanyName_KRDictLoanwordKorean_S = "샌드위치"
        Case "santaclaus": CompanyName_KRDictLoanwordKorean_S = "산타클로스"
        Case "sapphire": CompanyName_KRDictLoanwordKorean_S = "사파이어"
        Case "satan": CompanyName_KRDictLoanwordKorean_S = "사탄"
        Case "sauce": CompanyName_KRDictLoanwordKorean_S = "소스"
        Case "saudiarabia": CompanyName_KRDictLoanwordKorean_S = "사우디아라비아"
        Case "sauna": CompanyName_KRDictLoanwordKorean_S = "사우나"
        Case "sausage": CompanyName_KRDictLoanwordKorean_S = "소시지"
        Case "saxophone": CompanyName_KRDictLoanwordKorean_S = "색소폰"
        Case "scale": CompanyName_KRDictLoanwordKorean_S = "스케일"
        Case "scaling": CompanyName_KRDictLoanwordKorean_S = "스케일링"
        Case "scandal": CompanyName_KRDictLoanwordKorean_S = "스캔들"
        Case "scanner": CompanyName_KRDictLoanwordKorean_S = "스캐너"
        Case "scarf": CompanyName_KRDictLoanwordKorean_S = "스카프"
        Case "scenario": CompanyName_KRDictLoanwordKorean_S = "시나리오"
        Case "scene": CompanyName_KRDictLoanwordKorean_S = "신"
        Case "schedule": CompanyName_KRDictLoanwordKorean_S = "스케줄"
        Case "schoolbus": CompanyName_KRDictLoanwordKorean_S = "스쿨버스"
        Case "scooter": CompanyName_KRDictLoanwordKorean_S = "스쿠터"
        Case "score": CompanyName_KRDictLoanwordKorean_S = "스코어"
        Case "scotchtape": CompanyName_KRDictLoanwordKorean_S = "스카치테이프"
        Case "scout": CompanyName_KRDictLoanwordKorean_S = "스카우트"
        Case "scrap": CompanyName_KRDictLoanwordKorean_S = "스크랩"
        Case "scrapbook": CompanyName_KRDictLoanwordKorean_S = "스크랩북"
        Case "screen": CompanyName_KRDictLoanwordKorean_S = "스크린"
        Case "screendoor": CompanyName_KRDictLoanwordKorean_S = "스크린도어"
        Case "scrum": CompanyName_KRDictLoanwordKorean_S = "스크럼"
        Case "scuba": CompanyName_KRDictLoanwordKorean_S = "스쿠버"
        Case "scubadiving": CompanyName_KRDictLoanwordKorean_S = "스쿠버다이빙"
        Case "searchlight": CompanyName_KRDictLoanwordKorean_S = "서치라이트"
        Case "season": CompanyName_KRDictLoanwordKorean_S = "시즌"
        Case "seesaw": CompanyName_KRDictLoanwordKorean_S = "시소"
        Case "self": CompanyName_KRDictLoanwordKorean_S = "셀프"
        Case "selfservice": CompanyName_KRDictLoanwordKorean_S = "셀프서비스"
        Case "seminar": CompanyName_KRDictLoanwordKorean_S = "세미나"
        Case "senegal": CompanyName_KRDictLoanwordKorean_S = "세네갈"
        Case "sense": CompanyName_KRDictLoanwordKorean_S = "센스"
        Case "sensor": CompanyName_KRDictLoanwordKorean_S = "센서"
        Case "serbia": CompanyName_KRDictLoanwordKorean_S = "세르비아"
        Case "serenade": CompanyName_KRDictLoanwordKorean_S = "세레나데"
        Case "series": CompanyName_KRDictLoanwordKorean_S = "시리즈"
        Case "serve": CompanyName_KRDictLoanwordKorean_S = "서브"
        Case "service": CompanyName_KRDictLoanwordKorean_S = "서비스"
        Case "servicecenter": CompanyName_KRDictLoanwordKorean_S = "서비스센터"
        Case "serving": CompanyName_KRDictLoanwordKorean_S = "서빙"
        Case "set": CompanyName_KRDictLoanwordKorean_S = "세트"
        Case "sex": CompanyName_KRDictLoanwordKorean_S = "섹스"
        Case "shabushabu": CompanyName_KRDictLoanwordKorean_S = "샤부샤부"
        Case "shamanism": CompanyName_KRDictLoanwordKorean_S = "샤머니즘"
        Case "shampoo": CompanyName_KRDictLoanwordKorean_S = "샴푸"
        Case "sheet": CompanyName_KRDictLoanwordKorean_S = "시트"
        Case "shock": CompanyName_KRDictLoanwordKorean_S = "쇼크"
        Case "shop": CompanyName_KRDictLoanwordKorean_S = "숍"
        Case "shopping": CompanyName_KRDictLoanwordKorean_S = "쇼핑"
        Case "shoppingbag": CompanyName_KRDictLoanwordKorean_S = "쇼핑백"
        Case "shoppingcenter": CompanyName_KRDictLoanwordKorean_S = "쇼핑센터"
        Case "shoppingmall": CompanyName_KRDictLoanwordKorean_S = "쇼핑몰"
        Case "short": CompanyName_KRDictLoanwordKorean_S = "쇼트"
        Case "shorttrack": CompanyName_KRDictLoanwordKorean_S = "쇼트트랙"
        Case "shot": CompanyName_KRDictLoanwordKorean_S = "샷"
        Case "shoulderbag": CompanyName_KRDictLoanwordKorean_S = "숄더백"
        Case "show": CompanyName_KRDictLoanwordKorean_S = "쇼"
        Case "shower": CompanyName_KRDictLoanwordKorean_S = "샤워"
        Case "showhost": CompanyName_KRDictLoanwordKorean_S = "쇼호스트"
        Case "showwindow": CompanyName_KRDictLoanwordKorean_S = "쇼윈도"
        Case "shutter": CompanyName_KRDictLoanwordKorean_S = "셔터"
        Case "shuttlebus": CompanyName_KRDictLoanwordKorean_S = "셔틀버스"
        Case "shuttlecock": CompanyName_KRDictLoanwordKorean_S = "셔틀콕"
        Case "si": CompanyName_KRDictLoanwordKorean_S = "시"
        Case "sign": CompanyName_KRDictLoanwordKorean_S = "사인"
        Case "signal": CompanyName_KRDictLoanwordKorean_S = "시그널"
        Case "silhouette": CompanyName_KRDictLoanwordKorean_S = "실루엣"
        Case "silicon": CompanyName_KRDictLoanwordKorean_S = "실리콘"
        Case "silk": CompanyName_KRDictLoanwordKorean_S = "실크"
        Case "simulation": CompanyName_KRDictLoanwordKorean_S = "시뮬레이션"
        Case "singapore": CompanyName_KRDictLoanwordKorean_S = "싱가포르"
        Case "single": CompanyName_KRDictLoanwordKorean_S = "싱글"
        Case "siren": CompanyName_KRDictLoanwordKorean_S = "사이렌"
        Case "sitcom": CompanyName_KRDictLoanwordKorean_S = "시트콤"
        Case "site": CompanyName_KRDictLoanwordKorean_S = "사이트"
        Case "size": CompanyName_KRDictLoanwordKorean_S = "사이즈"
        Case "skate": CompanyName_KRDictLoanwordKorean_S = "스케이트"
        Case "skateboard": CompanyName_KRDictLoanwordKorean_S = "스케이트보드"
        Case "skating": CompanyName_KRDictLoanwordKorean_S = "스케이팅"
        Case "sketch": CompanyName_KRDictLoanwordKorean_S = "스케치"
        Case "sketchbook": CompanyName_KRDictLoanwordKorean_S = "스케치북"
        Case "ski": CompanyName_KRDictLoanwordKorean_S = "스키"
        Case "skinlotion": CompanyName_KRDictLoanwordKorean_S = "스킨로션"
        Case "skirt": CompanyName_KRDictLoanwordKorean_S = "스커트"
        Case "skydiving": CompanyName_KRDictLoanwordKorean_S = "스카이다이빙"
        Case "slide": CompanyName_KRDictLoanwordKorean_S = "슬라이드"
        Case "sliding": CompanyName_KRDictLoanwordKorean_S = "슬라이딩"
        Case "slipper": CompanyName_KRDictLoanwordKorean_S = "슬리퍼"
        Case "slogan": CompanyName_KRDictLoanwordKorean_S = "슬로건"
        Case "slovakia": CompanyName_KRDictLoanwordKorean_S = "슬로바키아"
        Case "slovenia": CompanyName_KRDictLoanwordKorean_S = "슬로베니아"
        Case "slump": CompanyName_KRDictLoanwordKorean_S = "슬럼프"
        Case "smartphone": CompanyName_KRDictLoanwordKorean_S = "스마트폰"
        Case "smile": CompanyName_KRDictLoanwordKorean_S = "스마일"
        Case "smog": CompanyName_KRDictLoanwordKorean_S = "스모그"
        Case "snack": CompanyName_KRDictLoanwordKorean_S = "스낵"
        Case "snap": CompanyName_KRDictLoanwordKorean_S = "스냅"
        Case "snowboard": CompanyName_KRDictLoanwordKorean_S = "스노보드"
        Case "social": CompanyName_KRDictLoanwordKorean_S = "소셜"
        Case "socialmedia": CompanyName_KRDictLoanwordKorean_S = "소셜미디어"
        Case "socket": CompanyName_KRDictLoanwordKorean_S = "소켓"
        Case "sofa": CompanyName_KRDictLoanwordKorean_S = "소파"
        Case "softball": CompanyName_KRDictLoanwordKorean_S = "소프트볼"
        Case "software": CompanyName_KRDictLoanwordKorean_S = "소프트웨어"
        Case "sol": CompanyName_KRDictLoanwordKorean_S = "솔"
        Case "solo": CompanyName_KRDictLoanwordKorean_S = "솔로"
        Case "somalia": CompanyName_KRDictLoanwordKorean_S = "소말리아"
        Case "sonata": CompanyName_KRDictLoanwordKorean_S = "소나타"
        Case "soprano": CompanyName_KRDictLoanwordKorean_S = "소프라노"
        Case "sos": CompanyName_KRDictLoanwordKorean_S = "에스오에스"
        Case "sound": CompanyName_KRDictLoanwordKorean_S = "사운드"
        Case "soup": CompanyName_KRDictLoanwordKorean_S = "수프"
        Case "source": CompanyName_KRDictLoanwordKorean_S = "소스"
        Case "spa": CompanyName_KRDictLoanwordKorean_S = "스파"
        Case "spaghetti": CompanyName_KRDictLoanwordKorean_S = "스파게티"
        Case "spain": CompanyName_KRDictLoanwordKorean_S = "스페인"
        Case "spam": CompanyName_KRDictLoanwordKorean_S = "스팸"
        Case "spammail": CompanyName_KRDictLoanwordKorean_S = "스팸메일"
        Case "sparetire": CompanyName_KRDictLoanwordKorean_S = "스페어타이어"
        Case "speaker": CompanyName_KRDictLoanwordKorean_S = "스피커"
        Case "spectrum": CompanyName_KRDictLoanwordKorean_S = "스펙트럼"
        Case "speed": CompanyName_KRDictLoanwordKorean_S = "스피드"
        Case "spelling": CompanyName_KRDictLoanwordKorean_S = "스펠링"
        Case "sphinx": CompanyName_KRDictLoanwordKorean_S = "스핑크스"
        Case "spike": CompanyName_KRDictLoanwordKorean_S = "스파이크"
        Case "spin": CompanyName_KRDictLoanwordKorean_S = "스핀"
        Case "sponge": CompanyName_KRDictLoanwordKorean_S = "스펀지"
        Case "sponsor": CompanyName_KRDictLoanwordKorean_S = "스폰서"
        Case "spoon": CompanyName_KRDictLoanwordKorean_S = "스푼"
        Case "sports": CompanyName_KRDictLoanwordKorean_S = "스포츠"
        Case "sportscar": CompanyName_KRDictLoanwordKorean_S = "스포츠카"
        Case "sportscenter": CompanyName_KRDictLoanwordKorean_S = "스포츠센터"
        Case "sportsman": CompanyName_KRDictLoanwordKorean_S = "스포츠맨"
        Case "spotlight": CompanyName_KRDictLoanwordKorean_S = "스포트라이트"
        Case "spray": CompanyName_KRDictLoanwordKorean_S = "스프레이"
        Case "spring": CompanyName_KRDictLoanwordKorean_S = "스프링"
        Case "sprinkler": CompanyName_KRDictLoanwordKorean_S = "스프링클러"
        Case "spy": CompanyName_KRDictLoanwordKorean_S = "스파이"
        Case "srilanka": CompanyName_KRDictLoanwordKorean_S = "스리랑카"
        Case "stadium": CompanyName_KRDictLoanwordKorean_S = "스타디움"
        Case "staff": CompanyName_KRDictLoanwordKorean_S = "스태프"
        Case "stainless": CompanyName_KRDictLoanwordKorean_S = "스테인리스"
        Case "stamp": CompanyName_KRDictLoanwordKorean_S = "스탬프"
        Case "stand": CompanyName_KRDictLoanwordKorean_S = "스탠드"
        Case "stapler": CompanyName_KRDictLoanwordKorean_S = "스테이플러"
        Case "star": CompanyName_KRDictLoanwordKorean_S = "스타"
        Case "steadyseller": CompanyName_KRDictLoanwordKorean_S = "스테디셀러"
        Case "steak": CompanyName_KRDictLoanwordKorean_S = "스테이크"
        Case "steam": CompanyName_KRDictLoanwordKorean_S = "스팀"
        Case "step": CompanyName_KRDictLoanwordKorean_S = "스텝"
        Case "stereo": CompanyName_KRDictLoanwordKorean_S = "스테레오"
        Case "steroid": CompanyName_KRDictLoanwordKorean_S = "스테로이드"
        Case "steward": CompanyName_KRDictLoanwordKorean_S = "스튜어드"
        Case "stewardess": CompanyName_KRDictLoanwordKorean_S = "스튜어디스"
        Case "stick": CompanyName_KRDictLoanwordKorean_S = "스틱"
        Case "sticker": CompanyName_KRDictLoanwordKorean_S = "스티커"
        Case "stocking": CompanyName_KRDictLoanwordKorean_S = "스타킹"
        Case "storytelling": CompanyName_KRDictLoanwordKorean_S = "스토리텔링"
        Case "stove": CompanyName_KRDictLoanwordKorean_S = "스토브"
        Case "straight": CompanyName_KRDictLoanwordKorean_S = "스트레이트"
        Case "streaming": CompanyName_KRDictLoanwordKorean_S = "스트리밍"
        Case "stress": CompanyName_KRDictLoanwordKorean_S = "스트레스"
        Case "stretching": CompanyName_KRDictLoanwordKorean_S = "스트레칭"
        Case "strike": CompanyName_KRDictLoanwordKorean_S = "스트라이크"
        Case "striker": CompanyName_KRDictLoanwordKorean_S = "스트라이커"
        Case "studio": CompanyName_KRDictLoanwordKorean_S = "스튜디오"
        Case "study": CompanyName_KRDictLoanwordKorean_S = "스터디"
        Case "stuntman": CompanyName_KRDictLoanwordKorean_S = "스턴트맨"
        Case "style": CompanyName_KRDictLoanwordKorean_S = "스타일"
        Case "styrofoam": CompanyName_KRDictLoanwordKorean_S = "스티로폼"
        Case "sudan": CompanyName_KRDictLoanwordKorean_S = "수단"
        Case "sunglass": CompanyName_KRDictLoanwordKorean_S = "선글라스"
        Case "supercomputer": CompanyName_KRDictLoanwordKorean_S = "슈퍼컴퓨터"
        Case "superman": CompanyName_KRDictLoanwordKorean_S = "슈퍼맨"
        Case "supermarket": CompanyName_KRDictLoanwordKorean_S = "슈퍼마켓"
        Case "superwoman": CompanyName_KRDictLoanwordKorean_S = "슈퍼우먼"
        Case "supporters": CompanyName_KRDictLoanwordKorean_S = "서포터스"
        Case "surfing": CompanyName_KRDictLoanwordKorean_S = "서핑"
        Case "sweater": CompanyName_KRDictLoanwordKorean_S = "스웨터"
        Case "sweden": CompanyName_KRDictLoanwordKorean_S = "스웨덴"
        Case "swing": CompanyName_KRDictLoanwordKorean_S = "스윙"
        Case "switch": CompanyName_KRDictLoanwordKorean_S = "스위치"
        Case "symphony": CompanyName_KRDictLoanwordKorean_S = "심포니"
        Case "symposium": CompanyName_KRDictLoanwordKorean_S = "심포지엄"
        Case "syndrome": CompanyName_KRDictLoanwordKorean_S = "신드롬"
        Case "syria": CompanyName_KRDictLoanwordKorean_S = "시리아"
        Case "syrup": CompanyName_KRDictLoanwordKorean_S = "시럽"
        Case "system": CompanyName_KRDictLoanwordKorean_S = "시스템"
        Case Else: CompanyName_KRDictLoanwordKorean_S = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_T(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "table": CompanyName_KRDictLoanwordKorean_T = "테이블"
        Case "tablet": CompanyName_KRDictLoanwordKorean_T = "태블릿"
        Case "tabletpc": CompanyName_KRDictLoanwordKorean_T = "태블릿피시"
        Case "tabloid": CompanyName_KRDictLoanwordKorean_T = "타블로이드"
        Case "taboo": CompanyName_KRDictLoanwordKorean_T = "터부"
        Case "tackle": CompanyName_KRDictLoanwordKorean_T = "태클"
        Case "taco": CompanyName_KRDictLoanwordKorean_T = "타코"
        Case "tag": CompanyName_KRDictLoanwordKorean_T = "태그"
        Case "talk": CompanyName_KRDictLoanwordKorean_T = "토크"
        Case "talkshow": CompanyName_KRDictLoanwordKorean_T = "토크쇼"
        Case "tambourine": CompanyName_KRDictLoanwordKorean_T = "탬버린"
        Case "tango": CompanyName_KRDictLoanwordKorean_T = "탱고"
        Case "tank": CompanyName_KRDictLoanwordKorean_T = "탱크"
        Case "tanzania": CompanyName_KRDictLoanwordKorean_T = "탄자니아"
        Case "tape": CompanyName_KRDictLoanwordKorean_T = "테이프"
        Case "taxi": CompanyName_KRDictLoanwordKorean_T = "택시"
        Case "tea": CompanyName_KRDictLoanwordKorean_T = "티"
        Case "team": CompanyName_KRDictLoanwordKorean_T = "팀"
        Case "teamwork": CompanyName_KRDictLoanwordKorean_T = "팀워크"
        Case "teaspoon": CompanyName_KRDictLoanwordKorean_T = "티스푼"
        Case "telebanking": CompanyName_KRDictLoanwordKorean_T = "텔레뱅킹"
        Case "telepathy": CompanyName_KRDictLoanwordKorean_T = "텔레파시"
        Case "television": CompanyName_KRDictLoanwordKorean_T = "텔레비전"
        Case "tempo": CompanyName_KRDictLoanwordKorean_T = "템포"
        Case "tennis": CompanyName_KRDictLoanwordKorean_T = "테니스"
        Case "tenor": CompanyName_KRDictLoanwordKorean_T = "테너"
        Case "tent": CompanyName_KRDictLoanwordKorean_T = "텐트"
        Case "terminal": CompanyName_KRDictLoanwordKorean_T = "터미널"
        Case "terrace": CompanyName_KRDictLoanwordKorean_T = "테라스"
        Case "terror": CompanyName_KRDictLoanwordKorean_T = "테러"
        Case "terrorism": CompanyName_KRDictLoanwordKorean_T = "테러리즘"
        Case "terrorist": CompanyName_KRDictLoanwordKorean_T = "테러리스트"
        Case "test": CompanyName_KRDictLoanwordKorean_T = "테스트"
        Case "text": CompanyName_KRDictLoanwordKorean_T = "텍스트"
        Case "thinner": CompanyName_KRDictLoanwordKorean_T = "시너"
        Case "thrill": CompanyName_KRDictLoanwordKorean_T = "스릴"
        Case "thriller": CompanyName_KRDictLoanwordKorean_T = "스릴러"
        Case "tibet": CompanyName_KRDictLoanwordKorean_T = "티베트"
        Case "ticket": CompanyName_KRDictLoanwordKorean_T = "티켓"
        Case "tie": CompanyName_KRDictLoanwordKorean_T = "타이"
        Case "tights": CompanyName_KRDictLoanwordKorean_T = "타이츠"
        Case "tile": CompanyName_KRDictLoanwordKorean_T = "타일"
        Case "time": CompanyName_KRDictLoanwordKorean_T = "타임"
        Case "timecapsule": CompanyName_KRDictLoanwordKorean_T = "타임캡슐"
        Case "timemachine": CompanyName_KRDictLoanwordKorean_T = "타임머신"
        Case "timeout": CompanyName_KRDictLoanwordKorean_T = "타임아웃"
        Case "timing": CompanyName_KRDictLoanwordKorean_T = "타이밍"
        Case "timpani": CompanyName_KRDictLoanwordKorean_T = "팀파니"
        Case "tip": CompanyName_KRDictLoanwordKorean_T = "팁"
        Case "tire": CompanyName_KRDictLoanwordKorean_T = "타이어"
        Case "tissue": CompanyName_KRDictLoanwordKorean_T = "티슈"
        Case "title": CompanyName_KRDictLoanwordKorean_T = "타이틀"
        Case "toast": CompanyName_KRDictLoanwordKorean_T = "토스트"
        Case "toaster": CompanyName_KRDictLoanwordKorean_T = "토스터"
        Case "toefl": CompanyName_KRDictLoanwordKorean_T = "토플"
        Case "token": CompanyName_KRDictLoanwordKorean_T = "토큰"
        Case "tollgate": CompanyName_KRDictLoanwordKorean_T = "톨게이트"
        Case "tomato": CompanyName_KRDictLoanwordKorean_T = "토마토"
        Case "tomatoketchup": CompanyName_KRDictLoanwordKorean_T = "토마토케첩"
        Case "ton": CompanyName_KRDictLoanwordKorean_T = "톤"
        Case "tone": CompanyName_KRDictLoanwordKorean_T = "톤"
        Case "toner": CompanyName_KRDictLoanwordKorean_T = "토너"
        Case "top": CompanyName_KRDictLoanwordKorean_T = "톱"
        Case "topic": CompanyName_KRDictLoanwordKorean_T = "토픽"
        Case "topping": CompanyName_KRDictLoanwordKorean_T = "토핑"
        Case "toss": CompanyName_KRDictLoanwordKorean_T = "토스"
        Case "totem": CompanyName_KRDictLoanwordKorean_T = "토템"
        Case "totemism": CompanyName_KRDictLoanwordKorean_T = "토테미즘"
        Case "touch": CompanyName_KRDictLoanwordKorean_T = "터치"
        Case "touchscreen": CompanyName_KRDictLoanwordKorean_T = "터치스크린"
        Case "tournament": CompanyName_KRDictLoanwordKorean_T = "토너먼트"
        Case "towel": CompanyName_KRDictLoanwordKorean_T = "타월"
        Case "tower": CompanyName_KRDictLoanwordKorean_T = "타워"
        Case "track": CompanyName_KRDictLoanwordKorean_T = "트랙"
        Case "trade": CompanyName_KRDictLoanwordKorean_T = "트레이드"
        Case "trailer": CompanyName_KRDictLoanwordKorean_T = "트레일러"
        Case "trainer": CompanyName_KRDictLoanwordKorean_T = "트레이너"
        Case "training": CompanyName_KRDictLoanwordKorean_T = "트레이닝"
        Case "trap": CompanyName_KRDictLoanwordKorean_T = "트랩"
        Case "trenchcoat": CompanyName_KRDictLoanwordKorean_T = "트렌치코트"
        Case "triangle": CompanyName_KRDictLoanwordKorean_T = "트라이앵글"
        Case "trio": CompanyName_KRDictLoanwordKorean_T = "트리오"
        Case "troika": CompanyName_KRDictLoanwordKorean_T = "트로이카"
        Case "trombone": CompanyName_KRDictLoanwordKorean_T = "트롬본"
        Case "trophy": CompanyName_KRDictLoanwordKorean_T = "트로피"
        Case "truck": CompanyName_KRDictLoanwordKorean_T = "트럭"
        Case "trump": CompanyName_KRDictLoanwordKorean_T = "트럼프"
        Case "trumpet": CompanyName_KRDictLoanwordKorean_T = "트럼펫"
        Case "trunk": CompanyName_KRDictLoanwordKorean_T = "트렁크"
        Case "tsunami": CompanyName_KRDictLoanwordKorean_T = "쓰나미"
        Case "tuba": CompanyName_KRDictLoanwordKorean_T = "튜바"
        Case "tube": CompanyName_KRDictLoanwordKorean_T = "튜브"
        Case "tulip": CompanyName_KRDictLoanwordKorean_T = "튤립"
        Case "tumbler": CompanyName_KRDictLoanwordKorean_T = "텀블러"
        Case "tunnel": CompanyName_KRDictLoanwordKorean_T = "터널"
        Case "turkey": CompanyName_KRDictLoanwordKorean_T = "터키"
        Case "tuxedo": CompanyName_KRDictLoanwordKorean_T = "턱시도"
        Case "tv": CompanyName_KRDictLoanwordKorean_T = "티브이"
        Case "twopiece": CompanyName_KRDictLoanwordKorean_T = "투피스"
        Case "type": CompanyName_KRDictLoanwordKorean_T = "타입"
        Case "typing": CompanyName_KRDictLoanwordKorean_T = "타이핑"
        Case Else: CompanyName_KRDictLoanwordKorean_T = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_U(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "ufo": CompanyName_KRDictLoanwordKorean_U = "유에프오"
        Case "uganda": CompanyName_KRDictLoanwordKorean_U = "우간다"
        Case "ukraine": CompanyName_KRDictLoanwordKorean_U = "우크라이나"
        Case "un": CompanyName_KRDictLoanwordKorean_U = "유엔"
        Case "unesco": CompanyName_KRDictLoanwordKorean_U = "유네스코"
        Case "unicorn": CompanyName_KRDictLoanwordKorean_U = "유니콘"
        Case "uniform": CompanyName_KRDictLoanwordKorean_U = "유니폼"
        Case "universiade": CompanyName_KRDictLoanwordKorean_U = "유니버시아드"
        Case "update": CompanyName_KRDictLoanwordKorean_U = "업데이트"
        Case "upgrade": CompanyName_KRDictLoanwordKorean_U = "업그레이드"
        Case "upload": CompanyName_KRDictLoanwordKorean_U = "업로드"
        Case "uranium": CompanyName_KRDictLoanwordKorean_U = "우라늄"
        Case "uruguay": CompanyName_KRDictLoanwordKorean_U = "우루과이"
        Case "usb": CompanyName_KRDictLoanwordKorean_U = "유에스비"
        Case "utopia": CompanyName_KRDictLoanwordKorean_U = "유토피아"
        Case "uturn": CompanyName_KRDictLoanwordKorean_U = "유턴"
        Case "uzbekistan": CompanyName_KRDictLoanwordKorean_U = "우즈베키스탄"
        Case Else: CompanyName_KRDictLoanwordKorean_U = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_V(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "vaccine": CompanyName_KRDictLoanwordKorean_V = "백신"
        Case "valentineday": CompanyName_KRDictLoanwordKorean_V = "밸런타인데이"
        Case "valve": CompanyName_KRDictLoanwordKorean_V = "밸브"
        Case "vanilla": CompanyName_KRDictLoanwordKorean_V = "바닐라"
        Case "variety": CompanyName_KRDictLoanwordKorean_V = "버라이어티"
        Case "varietyshow": CompanyName_KRDictLoanwordKorean_V = "버라이어티쇼"
        Case "vaseline": CompanyName_KRDictLoanwordKorean_V = "바셀린"
        Case "veil": CompanyName_KRDictLoanwordKorean_V = "베일"
        Case "velvet": CompanyName_KRDictLoanwordKorean_V = "벨벳"
        Case "veneer": CompanyName_KRDictLoanwordKorean_V = "베니어"
        Case "venezuela": CompanyName_KRDictLoanwordKorean_V = "베네수엘라"
        Case "venture": CompanyName_KRDictLoanwordKorean_V = "벤처"
        Case "veranda": CompanyName_KRDictLoanwordKorean_V = "베란다"
        Case "version": CompanyName_KRDictLoanwordKorean_V = "버전"
        Case "video": CompanyName_KRDictLoanwordKorean_V = "비디오"
        Case "vietnam": CompanyName_KRDictLoanwordKorean_V = "베트남"
        Case "vinyl": CompanyName_KRDictLoanwordKorean_V = "비닐"
        Case "viola": CompanyName_KRDictLoanwordKorean_V = "비올라"
        Case "violin": CompanyName_KRDictLoanwordKorean_V = "바이올린"
        Case "violinist": CompanyName_KRDictLoanwordKorean_V = "바이올리니스트"
        Case "virus": CompanyName_KRDictLoanwordKorean_V = "바이러스"
        Case "visa": CompanyName_KRDictLoanwordKorean_V = "비자"
        Case "vision": CompanyName_KRDictLoanwordKorean_V = "비전"
        Case "visual": CompanyName_KRDictLoanwordKorean_V = "비주얼"
        Case "vitamin": CompanyName_KRDictLoanwordKorean_V = "비타민"
        Case "vocal": CompanyName_KRDictLoanwordKorean_V = "보컬"
        Case "vodka": CompanyName_KRDictLoanwordKorean_V = "보드카"
        Case "voice": CompanyName_KRDictLoanwordKorean_V = "보이스"
        Case "voicephishing": CompanyName_KRDictLoanwordKorean_V = "보이스피싱"
        Case "volt": CompanyName_KRDictLoanwordKorean_V = "볼트"
        Case "volume": CompanyName_KRDictLoanwordKorean_V = "볼륨"
        Case "voucher": CompanyName_KRDictLoanwordKorean_V = "바우처"
        Case Else: CompanyName_KRDictLoanwordKorean_V = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_W(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "waiter": CompanyName_KRDictLoanwordKorean_W = "웨이터"
        Case "waltz": CompanyName_KRDictLoanwordKorean_W = "왈츠"
        Case "wash": CompanyName_KRDictLoanwordKorean_W = "워시"
        Case "washington": CompanyName_KRDictLoanwordKorean_W = "워싱턴"
        Case "watt": CompanyName_KRDictLoanwordKorean_W = "와트"
        Case "wax": CompanyName_KRDictLoanwordKorean_W = "왁스"
        Case "web": CompanyName_KRDictLoanwordKorean_W = "웹"
        Case "webpage": CompanyName_KRDictLoanwordKorean_W = "웹페이지"
        Case "website": CompanyName_KRDictLoanwordKorean_W = "웹사이트"
        Case "webtoon": CompanyName_KRDictLoanwordKorean_W = "웹툰"
        Case "wedding": CompanyName_KRDictLoanwordKorean_W = "웨딩"
        Case "weddingalbum": CompanyName_KRDictLoanwordKorean_W = "웨딩앨범"
        Case "weddingdress": CompanyName_KRDictLoanwordKorean_W = "웨딩드레스"
        Case "weddingshop": CompanyName_KRDictLoanwordKorean_W = "웨딩숍"
        Case "wheelchair": CompanyName_KRDictLoanwordKorean_W = "휠체어"
        Case "whiskey": CompanyName_KRDictLoanwordKorean_W = "위스키"
        Case "whiteboard": CompanyName_KRDictLoanwordKorean_W = "화이트보드"
        Case "whitecollar": CompanyName_KRDictLoanwordKorean_W = "화이트칼라"
        Case "wifi": CompanyName_KRDictLoanwordKorean_W = "와이파이"
        Case "wine": CompanyName_KRDictLoanwordKorean_W = "와인"
        Case "wink": CompanyName_KRDictLoanwordKorean_W = "윙크"
        Case "wiper": CompanyName_KRDictLoanwordKorean_W = "와이퍼"
        Case "wire": CompanyName_KRDictLoanwordKorean_W = "와이어"
        Case "wit": CompanyName_KRDictLoanwordKorean_W = "위트"
        Case "wool": CompanyName_KRDictLoanwordKorean_W = "울"
        Case "working": CompanyName_KRDictLoanwordKorean_W = "워킹"
        Case "workingholiday": CompanyName_KRDictLoanwordKorean_W = "워킹홀리데이"
        Case "workshop": CompanyName_KRDictLoanwordKorean_W = "워크숍"
        Case "worldcup": CompanyName_KRDictLoanwordKorean_W = "월드컵"
        Case "wrestling": CompanyName_KRDictLoanwordKorean_W = "레슬링"
        Case Else: CompanyName_KRDictLoanwordKorean_W = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_X(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "xylophone": CompanyName_KRDictLoanwordKorean_X = "실로폰"
        Case Else: CompanyName_KRDictLoanwordKorean_X = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_Y(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "yacht": CompanyName_KRDictLoanwordKorean_Y = "요트"
        Case "yard": CompanyName_KRDictLoanwordKorean_Y = "야드"
        Case "yeast": CompanyName_KRDictLoanwordKorean_Y = "이스트"
        Case "yellowcard": CompanyName_KRDictLoanwordKorean_Y = "옐로카드"
        Case "yemen": CompanyName_KRDictLoanwordKorean_Y = "예멘"
        Case "yoga": CompanyName_KRDictLoanwordKorean_Y = "요가"
        Case "yogurt": CompanyName_KRDictLoanwordKorean_Y = "요구르트"
        Case "youth": CompanyName_KRDictLoanwordKorean_Y = "유스"
        Case "youthhostel": CompanyName_KRDictLoanwordKorean_Y = "유스호스텔"
        Case "yoyo": CompanyName_KRDictLoanwordKorean_Y = "요요"
        Case Else: CompanyName_KRDictLoanwordKorean_Y = wordText
    End Select
End Function

Private Function CompanyName_KRDictLoanwordKorean_Z(ByVal wordText As String) As String
    Select Case LCase$(wordText)
        Case "zambia": CompanyName_KRDictLoanwordKorean_Z = "잠비아"
        Case "zero": CompanyName_KRDictLoanwordKorean_Z = "제로"
        Case "zigzag": CompanyName_KRDictLoanwordKorean_Z = "지그재그"
        Case "zipper": CompanyName_KRDictLoanwordKorean_Z = "지퍼"
        Case Else: CompanyName_KRDictLoanwordKorean_Z = wordText
    End Select
End Function
' END KRDICT LOANWORD DICTIONARY

Private Function CompanyName_SpellAlphabetKorean(ByVal textValue As String) As String
    Dim result As String
    Dim i As Long
    For i = 1 To Len(textValue)
        Dim ch As String
        ch = Mid$(textValue, i, 1)
        If CompanyName_IsAsciiLetter(ch) Then
            result = result & CompanyName_AlphabetKorean(ch)
        Else
            result = result & ch
        End If
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
    Phone_ApplyLeadingZeroOption = "0" & digits
End Function

Private Function Phone_ShouldWriteReviewInstead(ByVal digits As String, ByVal options As Object) As Boolean
    If Not Phone_NormalizeOptionEnabled(options, "repeatRunAsReview") Then Exit Function
    Phone_ShouldWriteReviewInstead = Phone_HasRepeatedDigitRun(digits, 4)
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
Private Function DigitsOnly(ByVal source As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(source)
        ch = Mid$(source, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

Private Function Phone_DigitsForKoreanRules(ByVal source As String) As String
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
    If Toolkit_ShowCancelMessage("사업자등록 상태조회") Then Exit Sub
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
    If Toolkit_ShowCancelMessage("네이버 지역검색") Then Exit Sub
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
    If Toolkit_ShowCancelMessage("네이버 지역검색") Then Exit Sub
    MsgBox "네이버 주소 검증 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical, "네이버 지역검색"
End Sub

Public Sub NaverOfficialMall_FindLinksForSelection()
    Const PURPOSE As String = "네이버 공식몰 링크 검색"

    If Not NaverLocal_HasCredentials() Then
        MsgBox "먼저 [네이버 인증키 설정] 버튼으로 Client ID와 Client Secret을 저장하세요.", vbExclamation, PURPOSE
        Exit Sub
    End If

    Dim nameRange As Range
    If Not Toolkit_GetRangeFromUserForm(PURPOSE, _
        "공식몰 링크를 찾을 업체명/기업체명 범위를 선택하세요." & vbCrLf & _
        "선택 범위 오른쪽에 공식몰 후보 URL과 판정 등급을 추가합니다.", _
        nameRange) Then Exit Sub
    If nameRange Is Nothing Then Exit Sub

    Set nameRange = nameRange.Areas(1)
    If nameRange.Columns.Count > 1 Then Set nameRange = nameRange.Columns(1)

    Dim headerAnswer As VbMsgBoxResult
    headerAnswer = MsgBox("선택 범위의 첫 행을 헤더로 볼까요?", vbQuestion + vbYesNoCancel, PURPOSE)
    If headerAnswer = vbCancel Then Exit Sub
    Dim hasHeader As Boolean
    hasHeader = (headerAnswer = vbYes)

    Dim defaultDisplay As Long
    defaultDisplay = CLng(Val(GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_MALL_DISPLAY_KEY, "10")))
    If defaultDisplay < 1 Or defaultDisplay > 20 Then defaultDisplay = 10

    Dim displayInput As Variant
    displayInput = Application.InputBox("업체당 확인할 웹문서 검색 결과 수를 입력하세요." & vbCrLf & _
                                        "권장값: 10, 허용 범위: 1~20", _
                                        PURPOSE, defaultDisplay, Type:=1)
    If VarType(displayInput) = vbBoolean And displayInput = False Then Exit Sub

    Dim displayCount As Long
    displayCount = CLng(displayInput)
    If displayCount < 1 Or displayCount > 20 Then
        MsgBox "검색 결과 수는 1~20 사이로 입력하세요.", vbExclamation, PURPOSE
        Exit Sub
    End If

    Dim defaultDelay As Double
    defaultDelay = CDbl(Val(GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_MALL_DELAY_KEY, "0.3")))
    If defaultDelay < 0 Then defaultDelay = 0.3

    Dim delayInput As Variant
    delayInput = Application.InputBox("네이버 API 호출 사이 지연 시간을 초 단위로 입력하세요." & vbCrLf & _
                                      "권장값: 0.3", PURPOSE, defaultDelay, Type:=1)
    If VarType(delayInput) = vbBoolean And delayInput = False Then Exit Sub

    Dim delaySeconds As Double
    delaySeconds = CDbl(delayInput)
    If delaySeconds < 0 Then
        MsgBox "호출 지연 시간은 0 이상으로 입력하세요.", vbExclamation, PURPOSE
        Exit Sub
    End If

    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_MALL_DISPLAY_KEY, CStr(displayCount)
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_MALL_DELAY_KEY, CStr(delaySeconds)

    Dim firstDataRow As Long
    firstDataRow = nameRange.Row + IIf(hasHeader, 1, 0)

    Dim lastRow As Long
    lastRow = nameRange.Row + nameRange.Rows.Count - 1
    If firstDataRow > lastRow Then
        MsgBox "검색할 업체명 데이터가 없습니다.", vbInformation, PURPOSE
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = nameRange.Worksheet

    Dim outputCol As Long
    outputCol = nameRange.Column + nameRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outputCol, 9, PURPOSE) Then Exit Sub

    Dim totalRows As Long
    totalRows = lastRow - firstDataRow + 1
    If totalRows > 500 Then
        If MsgBox("검색 대상이 " & totalRows & "행입니다." & vbCrLf & _
                  "행마다 네이버 웹문서 API를 호출하므로 시간이 걸릴 수 있습니다. 계속할까요?", _
                  vbQuestion + vbYesNo, PURPOSE) <> vbYes Then Exit Sub
    End If

    Dim outputData() As Variant
    ReDim outputData(1 To totalRows, 1 To 9)

    Dim rowIndex As Long
    Dim outIndex As Long
    Dim companyName As String
    Dim queryText As String
    Dim items As Collection
    Dim bestItem As Object
    Dim gradeLabel As String
    Dim reasonText As String
    Dim reviewUrl As String
    Dim reviewReason As String
    Dim requestCount As Long
    Dim foundCount As Long
    Dim noMallCount As Long
    Dim errorCount As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outputCol).Resize(, 9).Insert Shift:=xlToRight
    If hasHeader Then NaverOfficialMall_WriteHeaders ws, nameRange.Row, outputCol

    Toolkit_ProgressShow PURPOSE, "검색 준비 중", totalRows

    For rowIndex = firstDataRow To lastRow
        outIndex = rowIndex - firstDataRow + 1
        companyName = Trim$(CStr(ws.Cells(rowIndex, nameRange.Column).Value))
        Toolkit_ProgressUpdate outIndex, totalRows, "공식몰 검색 중: " & CStr(rowIndex) & "행"

        If Len(companyName) > 0 Then
            requestCount = requestCount + 1
            queryText = companyName & " 공식몰"
            outputData(outIndex, 1) = queryText
            outputData(outIndex, 9) = Now

            Set items = Nothing
            Set bestItem = Nothing
            gradeLabel = ""
            reasonText = ""
            reviewUrl = ""
            reviewReason = ""

            On Error Resume Next
            Set items = NaverOfficialMall_SearchWeb(queryText, displayCount)
            If Err.Number <> 0 Then
                errorCount = errorCount + 1
                gradeLabel = "조회 실패"
                reasonText = Err.Description
                Toolkit_LogApiEvent ws.Parent, PURPOSE, "webkr", queryText, "ERROR", Err.Description
                Err.Clear
                On Error GoTo ErrHandler
                outputData(outIndex, 5) = gradeLabel
                outputData(outIndex, 6) = reasonText
                NaverLocal_Delay delaySeconds
                GoTo NextMallRow
            End If
            On Error GoTo ErrHandler

            Set bestItem = NaverOfficialMall_FindBestCandidate(items, companyName, gradeLabel, reasonText, reviewUrl, reviewReason)
            If bestItem Is Nothing Then
                noMallCount = noMallCount + 1
            Else
                foundCount = foundCount + 1
                outputData(outIndex, 2) = CStr(bestItem("link"))
                outputData(outIndex, 3) = CStr(bestItem("title"))
                outputData(outIndex, 4) = NaverOfficialMall_Domain(CStr(bestItem("link")))
            End If

            outputData(outIndex, 5) = gradeLabel
            outputData(outIndex, 6) = reasonText
            outputData(outIndex, 7) = reviewUrl
            outputData(outIndex, 8) = reviewReason
            NaverLocal_Delay delaySeconds
        End If
NextMallRow:
    Next rowIndex

    ws.Range(ws.Cells(firstDataRow, outputCol), ws.Cells(lastRow, outputCol + 8)).Value = outputData
    If hasHeader Then ws.Range(ws.Cells(nameRange.Row, outputCol), ws.Cells(nameRange.Row, outputCol + 8)).Font.Bold = True

    For rowIndex = firstDataRow To lastRow
        NaverOfficialMall_ColorGradeCell ws.Cells(rowIndex, outputCol + 4)
    Next rowIndex

    ws.Columns(outputCol).Resize(, 9).AutoFit
    Toolkit_ProgressClose
    Application.ScreenUpdating = True

    MsgBox "네이버 공식몰 링크 검색을 완료했습니다." & vbCrLf & _
           "조회: " & requestCount & "건" & vbCrLf & _
           "자체 공식몰 후보: " & foundCount & "건" & vbCrLf & _
           "공식몰 없음/확인 필요: " & noMallCount & "건" & vbCrLf & _
           "조회 실패: " & errorCount & "건", vbInformation, PURPOSE
    Exit Sub

ErrHandler:
    Toolkit_ProgressClose
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage(PURPOSE) Then Exit Sub
    MsgBox "네이버 공식몰 링크 검색 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical, PURPOSE
End Sub

Public Function NaverOfficialMall_ClassifyUrlForTest(ByVal urlValue As String, _
                                                     ByVal titleText As String, _
                                                     ByVal descriptionText As String, _
                                                     ByVal companyName As String) As String
    NaverOfficialMall_ClassifyUrlForTest = NaverOfficialMall_ClassifyUrl(urlValue, titleText, descriptionText, companyName)
End Function

Private Sub NaverOfficialMall_WriteHeaders(ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long)
    Dim headers As Variant
    headers = Array("공식몰 검색어", "공식몰 후보 URL", "후보 제목", "후보 도메인", "판정등급", "판정근거", "제외/검토 URL", "제외/검토 사유", "검색일시")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowIndex, colIndex + i).Value = headers(i)
    Next i
End Sub

Private Function NaverOfficialMall_SearchWeb(ByVal queryText As String, ByVal displayCount As Long) As Collection
    If displayCount < 1 Then displayCount = 1
    If displayCount > 20 Then displayCount = 20

    Static searchCache As Object
    If searchCache Is Nothing Then Set searchCache = CreateObject("Scripting.Dictionary")

    Dim cacheKey As String
    cacheKey = NaverLocal_GetClientId() & "|webkr|" & queryText & "|" & CStr(displayCount)
    If searchCache.Exists(cacheKey) Then
        Set NaverOfficialMall_SearchWeb = searchCache(cacheKey)
        Exit Function
    End If

    Dim url As String
    url = NAVER_WEB_API_BASE_URL & "?query=" & NaverLocal_UrlEncode(queryText) & _
          "&display=" & CStr(displayCount) & "&start=1"

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts API_RESOLVE_TIMEOUT_MS, API_CONNECT_TIMEOUT_MS, API_SEND_TIMEOUT_MS, API_RECEIVE_TIMEOUT_MS
    http.Open "GET", url, False
    http.setRequestHeader "X-Naver-Client-Id", NaverLocal_GetClientId()
    http.setRequestHeader "X-Naver-Client-Secret", NaverLocal_GetClientSecret()
    http.setRequestHeader "Accept", "application/json"
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 782, "NaverOfficialMall_SearchWeb", "HTTP " & http.Status & ": " & http.responseText
    End If

    Set NaverOfficialMall_SearchWeb = NaverOfficialMall_ExtractWebItems(CStr(http.responseText))
    searchCache.Add cacheKey, NaverOfficialMall_SearchWeb
End Function

Private Function NaverOfficialMall_ExtractWebItems(ByVal jsonText As String) As Collection
    Dim result As New Collection
    Dim itemObjects As Collection
    Set itemObjects = NaverLocal_ExtractArrayObjects(jsonText, "items")

    Dim itemJson As Variant
    Dim item As Object
    For Each itemJson In itemObjects
        Set item = CreateObject("Scripting.Dictionary")
        item("title") = NaverLocal_StripHtml(Biz_JsonString(CStr(itemJson), "title"))
        item("link") = Biz_JsonString(CStr(itemJson), "link")
        item("description") = NaverLocal_StripHtml(Biz_JsonString(CStr(itemJson), "description"))
        result.Add item
    Next itemJson

    Set NaverOfficialMall_ExtractWebItems = result
End Function

Private Function NaverOfficialMall_FindBestCandidate(ByVal items As Collection, ByVal companyName As String, _
                                                     ByRef gradeLabel As String, ByRef reasonText As String, _
                                                     ByRef reviewUrl As String, ByRef reviewReason As String) As Object
    gradeLabel = "공식몰 없음"
    reasonText = "검색 결과에서 자체 쇼핑몰 후보를 찾지 못했습니다."

    If items Is Nothing Then
        reasonText = "검색 결과가 없습니다."
        Exit Function
    End If
    If items.Count = 0 Then
        reasonText = "검색 결과가 없습니다."
        Exit Function
    End If

    Dim bestScore As Double
    bestScore = -1
    Dim bestReason As String
    Dim bestItem As Object
    Dim firstUnknownUrl As String
    Dim firstUnknownReason As String
    Dim firstExcludedUrl As String
    Dim firstExcludedReason As String

    Dim item As Object
    For Each item In items
        Dim classCode As String
        classCode = NaverOfficialMall_ClassifyUrl(CStr(item("link")), CStr(item("title")), CStr(item("description")), companyName)

        If classCode = "candidate:self_mall" Then
            Dim candidateScore As Double
            Dim candidateReason As String
            candidateScore = NaverOfficialMall_CandidateScore(item, companyName, candidateReason)
            If candidateScore > bestScore Then
                bestScore = candidateScore
                bestReason = candidateReason
                Set bestItem = item
            End If
        ElseIf classCode = "candidate:unknown_site" Then
            If Len(firstUnknownUrl) = 0 Then
                firstUnknownUrl = CStr(item("link"))
                firstUnknownReason = NaverOfficialMall_ReasonForCode(classCode)
            End If
        Else
            If Len(firstExcludedUrl) = 0 Then
                firstExcludedUrl = CStr(item("link"))
                firstExcludedReason = NaverOfficialMall_ReasonForCode(classCode)
            End If
        End If
    Next item

    If Not bestItem Is Nothing Then
        If bestScore >= 80 Then
            gradeLabel = "정상"
        ElseIf bestScore >= 60 Then
            gradeLabel = "가능성 높음"
        Else
            gradeLabel = "확인 필요"
        End If
        reasonText = bestReason
        Set NaverOfficialMall_FindBestCandidate = bestItem
    ElseIf Len(firstUnknownUrl) > 0 Then
        gradeLabel = "확인 필요"
        reasonText = "자체 도메인 후보는 있으나 쇼핑 기능 단서가 부족합니다."
        reviewUrl = firstUnknownUrl
        reviewReason = firstUnknownReason
    ElseIf Len(firstExcludedUrl) > 0 Then
        gradeLabel = "공식몰 없음"
        reasonText = "자체몰 후보 없이 제외 대상 결과만 확인되었습니다."
        reviewUrl = firstExcludedUrl
        reviewReason = firstExcludedReason
    End If
End Function

Private Function NaverOfficialMall_ClassifyUrl(ByVal urlValue As String, ByVal titleText As String, _
                                               ByVal descriptionText As String, ByVal companyName As String) As String
    Dim domain As String
    domain = NaverOfficialMall_Domain(urlValue)
    If Len(domain) = 0 Then
        NaverOfficialMall_ClassifyUrl = "no_url"
        Exit Function
    End If

    If NaverOfficialMall_DomainMatches(domain, "brand.naver.com") Then
        NaverOfficialMall_ClassifyUrl = "exclude:naver_brand_store"
    ElseIf NaverOfficialMall_DomainMatches(domain, "smartstore.naver.com") Then
        NaverOfficialMall_ClassifyUrl = "exclude:naver_smartstore"
    ElseIf NaverOfficialMall_IsOpenMarketDomain(domain) Then
        NaverOfficialMall_ClassifyUrl = "exclude:open_market"
    ElseIf NaverOfficialMall_IsSocialOrBlogDomain(domain) Then
        NaverOfficialMall_ClassifyUrl = "exclude:social_or_blog"
    ElseIf NaverOfficialMall_HasMallEvidence(urlValue, titleText, descriptionText, domain) Then
        NaverOfficialMall_ClassifyUrl = "candidate:self_mall"
    Else
        NaverOfficialMall_ClassifyUrl = "candidate:unknown_site"
    End If
End Function

Private Function NaverOfficialMall_Domain(ByVal urlValue As String) As String
    Dim textValue As String
    textValue = LCase$(Trim$(CStr(urlValue)))
    If Len(textValue) = 0 Then Exit Function

    Dim protocolPos As Long
    protocolPos = InStr(1, textValue, "://", vbTextCompare)
    If protocolPos > 0 Then textValue = Mid$(textValue, protocolPos + 3)

    Dim slashPos As Long
    slashPos = InStr(1, textValue, "/", vbTextCompare)
    If slashPos > 0 Then textValue = Left$(textValue, slashPos - 1)

    Dim questionPos As Long
    questionPos = InStr(1, textValue, "?", vbTextCompare)
    If questionPos > 0 Then textValue = Left$(textValue, questionPos - 1)

    Dim portPos As Long
    portPos = InStr(1, textValue, ":", vbTextCompare)
    If portPos > 0 Then textValue = Left$(textValue, portPos - 1)

    If Left$(textValue, 4) = "www." Then textValue = Mid$(textValue, 5)
    NaverOfficialMall_Domain = textValue
End Function

Private Function NaverOfficialMall_DomainMatches(ByVal domain As String, ByVal matchDomain As String) As Boolean
    domain = LCase$(Trim$(domain))
    matchDomain = LCase$(Trim$(matchDomain))
    NaverOfficialMall_DomainMatches = (domain = matchDomain Or Right$(domain, Len(matchDomain) + 1) = "." & matchDomain)
End Function

Private Function NaverOfficialMall_IsOpenMarketDomain(ByVal domain As String) As Boolean
    Dim domains As Variant
    domains = Array("shopping.naver.com", "search.shopping.naver.com", "coupang.com", "11st.co.kr", _
                    "gmarket.co.kr", "auction.co.kr", "ssg.com", "lotteon.com", "tmon.co.kr", _
                    "wemakeprice.com", "interpark.com", "danawa.com", "musinsa.com", "zigzag.kr", _
                    "ably.co.kr", "kakaostyle.com", "store.kakao.com", "amazon.com", "aliexpress.com", _
                    "qoo10.com", "wemakeprice.co.kr")

    Dim item As Variant
    For Each item In domains
        If NaverOfficialMall_DomainMatches(domain, CStr(item)) Then
            NaverOfficialMall_IsOpenMarketDomain = True
            Exit Function
        End If
    Next item
End Function

Private Function NaverOfficialMall_IsSocialOrBlogDomain(ByVal domain As String) As Boolean
    Dim domains As Variant
    domains = Array("blog.naver.com", "cafe.naver.com", "post.naver.com", "tistory.com", "brunch.co.kr", _
                    "instagram.com", "facebook.com", "youtube.com", "youtu.be", "twitter.com", "x.com", _
                    "linkedin.com", "threads.net")

    Dim item As Variant
    For Each item In domains
        If NaverOfficialMall_DomainMatches(domain, CStr(item)) Then
            NaverOfficialMall_IsSocialOrBlogDomain = True
            Exit Function
        End If
    Next item
End Function

Private Function NaverOfficialMall_HasMallEvidence(ByVal urlValue As String, ByVal titleText As String, _
                                                   ByVal descriptionText As String, ByVal domain As String) As Boolean
    Dim combined As String
    combined = LCase$(CStr(urlValue) & " " & CStr(titleText) & " " & CStr(descriptionText))

    If Left$(domain, 5) = "shop." Or Left$(domain, 5) = "mall." Or Left$(domain, 6) = "store." Then
        NaverOfficialMall_HasMallEvidence = True
        Exit Function
    End If

    Dim keywords As Variant
    keywords = Array("공식몰", "공식 온라인몰", "공식 쇼핑몰", "자사몰", "온라인몰", "쇼핑몰", _
                     "브랜드몰", "구매하기", "장바구니", "결제", "상품상세", "product", _
                     "products", "goods", "shop", "mall", "store", "cart", "checkout", _
                     "official mall", "official store", "online store")

    Dim keyword As Variant
    For Each keyword In keywords
        If InStr(1, combined, CStr(keyword), vbTextCompare) > 0 Then
            NaverOfficialMall_HasMallEvidence = True
            Exit Function
        End If
    Next keyword
End Function

Private Function NaverOfficialMall_CandidateScore(ByVal item As Object, ByVal companyName As String, ByRef reasonText As String) As Double
    Dim score As Double
    score = 50

    Dim titleText As String
    Dim descriptionText As String
    Dim linkText As String
    titleText = CStr(item("title"))
    descriptionText = CStr(item("description"))
    linkText = CStr(item("link"))

    Dim nameScore As Double
    nameScore = NaverLocal_Similarity(companyName, titleText)
    score = score + (nameScore * 0.25)

    Dim domain As String
    domain = NaverOfficialMall_Domain(linkText)
    Dim normalizedCompany As String
    normalizedCompany = NaverLocal_NormalizeForMatch(companyName)
    If Len(normalizedCompany) >= 3 Then
        If InStr(1, NaverLocal_NormalizeForMatch(domain), normalizedCompany, vbTextCompare) > 0 Then score = score + 15
    End If

    Dim combined As String
    combined = LCase$(linkText & " " & titleText & " " & descriptionText)
    If InStr(1, combined, "공식", vbTextCompare) > 0 Or InStr(1, combined, "official", vbTextCompare) > 0 Then score = score + 10
    If InStr(1, combined, "장바구니", vbTextCompare) > 0 Or InStr(1, combined, "cart", vbTextCompare) > 0 Then score = score + 8
    If InStr(1, combined, "결제", vbTextCompare) > 0 Or InStr(1, combined, "checkout", vbTextCompare) > 0 Then score = score + 8

    If score > 100 Then score = 100
    NaverOfficialMall_CandidateScore = score

    reasonText = "자체몰 후보. 업체명 유사도 " & Format$(nameScore, "0.0") & _
                 ", 판정점수 " & Format$(score, "0.0") & _
                 ". 네이버 브랜드스토어/스마트스토어/오픈마켓은 제외했습니다."
End Function

Private Function NaverOfficialMall_ReasonForCode(ByVal classCode As String) As String
    Select Case classCode
        Case "exclude:naver_brand_store"
            NaverOfficialMall_ReasonForCode = "네이버 브랜드스토어 제외: 자체 쇼핑몰로 인정하지 않음"
        Case "exclude:naver_smartstore"
            NaverOfficialMall_ReasonForCode = "네이버 스마트스토어 제외: 자체 쇼핑몰로 인정하지 않음"
        Case "exclude:open_market"
            NaverOfficialMall_ReasonForCode = "오픈마켓/입점몰 제외"
        Case "exclude:social_or_blog"
            NaverOfficialMall_ReasonForCode = "블로그/카페/SNS/영상 채널 제외"
        Case "candidate:unknown_site"
            NaverOfficialMall_ReasonForCode = "자체 도메인 가능성은 있으나 쇼핑 기능 단서가 부족함"
        Case "no_url"
            NaverOfficialMall_ReasonForCode = "URL 없음"
        Case Else
            NaverOfficialMall_ReasonForCode = classCode
    End Select
End Function

Private Sub NaverOfficialMall_ColorGradeCell(ByVal cell As Range)
    Select Case CStr(cell.Value)
        Case "정상"
            cell.Interior.Color = RGB(198, 239, 206)
        Case "가능성 높음"
            cell.Interior.Color = RGB(255, 235, 156)
        Case "확인 필요"
            cell.Interior.Color = RGB(244, 176, 132)
        Case Else
            cell.Interior.Color = RGB(255, 199, 206)
    End Select
End Sub

Private Function NaverLocal_GetBuildOptions() As Object
    Dim defaultIndustry As String
    Dim defaultDisplayCount As Long
    Dim defaultDelaySeconds As Double
    Dim defaultSortMode As String
    Dim defaultDongBasis As String
    Dim defaultSido As String
    Dim defaultSigungu As String

    defaultIndustry = GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_INDUSTRY_KEY, "서점")
    defaultDisplayCount = CLng(Val(GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_DISPLAY_KEY, "5")))
    If defaultDisplayCount < 1 Or defaultDisplayCount > 5 Then defaultDisplayCount = 5
    defaultDelaySeconds = CDbl(Val(GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_DELAY_KEY, "0.4")))
    If defaultDelaySeconds < 0 Then defaultDelaySeconds = 0.4
    defaultSortMode = LCase$(Trim$(GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_SORT_KEY, "comment")))
    If defaultSortMode <> "random" Then defaultSortMode = "comment"
    defaultDongBasis = LCase$(Trim$(GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_BASIS_KEY, "admin")))
    If defaultDongBasis <> "legal" Then defaultDongBasis = "admin"
    defaultSido = GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_SIDO_KEY, "")
    defaultSigungu = GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_SIGUNGU_KEY, "")

    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmNaverLocalBuildOptions")
    form.Configure defaultIndustry, defaultDisplayCount, defaultDelaySeconds, defaultSortMode, defaultDongBasis, defaultSido, defaultSigungu, NaverLocal_GetRegionOptionsText(ActiveWorkbook)
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
    NaverLocal_SaveBuildOptions options
    Set NaverLocal_GetBuildOptions = options
    Exit Function

Fallback:
    Dim industry As String
    industry = InputBox("검색할 업종/키워드를 입력하세요. 예: 서점, 카페, 음식점", "네이버 신규 리스트", defaultIndustry)
    If Len(Trim$(industry)) = 0 Then Exit Function
    Set options = CreateObject("Scripting.Dictionary")
    options("industry") = Trim$(industry)
    options("displayCount") = defaultDisplayCount
    options("delaySeconds") = defaultDelaySeconds
    options("sortMode") = defaultSortMode
    options("sidoFilter") = defaultSido
    options("sigunguFilter") = defaultSigungu
    options("dongBasis") = defaultDongBasis
    NaverLocal_SaveBuildOptions options
    Set NaverLocal_GetBuildOptions = options
End Function

Private Sub NaverLocal_SaveBuildOptions(ByVal options As Object)
    On Error Resume Next
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_INDUSTRY_KEY, CStr(options("industry"))
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_DISPLAY_KEY, CStr(options("displayCount"))
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_DELAY_KEY, CStr(options("delaySeconds"))
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_SORT_KEY, CStr(options("sortMode"))
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_BASIS_KEY, CStr(options("dongBasis"))
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_SIDO_KEY, CStr(options("sidoFilter"))
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUILD_SIGUNGU_KEY, CStr(options("sigunguFilter"))
    On Error GoTo 0
End Sub

Private Function NaverLocal_GetRegionOptionsText(ByVal wb As Workbook) As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim result As String
    NaverLocal_AddRegionOptions result, seen, "admin", NaverLocal_GetMasterSheetForBasis(wb, "admin")
    NaverLocal_AddRegionOptions result, seen, "legal", NaverLocal_GetMasterSheetForBasis(wb, "legal")
    NaverLocal_GetRegionOptionsText = result
End Function

Private Sub NaverLocal_AddRegionOptions(ByRef outputText As String, ByVal seen As Object, ByVal basis As String, ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long: headerRow = 1
    Dim sidoCol As Long, sigunguCol As Long
    sidoCol = NaverLocal_FindHeaderColumn(ws, headerRow, "광역시도")
    If sidoCol = 0 Then sidoCol = NaverLocal_FindHeaderColumn(ws, headerRow, "시도명")
    sigunguCol = NaverLocal_FindHeaderColumn(ws, headerRow, "시군구")
    If sigunguCol = 0 Then sigunguCol = NaverLocal_FindHeaderColumn(ws, headerRow, "시군구명")
    If sidoCol = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, sidoCol).End(xlUp).Row

    Dim r As Long
    For r = headerRow + 1 To lastRow
        Dim sido As String
        Dim sigungu As String
        sido = Trim$(CStr(ws.Cells(r, sidoCol).Value))
        If sigunguCol > 0 Then sigungu = Trim$(CStr(ws.Cells(r, sigunguCol).Value)) Else sigungu = ""
        If Len(sido) > 0 Then
            Dim key As String
            key = basis & "|" & sido & "|" & sigungu
            If Not seen.Exists(key) Then
                seen.Add key, True
                If Len(outputText) > 0 Then outputText = outputText & vbLf
                outputText = outputText & basis & vbTab & sido & vbTab & sigungu
            End If
        End If
    Next r
End Sub

Private Function NaverLocal_GetMasterSheetForBasis(ByVal wb As Workbook, ByVal basis As String) As Worksheet
    If LCase$(Trim$(basis)) = "legal" Then
        Set NaverLocal_GetMasterSheetForBasis = NaverLocal_GetUserLegalDongMasterSheet(wb)
        If NaverLocal_GetMasterSheetForBasis Is Nothing Then Set NaverLocal_GetMasterSheetForBasis = NaverLocal_GetEmbeddedLegalDongMasterSheet()
    Else
        Set NaverLocal_GetMasterSheetForBasis = NaverLocal_GetUserDongMasterSheet(wb)
        If NaverLocal_GetMasterSheetForBasis Is Nothing Then Set NaverLocal_GetMasterSheetForBasis = NaverLocal_GetEmbeddedDongMasterSheet()
    End If
End Function

Private Function NaverLocal_GetDongNamesForSearch(ByVal wb As Workbook, ByVal options As Object) As Collection
    Dim result As New Collection
    Dim ws As Worksheet
    Set ws = NaverLocal_GetMasterSheetForBasis(wb, NaverLocal_DongBasis(options))
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

    Dim queryMode As String
    Dim displayCount As Long
    Dim delaySeconds As Double
    Dim sortMode As String
    Dim hasHeader As Boolean
    Dim useName As Boolean
    Dim useAddress As Boolean
    Dim nameRangeAddress As String
    Dim addressRangeAddress As String
    Dim nameRange As Range
    Dim addressRange As Range

    queryMode = CStr(form.QueryMode)
    displayCount = CLng(form.DisplayCount)
    delaySeconds = CDbl(form.DelaySeconds)
    sortMode = CStr(form.SortMode)
    hasHeader = CBool(form.HasHeader)
    useName = CBool(form.UseName)
    useAddress = CBool(form.UseAddress)
    nameRangeAddress = CStr(form.NameRangeAddress)
    addressRangeAddress = CStr(form.AddressRangeAddress)
    If useName Then Set nameRange = form.NameRange
    If useAddress Then Set addressRange = form.AddressRange
    Unload form

    If useName And nameRange Is Nothing And Len(Trim$(nameRangeAddress)) > 0 Then
        Set nameRange = NaverLocal_RangeFromAddressText(nameRangeAddress)
        If nameRange Is Nothing Then
            MsgBox "입력한 업체명 범위 주소를 해석하지 못했습니다." & vbCrLf & _
                   "다음 범위 선택창에서 다시 지정하세요.", vbInformation, "네이버 지역검색"
        End If
    End If
    If useAddress And addressRange Is Nothing And Len(Trim$(addressRangeAddress)) > 0 Then
        Set addressRange = NaverLocal_RangeFromAddressText(addressRangeAddress)
        If addressRange Is Nothing Then
            MsgBox "입력한 주소 범위 주소를 해석하지 못했습니다." & vbCrLf & _
                   "다음 범위 선택창에서 다시 지정하세요.", vbInformation, "네이버 지역검색"
        End If
    End If

    If useName And nameRange Is Nothing Then
        If Not Toolkit_GetRangeFromUserForm("네이버 주소 검증 - 업체명 범위", _
            "업체명 또는 기업체명 범위를 선택하세요.", nameRange) Then Exit Function
    End If
    If useAddress And addressRange Is Nothing Then
        If Not Toolkit_GetRangeFromUserForm("네이버 주소 검증 - 주소 범위", _
            "주소 범위를 선택하세요.", addressRange) Then Exit Function
    End If

    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    options("queryMode") = queryMode
    options("displayCount") = displayCount
    options("delaySeconds") = delaySeconds
    options("sortMode") = sortMode
    options("hasHeader") = hasHeader
    options("useName") = useName
    options("useAddress") = useAddress
    If useName Then Set options("nameRange") = nameRange
    If useAddress Then Set options("addressRange") = addressRange
    Set NaverLocal_GetVerifyOptions = options
    Exit Function

Fallback:
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    On Error GoTo 0
    MsgBox "주소 검증 옵션창을 열 수 없습니다." & vbCrLf & Err.Description, vbCritical, "네이버 지역검색"
End Function

Private Function NaverLocal_RangeFromAddressText(ByVal addressText As String) As Range
    Dim cleaned As String
    cleaned = Trim$(addressText)
    If Len(cleaned) = 0 Then Exit Function

    On Error Resume Next
    Set NaverLocal_RangeFromAddressText = Application.Range(cleaned)
    On Error GoTo 0
End Function

Private Function NaverLocal_VerifyPromptList(ByVal useName As Boolean, ByVal useAddress As Boolean, _
                                             ByVal nameRangeAddress As String, ByVal addressRangeAddress As String) As String
    If useName And Len(Trim$(nameRangeAddress)) = 0 Then NaverLocal_VerifyPromptList = "name"
    If useAddress And Len(Trim$(addressRangeAddress)) = 0 Then
        If Len(NaverLocal_VerifyPromptList) > 0 Then NaverLocal_VerifyPromptList = NaverLocal_VerifyPromptList & ","
        NaverLocal_VerifyPromptList = NaverLocal_VerifyPromptList & "address"
    End If
End Function

Public Function NaverLocal_VerifyPromptListForTest(ByVal useName As Boolean, ByVal useAddress As Boolean, _
                                                   ByVal nameRangeAddress As String, ByVal addressRangeAddress As String) As String
    NaverLocal_VerifyPromptListForTest = NaverLocal_VerifyPromptList(useName, useAddress, nameRangeAddress, addressRangeAddress)
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
        Toolkit_RaiseIfCancelled
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
    If phoneCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 오류", errorPhone, "000-0000, 0000-0000 등 명백한 임시값/영값 패턴"): outRow = outRow + 1
    If phoneCol > 0 Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 확인 필요", invalidPhone, "숫자 길이/대표번호 패턴/4회 이상 반복 숫자 기준"): outRow = outRow + 1
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
    If selectedItems.Exists("3") Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 오류", errorPhone, "000-0000, 0000-0000 등 명백한 임시값/영값 패턴"): outRow = outRow + 1
    If selectedItems.Exists("3") Then Roster_WriteArray outWs, outRow, 1, Array("전화번호 확인 필요", invalidPhone, "숫자 길이/대표번호 패턴/4회 이상 반복 숫자 기준"): outRow = outRow + 1
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
    Dim digits As String: digits = Phone_DigitsForKoreanRules(phoneText)
    Dim phoneType As String
    Dim statusText As String

    If Len(digits) = 0 Then
        phoneType = "빈값"
        statusText = "확인 필요"
    ElseIf Phone_IsPlaceholderZeroPattern(digits) Then
        phoneType = "오류 패턴"
        statusText = "오류"
    ElseIf Phone_IsSequentialPlaceholderDigits(digits) Then
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

    If statusText = "정상" And Phone_HasRepeatedDigitRun(digits, 4) Then statusText = "확인 필요"
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
Private Const SD_BAND_SETTINGS_SHEET As String = "_SampleDesignBandSettings"

Public Sub SampleDesign_CreateTemplate()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "표본설계 시트를 만들 통합문서를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim setup As Object
    Set setup = SampleDesign_GetSetupFromUser()
    If setup Is Nothing Then Exit Sub

    Dim templateMode As String
    templateMode = CStr(setup("inputMode"))
    If Len(templateMode) = 0 Then Exit Sub

    Dim guideWs As Worksheet
    Dim inputWs As Worksheet
    Dim tableWs As Worksheet
    Dim resultWs As Worksheet
    Dim reviewWs As Worksheet
    Set guideWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_가이드")
    Set inputWs = SampleDesign_PrepareSheet(wb, "표본설계_입력", "표본설계 입력 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?")
    If inputWs Is Nothing Then Exit Sub
    If templateMode = "table" Then
        Set tableWs = SampleDesign_PrepareSheet(wb, "표본설계_집계표입력", "표본설계 집계표입력 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?")
        If tableWs Is Nothing Then Exit Sub
    End If
    Set resultWs = SampleDesign_PrepareSheet(wb, "표본설계_할당결과", "표본설계 할당결과 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?")
    If resultWs Is Nothing Then Exit Sub
    Set reviewWs = SampleDesign_PrepareSheet(wb, "표본설계_검토", "표본설계 검토 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?")
    If reviewWs Is Nothing Then Exit Sub

    SampleDesign_WriteGuideSheet guideWs, setup, templateMode
    SampleDesign_WriteInputTemplate inputWs, setup
    If templateMode = "table" Then SampleDesign_WriteTableInputTemplate tableWs
    SampleDesign_WriteEmptyResult resultWs
    SampleDesign_WriteEmptyReview reviewWs

    If templateMode = "table" Then
        tableWs.Activate
        tableWs.Range("A8").Select
        MsgBox "집계표 입력용 시트와 계산용 리스트 입력 시트를 만들었습니다." & vbCrLf & _
               "1) 표본설계_가이드 시트를 먼저 확인하세요." & vbCrLf & _
               "2) 집계표를 표본설계_집계표입력 시트에 붙여넣은 뒤 [할당 계산]을 실행하세요.", _
               vbInformation, "표본설계"
    Else
        inputWs.Activate
        inputWs.Range("A1").Select
        MsgBox "표본설계 입력/결과/검토 시트를 만들었습니다." & vbCrLf & _
               "표본설계_가이드 시트의 체크리스트를 확인한 뒤 모집단 표를 채우고 [할당 계산]을 실행하세요.", _
               vbInformation, "표본설계"
    End If
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

    Dim minN As Long
    Dim maxN As Long
    Dim alpha As Double
    Dim methodName As String
    SampleDesign_ReadRunSettings inputWs, methodName, alpha, minN, maxN

    If Not SampleDesign_SelectMethod(methodName, alpha) Then Exit Sub
    inputWs.Range("B5").Value = methodName
    inputWs.Range("B6").Value = alpha

    Dim issues As Collection
    Set issues = New Collection

    If SampleDesign_PopulationDataRowCount(inputWs) = 0 Then
        Dim tableWs As Worksheet
        Set tableWs = SampleDesign_FindSheet(wb, "표본설계_집계표입력")
        If Not tableWs Is Nothing Then SampleDesign_ImportTableInputToPopulation inputWs, tableWs, issues
    End If

    Dim schema As Object
    Set schema = SampleDesign_ReadInputSchema(inputWs)
    If schema Is Nothing Then
        MsgBox "표본설계_입력 시트에서 '모집단수' 헤더를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim defaultTargetN As Long
    defaultTargetN = CLng(Val(inputWs.Range("B4").Value))

    Dim targetGroups As Collection
    Set targetGroups = SampleDesign_ReadTargetGroups(inputWs, schema, defaultTargetN, issues)
    If targetGroups.Count = 0 Then
        MsgBox "목표 표본 수를 찾지 못했습니다." & vbCrLf & _
               "B4에 전체 목표 표본 수를 입력하거나 목표표본 그룹표에 그룹별 목표 표본 수를 입력하세요.", _
               vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim records As Collection
    Set records = New Collection
    SampleDesign_ReadPopulation inputWs, schema, targetGroups, records, issues

    If records.Count = 0 Then
        MsgBox "계산할 모집단 행을 찾지 못했습니다." & vbCrLf & _
               "표본설계_입력 시트의 모집단 표에 1 이상의 모집단수를 입력하세요.", _
               vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim resultWs As Worksheet
    Dim reviewWs As Worksheet
    Set resultWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_할당결과")
    Set reviewWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_검토")

    Dim allocation As Collection
    Set allocation = SampleDesign_Allocate(records, targetGroups, methodName, alpha, minN, maxN, issues)

    Dim targetN As Long
    targetN = SampleDesign_TotalTargetN(targetGroups)

    Application.ScreenUpdating = False
    SampleDesign_WriteAllocationResult resultWs, allocation, targetN, methodName, alpha, schema("quotaHeaders")
    SampleDesign_WriteReviewReport reviewWs, allocation, targetGroups, issues, targetN, methodName, alpha
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
    Dim headerRow As Long
    headerRow = SampleDesign_FindHeaderRow(resultWs, "최종할당", 20)
    If headerRow = 0 Then
        MsgBox "할당결과 시트에서 최종할당 헤더를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim idCol As Long, popCol As Long, finalCol As Long
    idCol = SampleDesign_FindHeaderColumnInRow(resultWs, headerRow, "쿼터ID")
    popCol = SampleDesign_FindHeaderColumnInRow(resultWs, headerRow, "모집단수")
    finalCol = SampleDesign_FindHeaderColumnInRow(resultWs, headerRow, "최종할당")
    If idCol = 0 Or popCol = 0 Or finalCol = 0 Then
        MsgBox "진행관리표 생성을 위한 필수 헤더를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim outCol As Long
    outCol = 1
    outWs.Cells(3, outCol).Value = "쿼터ID": outCol = outCol + 1
    Dim c As Long
    For c = idCol + 1 To popCol - 1
        outWs.Cells(3, outCol).Value = resultWs.Cells(headerRow, c).Value
        outCol = outCol + 1
    Next c
    Dim finalOutCol As Long: finalOutCol = outCol
    SampleDesign_WriteArray outWs, 3, outCol, Array("최종할당", "회수", "부족", "초과", "달성률", "상태")

    Dim lastRow As Long
    lastRow = resultWs.Cells(resultWs.Rows.Count, idCol).End(xlUp).Row
    If lastRow <= headerRow Then
        MsgBox "할당결과 시트에 진행관리로 옮길 데이터가 없습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim r As Long
    Dim outRow As Long
    outRow = 4
    For r = headerRow + 1 To lastRow
        If Trim$(CStr(resultWs.Cells(r, idCol).Value)) <> "총계" And Len(Trim$(CStr(resultWs.Cells(r, idCol).Value))) > 0 Then
            outWs.Cells(outRow, 1).Value = resultWs.Cells(r, idCol).Value
            outCol = 2
            For c = idCol + 1 To popCol - 1
                outWs.Cells(outRow, outCol).Value = resultWs.Cells(r, c).Value
                outCol = outCol + 1
            Next c
            outWs.Cells(outRow, finalOutCol).Value = resultWs.Cells(r, finalCol).Value
            outWs.Cells(outRow, finalOutCol + 1).Value = 0
            outWs.Cells(outRow, finalOutCol + 2).FormulaR1C1 = "=MAX(RC[-2]-RC[-1],0)"
            outWs.Cells(outRow, finalOutCol + 3).FormulaR1C1 = "=MAX(RC[-2]-RC[-3],0)"
            outWs.Cells(outRow, finalOutCol + 4).FormulaR1C1 = "=IF(RC[-4]=0,0,RC[-3]/RC[-4])"
            outWs.Cells(outRow, finalOutCol + 5).FormulaR1C1 = "=IF(RC[-5]=0,""할당없음"",IF(RC[-4]=0,""미착수"",IF(RC[-2]>0,""초과"",IF(RC[-3]>0,""미달"",""충족""))))"
            outRow = outRow + 1
        End If
    Next r

    With outWs.Range(outWs.Cells(3, 1), outWs.Cells(Application.Max(3, outRow - 1), finalOutCol + 5))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    outWs.Range(outWs.Cells(3, 1), outWs.Cells(3, finalOutCol + 5)).Font.Bold = True
    outWs.Range(outWs.Cells(3, 1), outWs.Cells(3, finalOutCol + 5)).Interior.Color = RGB(220, 235, 250)
    outWs.Columns(finalOutCol + 4).NumberFormat = "0.0%"
    SampleDesign_ApplyProgressFormatting outWs, 4, outRow - 1, finalOutCol + 5
    outWs.Activate
    MsgBox "표본설계 진행관리 시트를 생성했습니다.", vbInformation, "표본설계"
End Sub

Public Sub SampleDesign_CreateBandColumn()
    On Error GoTo ErrHandler

    Dim sourceRange As Range
    If Not Toolkit_GetRangeFromUserForm("표본설계 구간 만들기", _
        "구간으로 바꿀 숫자 열을 선택하세요." & vbCrLf & _
        "예: B열의 실제 매출액 또는 종사자 수를 선택하면 오른쪽에 구간 열을 추가합니다.", _
        sourceRange) Then Exit Sub
    If sourceRange Is Nothing Then Exit Sub
    Set sourceRange = sourceRange.Areas(1)

    If sourceRange.Columns.Count <> 1 Then
        MsgBox "구간 만들기는 한 번에 숫자 열 1개만 처리합니다.", vbExclamation, "표본설계 구간 만들기"
        Exit Sub
    End If

    Set sourceRange = SampleDesign_TrimBandSourceRange(sourceRange)
    If sourceRange Is Nothing Then
        MsgBox "선택 범위에서 실제 데이터를 찾지 못했습니다.", vbExclamation, "표본설계 구간 만들기"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = sourceRange.Worksheet

    Dim hasHeader As Boolean
    hasHeader = SampleDesign_BandHasHeader(sourceRange)

    Dim firstDataRow As Long
    If hasHeader Then
        firstDataRow = sourceRange.Row + 1
    Else
        firstDataRow = sourceRange.Row
    End If

    Dim lastRow As Long
    lastRow = sourceRange.Row + sourceRange.Rows.Count - 1
    If firstDataRow > lastRow Then
        MsgBox "구간을 적용할 숫자 데이터가 없습니다.", vbExclamation, "표본설계 구간 만들기"
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ws.Parent

    Dim specText As String
    Dim outputHeader As String
    outputHeader = SampleDesign_DefaultBandHeader(sourceRange, hasHeader)
    If Not SampleDesign_ShowBandBuilder(wb, outputHeader, specText) Then Exit Sub

    Dim bands As Collection
    Set bands = New Collection
    If Not SampleDesign_ParseBandSpecs(specText, bands) Then Exit Sub
    If bands.Count = 0 Then
        MsgBox "구간 설정을 1개 이상 입력하세요.", vbExclamation, "표본설계 구간 만들기"
        Exit Sub
    End If

    Dim outputCol As Long
    outputCol = sourceRange.Column + 1
    If Not Toolkit_CanInsertColumns(ws, outputCol, 1, "표본설계 구간 만들기") Then Exit Sub

    Dim answer As VbMsgBoxResult
    answer = MsgBox("선택 열 오른쪽에 구간 열을 추가합니다." & vbCrLf & vbCrLf & _
                    "원본 열: " & sourceRange.Address(False, False) & vbCrLf & _
                    "결과 헤더: " & outputHeader & vbCrLf & _
                    "구간 수: " & bands.Count & vbCrLf & vbCrLf & _
                    "계속할까요?", _
                    vbQuestion + vbYesNo, "표본설계 구간 만들기")
    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    ws.Columns(outputCol).Insert Shift:=xlToRight

    If hasHeader Then
        ws.Cells(sourceRange.Row, outputCol).Value = outputHeader
        ws.Cells(sourceRange.Row, outputCol).Font.Bold = True
        ws.Cells(sourceRange.Row, outputCol).Interior.Color = RGB(220, 235, 250)
    End If

    Dim rowCount As Long
    rowCount = lastRow - firstDataRow + 1
    Dim formulaData() As Variant
    ReDim formulaData(1 To rowCount, 1 To 1)

    Dim r As Long
    Dim outIndex As Long
    For r = firstDataRow To lastRow
        outIndex = r - firstDataRow + 1
        formulaData(outIndex, 1) = SampleDesign_BandFormulaForAddress(ws.Cells(r, sourceRange.Column).Address(False, False), bands)
    Next r

    ws.Range(ws.Cells(firstDataRow, outputCol), ws.Cells(lastRow, outputCol)).Formula = formulaData
    ws.Columns(outputCol).AutoFit
    SampleDesign_SaveBandSettings wb, outputHeader, specText
    Application.ScreenUpdating = True

    MsgBox "구간 열을 생성했습니다." & vbCrLf & _
           "수식 입력: " & rowCount & "행" & vbCrLf & _
           "결과 열에는 IFS 수식이 입력되어 기준을 직접 확인하고 수정할 수 있습니다." & vbCrLf & _
           "설정은 숨김 시트에 임시 저장되었습니다.", _
           vbInformation, "표본설계 구간 만들기"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "구간 생성 중 오류가 발생했습니다." & vbCrLf & Err.Description, _
           vbExclamation, "표본설계 구간 만들기"
End Sub

Public Sub SampleDesign_ReviewQuota()
    SampleDesign_RunAllocation
End Sub

Public Sub SampleDesign_ListifyPopulationTable()
    Dim modeChoice As String
    modeChoice = Toolkit_ButtonMenu( _
        "모집단 구성표 리스트화", _
        "데이터 구조 지정 방식을 선택하세요.", _
        "1" & vbTab & "자동 감지 후 미리보기" & vbLf & _
        "2" & vbTab & "범위 직접 지정 후 미리보기", _
        "1")
    If Len(modeChoice) = 0 Then Exit Sub
    If modeChoice = "2" Then
        SampleDesign_ListifyPopulationTableManual
        Exit Sub
    End If

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

Public Sub SampleDesign_ListifyPopulationTableManual()
    Dim rowAxisRange As Range
    Dim columnHeaderRange As Range
    Dim valueRange As Range

    If Not Toolkit_GetRangeFromUserForm("모집단 구성표 리스트화 - 행축 데이터", _
        "세로로 반복되는 항목 범위를 선택하세요." & vbCrLf & _
        "예: 업종, 지역, 성별처럼 행 방향으로 내려가는 항목 열입니다. 여러 열을 함께 선택할 수 있습니다.", _
        rowAxisRange) Then Exit Sub
    If rowAxisRange Is Nothing Then Exit Sub
    Set rowAxisRange = rowAxisRange.Areas(1)

    If Not Toolkit_GetRangeFromUserForm("모집단 구성표 리스트화 - 가로 헤더", _
        "가로축 항목 헤더 범위를 선택하세요." & vbCrLf & _
        "예: 종사자 규모, 연령대, 성별처럼 값 영역 위에 있는 헤더 행입니다.", _
        columnHeaderRange) Then Exit Sub
    If columnHeaderRange Is Nothing Then Exit Sub
    Set columnHeaderRange = columnHeaderRange.Areas(1).Rows(1)

    If Not Toolkit_GetRangeFromUserForm("모집단 구성표 리스트화 - 값 영역", _
        "숫자 값 영역을 선택하세요." & vbCrLf & _
        "행 수는 행축 데이터 범위와 같고, 열 수는 가로 헤더 범위와 같아야 합니다.", _
        valueRange) Then Exit Sub
    If valueRange Is Nothing Then Exit Sub
    Set valueRange = valueRange.Areas(1)

    If Not PopList_ManualRangesValid(rowAxisRange, columnHeaderRange, valueRange) Then Exit Sub

    Dim tableTitle As String
    tableTitle = InputBox("표 제목을 입력하세요. 비워두면 시트명과 값 영역 주소를 사용합니다.", _
                          "모집단 구성표 리스트화", _
                          PopList_ManualDefaultTitle(rowAxisRange, valueRange))
    If Len(Trim$(tableTitle)) = 0 Then tableTitle = PopList_ManualDefaultTitle(rowAxisRange, valueRange)

    Dim columnAxisName As String
    Dim valueName As String
    Dim outputSheetName As String
    columnAxisName = "가로구분"
    valueName = "값"
    outputSheetName = "모집단_리스트"

    If Not PopList_ShowManualPreviewForm(tableTitle, rowAxisRange, columnHeaderRange, valueRange, _
                                         columnAxisName, valueName, outputSheetName) Then Exit Sub

    outputSheetName = PopList_CleanSheetName(outputSheetName, "모집단_리스트")
    Dim outWs As Worksheet
    Set outWs = SampleDesign_PrepareOutputSheet(rowAxisRange.Worksheet.Parent, outputSheetName)

    Dim writtenRows As Long
    writtenRows = PopList_WriteManualOutput(tableTitle, rowAxisRange, columnHeaderRange, valueRange, outWs, columnAxisName, valueName)
    outWs.Activate

    MsgBox "직접 지정한 구성표에서 " & writtenRows & "개 리스트 행을 생성했습니다." & vbCrLf & _
           "결과 시트: " & outWs.Name, vbInformation, "모집단 구성표 리스트화"
End Sub

Public Sub SampleDesign_PivotPopulationList()
    Dim sourceRange As Range
    If Not Toolkit_GetRangeFromUserForm("리스트→집계표", _
        "집계표 형태로 되돌릴 리스트 범위를 선택하세요." & vbCrLf & _
        "구성표 리스트화 결과 또는 표본설계 할당결과 범위를 사용할 수 있습니다.", _
        sourceRange) Then Exit Sub
    If sourceRange Is Nothing Then Exit Sub
    Set sourceRange = sourceRange.Areas(1)

    Dim schema As Object
    Set schema = SampleDesign_PivotDetectSchema(sourceRange)
    If schema Is Nothing Then
        MsgBox "선택 범위에서 집계표로 변환할 리스트 구조를 찾지 못했습니다." & vbCrLf & _
               "표제목/행축값/가로구분/값 또는 쿼터/최종할당 구조가 포함되도록 범위를 다시 선택하세요.", _
               vbExclamation, "리스트→집계표"
        Exit Sub
    End If

    Dim summary As String
    summary = "탐지된 구조" & vbCrLf & _
              "- 표 구분: " & CStr(schema("blockLabel")) & vbCrLf & _
              "- 행축: " & SampleDesign_JoinCollection(schema("rowAxisLabels"), ", ") & vbCrLf & _
              "- 열축: " & CStr(schema("columnAxisLabel")) & vbCrLf & _
              "- 값: " & CStr(schema("valueLabel")) & vbCrLf & vbCrLf & _
              "이 구조로 집계표를 생성할까요?"
    If MsgBox(summary, vbQuestion + vbYesNo, "리스트→집계표") <> vbYes Then Exit Sub

    Dim outWs As Worksheet
    Set outWs = SampleDesign_PrepareOutputSheet(sourceRange.Worksheet.Parent, "표본설계_집계표")
    Dim blockCount As Long
    blockCount = SampleDesign_PivotWriteOutput(sourceRange, schema, outWs)
    outWs.Activate

    MsgBox blockCount & "개 집계표 블록을 생성했습니다." & vbCrLf & _
           "결과 시트: " & outWs.Name, vbInformation, "리스트→집계표"
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

Private Function PopList_ShowManualPreviewForm(ByVal tableTitle As String, ByVal rowAxisRange As Range, _
                                               ByVal columnHeaderRange As Range, ByVal valueRange As Range, _
                                               ByRef columnAxisName As String, ByRef valueName As String, _
                                               ByRef outputSheetName As String) As Boolean
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmPopulationListifyPreview")
    form.Configure PopList_BuildManualSummary(tableTitle, rowAxisRange, columnHeaderRange, valueRange), _
                   PopList_BuildManualPreview(tableTitle, rowAxisRange, columnHeaderRange, valueRange, columnAxisName, valueName, 25), _
                   columnAxisName, valueName, outputSheetName
    form.Show

    If CBool(form.Confirmed) Then
        columnAxisName = Trim$(CStr(form.ColumnAxisName))
        valueName = Trim$(CStr(form.ValueName))
        outputSheetName = Trim$(CStr(form.OutputSheetName))
        If Len(columnAxisName) = 0 Then columnAxisName = "가로구분"
        If Len(valueName) = 0 Then valueName = "값"
        If Len(outputSheetName) = 0 Then outputSheetName = "모집단_리스트"
        PopList_ShowManualPreviewForm = True
    End If
    Unload form
    Exit Function

Fallback:
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    On Error GoTo 0

    If MsgBox("미리보기 창을 열지 못했습니다. 직접 지정한 구조로 바로 실행할까요?" & vbCrLf & vbCrLf & _
              PopList_BuildManualSummary(tableTitle, rowAxisRange, columnHeaderRange, valueRange), _
              vbQuestion + vbYesNo, "모집단 구성표 리스트화") = vbYes Then
        PopList_ShowManualPreviewForm = True
    End If
End Function

Private Function PopList_ManualRangesValid(ByVal rowAxisRange As Range, ByVal columnHeaderRange As Range, ByVal valueRange As Range) As Boolean
    If Not (rowAxisRange.Worksheet Is columnHeaderRange.Worksheet) Or Not (rowAxisRange.Worksheet Is valueRange.Worksheet) Then
        MsgBox "행축, 가로 헤더, 값 영역은 모두 같은 시트에서 선택해야 합니다.", vbExclamation, "모집단 구성표 리스트화"
        Exit Function
    End If

    If rowAxisRange.Rows.Count <> valueRange.Rows.Count Then
        MsgBox "행축 데이터 범위의 행 수와 값 영역의 행 수가 다릅니다." & vbCrLf & _
               "행축 행 수: " & rowAxisRange.Rows.Count & vbCrLf & _
               "값 영역 행 수: " & valueRange.Rows.Count, vbExclamation, "모집단 구성표 리스트화"
        Exit Function
    End If

    If columnHeaderRange.Columns.Count <> valueRange.Columns.Count Then
        MsgBox "가로 헤더 범위의 열 수와 값 영역의 열 수가 다릅니다." & vbCrLf & _
               "가로 헤더 열 수: " & columnHeaderRange.Columns.Count & vbCrLf & _
               "값 영역 열 수: " & valueRange.Columns.Count, vbExclamation, "모집단 구성표 리스트화"
        Exit Function
    End If

    If rowAxisRange.Rows.Count = 0 Or rowAxisRange.Columns.Count = 0 _
       Or columnHeaderRange.Columns.Count = 0 Or valueRange.Rows.Count = 0 Or valueRange.Columns.Count = 0 Then
        MsgBox "선택한 범위가 비어 있습니다.", vbExclamation, "모집단 구성표 리스트화"
        Exit Function
    End If

    PopList_ManualRangesValid = True
End Function

Private Function PopList_ManualDefaultTitle(ByVal rowAxisRange As Range, ByVal valueRange As Range) As String
    PopList_ManualDefaultTitle = rowAxisRange.Worksheet.Name & "!" & valueRange.Address(False, False)
End Function

Private Function PopList_ManualAxisHeaders(ByVal rowAxisRange As Range) As Collection
    Dim headers As Collection
    Set headers = New Collection

    Dim axisIndex As Long
    For axisIndex = 1 To rowAxisRange.Columns.Count
        Dim headerText As String
        If rowAxisRange.Row > 1 Then
            headerText = PopList_CellText(rowAxisRange.Worksheet.Cells(rowAxisRange.Row - 1, rowAxisRange.Column + axisIndex - 1))
        End If
        If Len(headerText) = 0 Then headerText = "행축" & axisIndex
        headers.Add headerText
    Next axisIndex

    Set PopList_ManualAxisHeaders = headers
End Function

Private Function PopList_BuildManualSummary(ByVal tableTitle As String, ByVal rowAxisRange As Range, _
                                            ByVal columnHeaderRange As Range, ByVal valueRange As Range) As String
    Dim headers As Collection
    Set headers = PopList_ManualAxisHeaders(rowAxisRange)

    PopList_BuildManualSummary = _
        "선택 방식: 범위 직접 지정" & vbCrLf & _
        "표 제목: " & tableTitle & vbCrLf & _
        "원본 시트: " & rowAxisRange.Worksheet.Name & vbCrLf & _
        "행축 데이터: " & rowAxisRange.Address(False, False) & " (" & rowAxisRange.Rows.Count & "행 x " & rowAxisRange.Columns.Count & "열)" & vbCrLf & _
        "행축 이름: " & PopList_JoinCollection(headers, ", ") & vbCrLf & _
        "가로 헤더: " & columnHeaderRange.Address(False, False) & " (" & columnHeaderRange.Columns.Count & "열)" & vbCrLf & _
        "값 영역: " & valueRange.Address(False, False) & " (" & valueRange.Rows.Count & "행 x " & valueRange.Columns.Count & "열)" & vbCrLf & _
        "합계 행/열은 이름이 계/합계/총계/전체/TOTAL인 경우 결과에서 제외됩니다."
End Function

Private Function PopList_BuildManualPreview(ByVal tableTitle As String, ByVal rowAxisRange As Range, _
                                            ByVal columnHeaderRange As Range, ByVal valueRange As Range, _
                                            ByVal columnAxisName As String, ByVal valueName As String, _
                                            ByVal maxRows As Long) As String
    Dim headers As Collection
    Set headers = PopList_ManualAxisHeaders(rowAxisRange)

    Dim result As String
    result = "표제목"

    Dim axisIndex As Long
    For axisIndex = 1 To headers.Count
        result = result & vbTab & CStr(headers(axisIndex))
    Next axisIndex
    result = result & vbTab & columnAxisName & vbTab & valueName & vbCrLf

    Dim lastSeen() As String
    ReDim lastSeen(1 To rowAxisRange.Columns.Count)

    Dim madeRows As Long
    Dim relRow As Long
    For relRow = 1 To valueRange.Rows.Count
        PopList_UpdateManualLastSeen rowAxisRange, relRow, lastSeen
        If Not PopList_ManualRowIsTotal(rowAxisRange, relRow) Then
            Dim relCol As Long
            For relCol = 1 To valueRange.Columns.Count
                Dim headerText As String
                headerText = PopList_CellText(columnHeaderRange.Cells(1, relCol))
                If Len(headerText) > 0 And Not PopList_IsTotalLabel(headerText) Then
                    If PopList_IsNumberLike(valueRange.Cells(relRow, relCol).Value) Then
                        result = result & tableTitle
                        For axisIndex = 1 To rowAxisRange.Columns.Count
                            result = result & vbTab & lastSeen(axisIndex)
                        Next axisIndex
                        result = result & vbTab & headerText & vbTab & PopList_ToNumber(valueRange.Cells(relRow, relCol).Value) & vbCrLf
                        madeRows = madeRows + 1
                        If madeRows >= maxRows Then
                            PopList_BuildManualPreview = result & "... (미리보기 " & maxRows & "행까지만 표시)"
                            Exit Function
                        End If
                    End If
                End If
            Next relCol
        End If
    Next relRow

    PopList_BuildManualPreview = result
End Function

Private Function PopList_WriteManualOutput(ByVal tableTitle As String, ByVal rowAxisRange As Range, _
                                           ByVal columnHeaderRange As Range, ByVal valueRange As Range, _
                                           ByVal outWs As Worksheet, ByVal columnAxisName As String, _
                                           ByVal valueName As String) As Long
    outWs.Cells.Clear
    outWs.Range("A1").Value = "모집단 구성표 리스트"
    outWs.Range("A1").Font.Bold = True
    outWs.Range("A1").Font.Size = 14
    outWs.Range("A2").Value = "원본 시트"
    outWs.Range("B2").Value = rowAxisRange.Worksheet.Name
    outWs.Range("A3").Value = "원본 범위"
    outWs.Range("B3").Value = "행축 " & rowAxisRange.Address(False, False) & _
                               " / 가로헤더 " & columnHeaderRange.Address(False, False) & _
                               " / 값 " & valueRange.Address(False, False)

    Dim headers As Collection
    Set headers = PopList_ManualAxisHeaders(rowAxisRange)

    Dim headerRow As Long
    headerRow = 5

    Dim outCol As Long
    outCol = 1
    outWs.Cells(headerRow, outCol).Value = "표제목"
    outCol = outCol + 1

    Dim axisIndex As Long
    For axisIndex = 1 To headers.Count
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

    Dim lastSeen() As String
    ReDim lastSeen(1 To rowAxisRange.Columns.Count)

    Dim outRow As Long
    outRow = headerRow + 1

    Dim relRow As Long
    For relRow = 1 To valueRange.Rows.Count
        PopList_UpdateManualLastSeen rowAxisRange, relRow, lastSeen
        If Not PopList_ManualRowIsTotal(rowAxisRange, relRow) Then
            Dim relCol As Long
            For relCol = 1 To valueRange.Columns.Count
                Dim headerText As String
                headerText = PopList_CellText(columnHeaderRange.Cells(1, relCol))
                If Len(headerText) > 0 And Not PopList_IsTotalLabel(headerText) Then
                    If PopList_IsNumberLike(valueRange.Cells(relRow, relCol).Value) Then
                        outWs.Cells(outRow, 1).Value = tableTitle

                        outCol = 2
                        For axisIndex = 1 To headers.Count
                            outWs.Cells(outRow, outCol).Value = CStr(headers(axisIndex))
                            outWs.Cells(outRow, outCol + 1).Value = lastSeen(axisIndex)
                            outCol = outCol + 2
                        Next axisIndex

                        outWs.Cells(outRow, columnAxisCol).Value = headerText
                        outWs.Cells(outRow, valueCol).Value = PopList_ToNumber(valueRange.Cells(relRow, relCol).Value)
                        outWs.Cells(outRow, sourceRowCol).Value = valueRange.Row + relRow - 1
                        outWs.Cells(outRow, sourceColCol).Value = valueRange.Column + relCol - 1
                        outRow = outRow + 1
                    End If
                End If
            Next relCol
        End If
    Next relRow

    Dim lastCol As Long
    lastCol = sourceColCol
    With outWs.Range(outWs.Cells(headerRow, 1), outWs.Cells(Application.Max(headerRow, outRow - 1), lastCol))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    outWs.Range(outWs.Cells(headerRow, 1), outWs.Cells(headerRow, lastCol)).Font.Bold = True
    outWs.Range(outWs.Cells(headerRow, 1), outWs.Cells(headerRow, lastCol)).Interior.Color = RGB(220, 235, 250)
    outWs.Columns(valueCol).NumberFormat = "#,##0"

    PopList_WriteManualOutput = outRow - headerRow - 1
End Function

Private Sub PopList_UpdateManualLastSeen(ByVal rowAxisRange As Range, ByVal relRow As Long, ByRef lastSeen() As String)
    Dim axisIndex As Long
    For axisIndex = 1 To rowAxisRange.Columns.Count
        Dim axisValue As String
        axisValue = PopList_CellText(rowAxisRange.Cells(relRow, axisIndex))
        If Len(axisValue) > 0 Then lastSeen(axisIndex) = axisValue
    Next axisIndex
End Sub

Private Function PopList_ManualRowIsTotal(ByVal rowAxisRange As Range, ByVal relRow As Long) As Boolean
    Dim axisIndex As Long
    For axisIndex = 1 To rowAxisRange.Columns.Count
        If PopList_IsTotalLabel(PopList_CellText(rowAxisRange.Cells(relRow, axisIndex))) Then
            PopList_ManualRowIsTotal = True
            Exit Function
        End If
    Next axisIndex
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

Private Function SampleDesign_PivotDetectSchema(ByVal sourceRange As Range) As Object
    Dim headerRelRow As Long
    headerRelRow = SampleDesign_PivotFindHeaderRelRow(sourceRange)
    If headerRelRow = 0 Then Exit Function

    Dim valueCol As Long
    valueCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "값")
    If valueCol = 0 Then valueCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "최종할당")
    If valueCol = 0 Then valueCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "모집단수")
    If valueCol = 0 Then Exit Function

    Dim columnCol As Long
    columnCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "가로구분")
    If columnCol = 0 Then columnCol = SampleDesign_PivotFallbackColumnAxisCol(sourceRange, headerRelRow, valueCol)
    If columnCol = 0 Then Exit Function

    Dim rowAxisCols As Collection
    Dim rowAxisLabels As Collection
    Set rowAxisCols = New Collection
    Set rowAxisLabels = New Collection

    Dim n As Long, relCol As Long
    For n = 1 To 20
        relCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "행축" & n & "값")
        If relCol > 0 Then
            rowAxisCols.Add relCol
            rowAxisLabels.Add SampleDesign_PivotAxisLabel(sourceRange, headerRelRow, n, relCol)
        End If
    Next n

    If rowAxisCols.Count = 0 Then
        Dim idCol As Long, popCol As Long
        idCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "쿼터ID")
        popCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "모집단수")
        If idCol > 0 And popCol > idCol + 1 Then
            For relCol = idCol + 1 To popCol - 1
                If relCol <> columnCol Then
                    rowAxisCols.Add relCol
                    rowAxisLabels.Add PopList_CellText(sourceRange.Cells(headerRelRow, relCol))
                End If
            Next relCol
        End If
    End If
    If rowAxisCols.Count = 0 Then Exit Function

    Dim schema As Object
    Set schema = CreateObject("Scripting.Dictionary")
    schema.Add "headerRelRow", headerRelRow
    schema.Add "dataStartRelRow", headerRelRow + 1
    schema.Add "blockCol", SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "표제목")
    schema.Add "blockLabel", IIf(SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "표제목") > 0, "표제목", "단일 표")
    schema.Add "rowAxisCols", rowAxisCols
    schema.Add "rowAxisLabels", rowAxisLabels
    schema.Add "columnCol", columnCol
    schema.Add "columnAxisLabel", PopList_CellText(sourceRange.Cells(headerRelRow, columnCol))
    schema.Add "valueCol", valueCol
    schema.Add "valueLabel", PopList_CellText(sourceRange.Cells(headerRelRow, valueCol))
    Set SampleDesign_PivotDetectSchema = schema
End Function

Private Function SampleDesign_PivotFindHeaderRelRow(ByVal sourceRange As Range) As Long
    Dim scanRows As Long
    scanRows = sourceRange.Rows.Count
    If scanRows > 30 Then scanRows = 30
    Dim relRow As Long
    For relRow = 1 To scanRows
        If SampleDesign_PivotHeaderCol(sourceRange, relRow, "값") > 0 _
           Or SampleDesign_PivotHeaderCol(sourceRange, relRow, "최종할당") > 0 _
           Or SampleDesign_PivotHeaderCol(sourceRange, relRow, "모집단수") > 0 Then
            SampleDesign_PivotFindHeaderRelRow = relRow
            Exit Function
        End If
    Next relRow
End Function

Private Function SampleDesign_PivotHeaderCol(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal headerText As String) As Long
    Dim relCol As Long
    For relCol = 1 To sourceRange.Columns.Count
        If Trim$(CStr(sourceRange.Cells(headerRelRow, relCol).Value)) = headerText Then
            SampleDesign_PivotHeaderCol = relCol
            Exit Function
        End If
    Next relCol
End Function

Private Function SampleDesign_PivotFallbackColumnAxisCol(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal valueCol As Long) As Long
    Dim idCol As Long, popCol As Long, relCol As Long
    idCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "쿼터ID")
    popCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "모집단수")
    If idCol > 0 And popCol > idCol + 1 Then
        For relCol = popCol - 1 To idCol + 1 Step -1
            If Len(PopList_CellText(sourceRange.Cells(headerRelRow, relCol))) > 0 Then
                SampleDesign_PivotFallbackColumnAxisCol = relCol
                Exit Function
            End If
        Next relCol
    End If
End Function

Private Function SampleDesign_PivotAxisLabel(ByVal sourceRange As Range, ByVal headerRelRow As Long, ByVal axisIndex As Long, ByVal valueCol As Long) As String
    Dim nameCol As Long
    nameCol = SampleDesign_PivotHeaderCol(sourceRange, headerRelRow, "행축" & axisIndex & "명")
    If nameCol > 0 Then
        Dim relRow As Long, lastRelRow As Long
        lastRelRow = sourceRange.Rows.Count
        If lastRelRow > headerRelRow + 100 Then lastRelRow = headerRelRow + 100
        For relRow = headerRelRow + 1 To lastRelRow
            SampleDesign_PivotAxisLabel = PopList_CellText(sourceRange.Cells(relRow, nameCol))
            If Len(SampleDesign_PivotAxisLabel) > 0 Then Exit Function
        Next relRow
    End If
    SampleDesign_PivotAxisLabel = PopList_CellText(sourceRange.Cells(headerRelRow, valueCol))
    If Len(SampleDesign_PivotAxisLabel) = 0 Then SampleDesign_PivotAxisLabel = "행축" & axisIndex
End Function

Private Function SampleDesign_PivotWriteOutput(ByVal sourceRange As Range, ByVal schema As Object, ByVal outWs As Worksheet) As Long
    outWs.Cells.Clear
    outWs.Range("A1").Value = "리스트→집계표"
    outWs.Range("A1").Font.Bold = True
    outWs.Range("A1").Font.Size = 14

    Dim blocks As Object
    Dim blockOrder As Collection
    Set blocks = CreateObject("Scripting.Dictionary")
    Set blockOrder = New Collection

    Dim relRow As Long
    For relRow = CLng(schema("dataStartRelRow")) To sourceRange.Rows.Count
        Dim valueData As Double
        valueData = SampleDesign_CellNumber(sourceRange.Cells(relRow, CLng(schema("valueCol"))).Value)
        If valueData <> 0 Or Len(Trim$(CStr(sourceRange.Cells(relRow, CLng(schema("valueCol"))).Value))) > 0 Then
            Dim blockKey As String
            blockKey = "집계표"
            If CLng(schema("blockCol")) > 0 Then blockKey = PopList_CellText(sourceRange.Cells(relRow, CLng(schema("blockCol"))))
            If Len(blockKey) = 0 Then blockKey = "집계표"

            Dim block As Object
            Set block = SampleDesign_PivotEnsureBlock(blocks, blockOrder, blockKey)

            Dim rowKey As String
            rowKey = SampleDesign_PivotRowKey(sourceRange, schema, relRow)
            If Len(rowKey) > 0 Then
                Dim colKey As String
                colKey = PopList_CellText(sourceRange.Cells(relRow, CLng(schema("columnCol"))))
                If Len(colKey) = 0 Then colKey = "(빈값)"

                If Not block("rows").Exists(rowKey) Then block("rows").Add rowKey, rowKey
                SampleDesign_AddUniqueCollection block("columns"), colKey

                Dim dataKey As String
                dataKey = rowKey & Chr$(31) & colKey
                If block("data").Exists(dataKey) Then
                    block("data")(dataKey) = CDbl(block("data")(dataKey)) + valueData
                Else
                    block("data").Add dataKey, valueData
                End If
            End If
        End If
    Next relRow

    Dim outRow As Long
    outRow = 3
    Dim blockName As Variant
    For Each blockName In blockOrder
        Dim currentBlock As Object
        Set currentBlock = blocks(CStr(blockName))
        outRow = SampleDesign_PivotWriteBlock(outWs, outRow, CStr(blockName), schema, currentBlock)
        outRow = outRow + 2
    Next blockName

    outWs.Columns.AutoFit
    SampleDesign_PivotWriteOutput = blockOrder.Count
End Function

Private Function SampleDesign_PivotEnsureBlock(ByVal blocks As Object, ByVal blockOrder As Collection, ByVal blockKey As String) As Object
    If Not blocks.Exists(blockKey) Then
        Dim block As Object
        Set block = CreateObject("Scripting.Dictionary")
        Dim rows As Object
        Set rows = CreateObject("Scripting.Dictionary")
        Dim cols As Collection
        Set cols = New Collection
        Dim data As Object
        Set data = CreateObject("Scripting.Dictionary")
        block.Add "rows", rows
        block.Add "columns", cols
        block.Add "data", data
        blocks.Add blockKey, block
        blockOrder.Add blockKey
    End If
    Set SampleDesign_PivotEnsureBlock = blocks(blockKey)
End Function

Private Function SampleDesign_PivotRowKey(ByVal sourceRange As Range, ByVal schema As Object, ByVal relRow As Long) As String
    Dim i As Long
    For i = 1 To schema("rowAxisCols").Count
        If i > 1 Then SampleDesign_PivotRowKey = SampleDesign_PivotRowKey & Chr$(30)
        SampleDesign_PivotRowKey = SampleDesign_PivotRowKey & PopList_CellText(sourceRange.Cells(relRow, CLng(schema("rowAxisCols")(i))))
    Next i
End Function

Private Sub SampleDesign_AddUniqueCollection(ByVal items As Collection, ByVal itemText As String)
    Dim item As Variant
    For Each item In items
        If StrComp(CStr(item), itemText, vbTextCompare) = 0 Then Exit Sub
    Next item
    items.Add itemText
End Sub

Private Function SampleDesign_PivotWriteBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal blockTitle As String, ByVal schema As Object, ByVal block As Object) As Long
    ws.Cells(startRow, 1).Value = blockTitle
    ws.Cells(startRow, 1).Font.Bold = True

    Dim headerRow As Long
    headerRow = startRow + 1
    Dim col As Long, i As Long
    col = 1
    For i = 1 To schema("rowAxisLabels").Count
        ws.Cells(headerRow, col).Value = schema("rowAxisLabels")(i)
        col = col + 1
    Next i

    Dim firstValueCol As Long
    firstValueCol = col
    For i = 1 To block("columns").Count
        ws.Cells(headerRow, col).Value = block("columns")(i)
        col = col + 1
    Next i
    ws.Cells(headerRow, col).Value = "합계"
    Dim totalCol As Long
    totalCol = col

    Dim outRow As Long
    outRow = headerRow + 1
    Dim rowKey As Variant
    For Each rowKey In block("rows").Keys
        Dim parts As Variant
        parts = Split(CStr(rowKey), Chr$(30))
        For i = LBound(parts) To UBound(parts)
            ws.Cells(outRow, 1 + i).Value = parts(i)
        Next i

        Dim rowTotal As Double
        For i = 1 To block("columns").Count
            Dim dataKey As String
            dataKey = CStr(rowKey) & Chr$(31) & CStr(block("columns")(i))
            If block("data").Exists(dataKey) Then
                ws.Cells(outRow, firstValueCol + i - 1).Value = CDbl(block("data")(dataKey))
                rowTotal = rowTotal + CDbl(block("data")(dataKey))
            End If
        Next i
        ws.Cells(outRow, totalCol).Value = rowTotal
        outRow = outRow + 1
    Next rowKey

    ws.Cells(outRow, 1).Value = "합계"
    For i = firstValueCol To totalCol
        ws.Cells(outRow, i).Formula = "=SUM(" & ws.Range(ws.Cells(headerRow + 1, i), ws.Cells(outRow - 1, i)).Address(False, False) & ")"
    Next i

    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(outRow, totalCol))
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, totalCol)).Font.Bold = True
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, totalCol)).Interior.Color = RGB(220, 235, 250)
    ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, totalCol)).Font.Bold = True
    ws.Range(ws.Cells(headerRow + 1, firstValueCol), ws.Cells(outRow, totalCol)).NumberFormat = "#,##0"

    SampleDesign_PivotWriteBlock = outRow + 1
End Function

Private Function SampleDesign_JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To items.Count
        If i > 1 Then SampleDesign_JoinCollection = SampleDesign_JoinCollection & delimiter
        SampleDesign_JoinCollection = SampleDesign_JoinCollection & CStr(items(i))
    Next i
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

Private Function SampleDesign_TrimBandSourceRange(ByVal sourceRange As Range) As Range
    Dim usedPart As Range
    On Error Resume Next
    Set usedPart = Application.Intersect(sourceRange, sourceRange.Worksheet.UsedRange)
    On Error GoTo 0
    If usedPart Is Nothing Then
        Set SampleDesign_TrimBandSourceRange = sourceRange
    Else
        Set SampleDesign_TrimBandSourceRange = usedPart.Areas(1)
    End If
End Function

Private Function SampleDesign_BandHasHeader(ByVal sourceRange As Range) As Boolean
    If sourceRange.Rows.Count < 2 Then Exit Function
    Dim numericValue As Double
    SampleDesign_BandHasHeader = Not SampleDesign_BandTryNumber(sourceRange.Cells(1, 1).Value, numericValue)
End Function

Private Function SampleDesign_DefaultBandHeader(ByVal sourceRange As Range, ByVal hasHeader As Boolean) As String
    Dim baseHeader As String
    If hasHeader Then baseHeader = Trim$(CStr(sourceRange.Cells(1, 1).Value))
    If Len(baseHeader) = 0 Then baseHeader = "값"
    SampleDesign_DefaultBandHeader = baseHeader & "_구간"
End Function

Private Function SampleDesign_ShowBandBuilder(ByVal wb As Workbook, ByRef outputHeader As String, ByRef specText As String) As Boolean
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmSampleDesignBandBuilder")
    form.Configure SampleDesign_LoadBandSpecText(wb), SampleDesign_LoadBandHeader(wb, outputHeader)
    form.Show
    If CBool(form.Cancelled) Then
        Unload form
        Exit Function
    End If

    outputHeader = Trim$(CStr(form.OutputHeader))
    specText = CStr(form.BandSpecText)
    Unload form
    If Len(outputHeader) = 0 Then outputHeader = "구간"
    SampleDesign_ShowBandBuilder = True
    Exit Function

Fallback:
    SampleDesign_ShowBandBuilder = SampleDesign_ShowBandBuilderFallback(outputHeader, specText)
End Function

Private Function SampleDesign_ShowBandBuilderFallback(ByRef outputHeader As String, ByRef specText As String) As Boolean
    Dim headerInput As Variant
    headerInput = Application.InputBox("결과 열 헤더를 입력하세요.", "표본설계 구간 만들기", outputHeader, Type:=2)
    If VarType(headerInput) = vbBoolean And headerInput = False Then Exit Function
    outputHeader = Trim$(CStr(headerInput))
    If Len(outputHeader) = 0 Then outputHeader = "구간"

    Dim countInput As Variant
    countInput = Application.InputBox("구간 개수를 입력하세요. 최대 15개까지 가능합니다.", "표본설계 구간 만들기", 5, Type:=1)
    If VarType(countInput) = vbBoolean And countInput = False Then Exit Function
    Dim bandCount As Long
    bandCount = CLng(countInput)
    If bandCount < 1 Or bandCount > 15 Then
        MsgBox "구간 개수는 1~15 사이로 입력하세요.", vbExclamation, "표본설계 구간 만들기"
        Exit Function
    End If

    Dim i As Long
    For i = 1 To bandCount
        Dim startText As String, endText As String, labelText As String
        Dim startInput As Variant, endInput As Variant, labelInput As Variant
        startInput = Application.InputBox(i & "구간 시작값을 입력하세요. 첫 구간은 비워도 됩니다.", "표본설계 구간 만들기", "", Type:=2)
        If VarType(startInput) = vbBoolean And startInput = False Then Exit Function
        endInput = Application.InputBox(i & "구간 종료값을 입력하세요. 마지막 구간은 비워도 됩니다.", "표본설계 구간 만들기", "", Type:=2)
        If VarType(endInput) = vbBoolean And endInput = False Then Exit Function
        labelInput = Application.InputBox(i & "구간명을 입력하세요. 비우면 조건으로 자동 생성됩니다.", "표본설계 구간 만들기", "", Type:=2)
        If VarType(labelInput) = vbBoolean And labelInput = False Then Exit Function
        startText = CStr(startInput)
        endText = CStr(endInput)
        labelText = CStr(labelInput)
        If Len(specText) > 0 Then specText = specText & vbLf
        specText = specText & CStr(i) & vbTab & Trim$(startText) & vbTab & "이상" & vbTab & Trim$(endText) & vbTab & "미만" & vbTab & Trim$(labelText)
    Next i

    SampleDesign_ShowBandBuilderFallback = True
End Function

Private Function SampleDesign_ParseBandSpecs(ByVal specText As String, ByRef bands As Collection) As Boolean
    Dim lines As Variant
    lines = Split(Replace(CStr(specText), vbCrLf, vbLf), vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(CStr(lines(i)))
        If Len(lineText) > 0 Then
            Dim parts As Variant
            parts = Split(lineText, vbTab)
            If UBound(parts) < 4 Then
                MsgBox "구간 설정 형식이 올바르지 않습니다.", vbExclamation, "표본설계 구간 만들기"
                Exit Function
            End If

            Dim band As Object
            Set band = CreateObject("Scripting.Dictionary")

            Dim startValue As Double, endValue As Double
            Dim hasStart As Boolean, hasEnd As Boolean
            hasStart = SampleDesign_BandTryNumber(parts(1), startValue)
            hasEnd = SampleDesign_BandTryNumber(parts(3), endValue)
            If Not hasStart And Not hasEnd Then
                MsgBox "각 구간은 시작값 또는 종료값 중 하나 이상이 필요합니다.", vbExclamation, "표본설계 구간 만들기"
                Exit Function
            End If
            If hasStart And hasEnd And startValue >= endValue Then
                MsgBox "구간 시작값은 종료값보다 작아야 합니다." & vbCrLf & _
                       "문제 구간: " & CStr(parts(0)), vbExclamation, "표본설계 구간 만들기"
                Exit Function
            End If

            band.Add "hasStart", hasStart
            band.Add "startValue", startValue
            band.Add "startOp", SampleDesign_NormalizeBandStartOp(CStr(parts(2)))
            band.Add "hasEnd", hasEnd
            band.Add "endValue", endValue
            band.Add "endOp", SampleDesign_NormalizeBandEndOp(CStr(parts(4)))
            If UBound(parts) >= 5 Then
                band.Add "label", Trim$(CStr(parts(5)))
            Else
                band.Add "label", ""
            End If
            If Len(CStr(band("label"))) = 0 Then band("label") = SampleDesign_BandAutoLabel(band)
            bands.Add band
        End If
    Next i

    SampleDesign_ParseBandSpecs = True
End Function

Private Function SampleDesign_BandTryNumber(ByVal valueData As Variant, ByRef valueOut As Double) As Boolean
    If IsError(valueData) Then Exit Function
    Dim textValue As String

    If IsNumeric(valueData) And VarType(valueData) <> vbString Then
        valueOut = CDbl(valueData)
        SampleDesign_BandTryNumber = True
        Exit Function
    End If

    textValue = Trim$(CStr(valueData))
    If Len(textValue) = 0 Then Exit Function

    If Left$(textValue, 1) = "'" Then textValue = Mid$(textValue, 2)
    textValue = Replace(textValue, ChrW(160), "")
    textValue = Replace(textValue, " ", "")
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, "\", "")
    textValue = Replace(textValue, ChrW$(8361), "")
    textValue = SampleDesign_StripBandNumberUnitSuffix(textValue)

    If Left$(textValue, 1) = "(" And Right$(textValue, 1) = ")" Then
        textValue = "-" & Mid$(textValue, 2, Len(textValue) - 2)
    End If

    Dim percentValue As Boolean
    If Right$(textValue, 1) = "%" Then
        percentValue = True
        textValue = Left$(textValue, Len(textValue) - 1)
    End If

    If IsNumeric(textValue) Then
        valueOut = CDbl(textValue)
        If percentValue Then valueOut = valueOut / 100
        SampleDesign_BandTryNumber = True
    End If
End Function

Private Function SampleDesign_StripBandNumberUnitSuffix(ByVal textValue As String) As String
    Do While Len(textValue) > 0
        Dim lastChar As String
        lastChar = Right$(textValue, 1)
        If (lastChar >= "0" And lastChar <= "9") Or lastChar = "." Or lastChar = "-" Or lastChar = ")" Or lastChar = "%" Then
            Exit Do
        End If
        textValue = Left$(textValue, Len(textValue) - 1)
    Loop

    SampleDesign_StripBandNumberUnitSuffix = textValue
End Function

Private Function SampleDesign_BandFormulaForAddress(ByVal sourceAddress As String, ByVal bands As Collection) As String
    Dim q As String
    q = Chr$(34)

    Dim valueExpr As String
    valueExpr = SampleDesign_BandValueFormula(sourceAddress)

    Dim formulaText As String
    formulaText = "=IF(" & sourceAddress & "=" & q & q & "," & q & q & ",IFERROR(IFS("

    Dim i As Long
    For i = 1 To bands.Count
        If i > 1 Then formulaText = formulaText & ","
        formulaText = formulaText & SampleDesign_BandConditionFormula(valueExpr, bands(i)) & "," & _
                      SampleDesign_FormulaStringLiteral(CStr(bands(i)("label")))
    Next i

    formulaText = formulaText & ")," & SampleDesign_FormulaStringLiteral("미분류") & "))"
    SampleDesign_BandFormulaForAddress = formulaText
End Function

Private Function SampleDesign_BandConditionFormula(ByVal valueExpr As String, ByVal band As Object) As String
    Dim conditionText As String
    Dim clauseCount As Long

    If CBool(band("hasStart")) Then
        clauseCount = clauseCount + 1
        conditionText = valueExpr & IIf(CStr(band("startOp")) = "초과", ">", ">=") & _
                        SampleDesign_FormulaNumberLiteral(CDbl(band("startValue")))
    End If

    If CBool(band("hasEnd")) Then
        If clauseCount > 0 Then conditionText = conditionText & ","
        clauseCount = clauseCount + 1
        conditionText = conditionText & valueExpr & IIf(CStr(band("endOp")) = "이하", "<=", "<") & _
                        SampleDesign_FormulaNumberLiteral(CDbl(band("endValue")))
    End If

    If clauseCount > 1 Then
        SampleDesign_BandConditionFormula = "AND(" & conditionText & ")"
    Else
        SampleDesign_BandConditionFormula = conditionText
    End If
End Function

Private Function SampleDesign_BandValueFormula(ByVal sourceAddress As String) As String
    SampleDesign_BandValueFormula = sourceAddress
End Function

Private Function SampleDesign_FormulaNumberLiteral(ByVal numberValue As Double) As String
    Dim textValue As String
    textValue = Trim$(Str$(numberValue))
    If Left$(textValue, 1) = "." Then textValue = "0" & textValue
    If Left$(textValue, 2) = "-." Then textValue = "-0" & Mid$(textValue, 2)
    SampleDesign_FormulaNumberLiteral = textValue
End Function

Private Function SampleDesign_FormulaStringLiteral(ByVal textValue As String) As String
    Dim q As String
    q = Chr$(34)
    SampleDesign_FormulaStringLiteral = q & Replace(CStr(textValue), q, q & q) & q
End Function

Private Function SampleDesign_NormalizeBandStartOp(ByVal opText As String) As String
    If InStr(1, opText, "초과", vbTextCompare) > 0 Then
        SampleDesign_NormalizeBandStartOp = "초과"
    Else
        SampleDesign_NormalizeBandStartOp = "이상"
    End If
End Function

Private Function SampleDesign_NormalizeBandEndOp(ByVal opText As String) As String
    If InStr(1, opText, "이하", vbTextCompare) > 0 Then
        SampleDesign_NormalizeBandEndOp = "이하"
    Else
        SampleDesign_NormalizeBandEndOp = "미만"
    End If
End Function

Private Function SampleDesign_BandLabelForValue(ByVal numericValue As Double, ByVal bands As Collection) As String
    Dim i As Long
    For i = 1 To bands.Count
        Dim band As Object
        Set band = bands(i)
        If SampleDesign_BandValueMatches(numericValue, band) Then
            SampleDesign_BandLabelForValue = CStr(band("label"))
            Exit Function
        End If
    Next i
End Function

Private Function SampleDesign_BandValueMatches(ByVal numericValue As Double, ByVal band As Object) As Boolean
    Dim okStart As Boolean
    Dim okEnd As Boolean

    If CBool(band("hasStart")) Then
        If CStr(band("startOp")) = "초과" Then
            okStart = (numericValue > CDbl(band("startValue")))
        Else
            okStart = (numericValue >= CDbl(band("startValue")))
        End If
    Else
        okStart = True
    End If

    If CBool(band("hasEnd")) Then
        If CStr(band("endOp")) = "이하" Then
            okEnd = (numericValue <= CDbl(band("endValue")))
        Else
            okEnd = (numericValue < CDbl(band("endValue")))
        End If
    Else
        okEnd = True
    End If

    SampleDesign_BandValueMatches = okStart And okEnd
End Function

Private Function SampleDesign_BandAutoLabel(ByVal band As Object) As String
    Dim labelText As String
    If CBool(band("hasStart")) Then
        labelText = SampleDesign_FormatBandNumber(CDbl(band("startValue"))) & " " & CStr(band("startOp"))
    End If
    If CBool(band("hasEnd")) Then
        If Len(labelText) > 0 Then labelText = labelText & "~"
        labelText = labelText & SampleDesign_FormatBandNumber(CDbl(band("endValue"))) & " " & CStr(band("endOp"))
    End If
    SampleDesign_BandAutoLabel = labelText
End Function

Private Function SampleDesign_FormatBandNumber(ByVal valueData As Double) As String
    If valueData = Fix(valueData) Then
        SampleDesign_FormatBandNumber = Format$(valueData, "#,##0")
    Else
        SampleDesign_FormatBandNumber = Format$(valueData, "#,##0.########")
    End If
End Function

Private Function SampleDesign_LoadBandHeader(ByVal wb As Workbook, ByVal defaultHeader As String) As String
    Dim ws As Worksheet
    Set ws = SampleDesign_FindSheet(wb, SD_BAND_SETTINGS_SHEET)
    If Not ws Is Nothing Then
        SampleDesign_LoadBandHeader = Trim$(CStr(ws.Range("B2").Value))
    End If
    If Len(SampleDesign_LoadBandHeader) = 0 Then SampleDesign_LoadBandHeader = defaultHeader
End Function

Private Function SampleDesign_LoadBandSpecText(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Set ws = SampleDesign_FindSheet(wb, SD_BAND_SETTINGS_SHEET)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 5 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
            If Len(SampleDesign_LoadBandSpecText) > 0 Then SampleDesign_LoadBandSpecText = SampleDesign_LoadBandSpecText & vbLf
            SampleDesign_LoadBandSpecText = SampleDesign_LoadBandSpecText & _
                CStr(ws.Cells(r, 1).Value) & vbTab & _
                CStr(ws.Cells(r, 2).Value) & vbTab & _
                CStr(ws.Cells(r, 3).Value) & vbTab & _
                CStr(ws.Cells(r, 4).Value) & vbTab & _
                CStr(ws.Cells(r, 5).Value) & vbTab & _
                CStr(ws.Cells(r, 6).Value)
        End If
    Next r
End Function

Private Sub SampleDesign_SaveBandSettings(ByVal wb As Workbook, ByVal outputHeader As String, ByVal specText As String)
    Dim ws As Worksheet
    Set ws = SampleDesign_FindSheet(wb, SD_BAND_SETTINGS_SHEET)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SD_BAND_SETTINGS_SHEET
    Else
        ws.Visible = xlSheetVisible
        ws.Cells.Clear
    End If

    ws.Range("A1").Value = "표본설계 구간 설정"
    ws.Range("A2").Value = "결과 헤더"
    ws.Range("B2").Value = outputHeader
    SampleDesign_WriteArray ws, 4, 1, Array("순서", "구간 시작값", "시작 조건", "구간 종료값", "종료 조건", "구간명")

    Dim lines As Variant
    lines = Split(Replace(CStr(specText), vbCrLf, vbLf), vbLf)
    Dim i As Long
    Dim outRow As Long
    outRow = 5
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) > 0 Then
            Dim parts As Variant
            parts = Split(CStr(lines(i)), vbTab)
            Dim c As Long
            Dim maxPart As Long
            maxPart = UBound(parts)
            If maxPart > 5 Then maxPart = 5
            For c = 0 To maxPart
                ws.Cells(outRow, c + 1).Value = parts(c)
            Next c
            outRow = outRow + 1
        End If
    Next i

    ws.Columns("A:F").AutoFit
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function SampleDesign_DefaultQuotaHeaders() As Variant
    SampleDesign_DefaultQuotaHeaders = Array("산업", "대분류", "지역", "조사지역", "대조지역", "존", "성별", "연령")
End Function

Private Function SampleDesign_DefaultQuotaEnabled() As Variant
    SampleDesign_DefaultQuotaEnabled = Array(True, True, True, False, False, False, False, False)
End Function

Private Function SampleDesign_DefaultSetup() As Object
    Dim setup As Object
    Set setup = CreateObject("Scripting.Dictionary")
    setup("inputMode") = "list"
    setup("targetN") = 800
    setup("methodName") = "단순 비례"
    setup("alpha") = 0.5
    setup("minN") = 0
    setup("maxN") = 0
    setup("quotaHeaders") = SampleDesign_DefaultQuotaHeaders()
    setup("quotaEnabled") = SampleDesign_DefaultQuotaEnabled()
    Set SampleDesign_DefaultSetup = setup
End Function

Private Function SampleDesign_GetSetupFromUser() As Object
    On Error GoTo Fallback

    Dim form As Object
    Set form = VBA.UserForms.Add("frmSampleDesignSetup")
    form.Show
    If CBool(form.Cancelled) Then
        Unload form
        Exit Function
    End If

    Dim setup As Object
    Set setup = SampleDesign_DefaultSetup()
    setup("inputMode") = CStr(form.InputMode)
    setup("targetN") = CLng(Val(form.TargetN))
    setup("methodName") = CStr(form.MethodName)
    setup("alpha") = CDbl(Val(form.AlphaValue))
    setup("minN") = CLng(Val(form.MinN))
    setup("maxN") = CLng(Val(form.MaxN))
    setup("quotaHeaders") = SampleDesign_ParsePipeText(form.QuotaHeaders, SampleDesign_DefaultQuotaHeaders())
    setup("quotaEnabled") = SampleDesign_ParsePipeFlags(form.QuotaEnabled, SampleDesign_DefaultQuotaEnabled())
    Unload form

    Set SampleDesign_GetSetupFromUser = setup
    Exit Function

Fallback:
    Set SampleDesign_GetSetupFromUser = SampleDesign_GetSetupFallback()
End Function

Private Function SampleDesign_GetSetupFallback() As Object
    Dim setup As Object
    Set setup = SampleDesign_DefaultSetup()

    Dim templateMode As String
    templateMode = SampleDesign_SelectInputTemplateMode()
    If Len(templateMode) = 0 Then Exit Function
    setup("inputMode") = templateMode

    Dim targetInput As Variant
    targetInput = Application.InputBox("전체 목표 표본수를 입력하세요.", "표본설계 기본 설정", setup("targetN"), Type:=1)
    If VarType(targetInput) = vbBoolean And targetInput = False Then Exit Function
    If CLng(targetInput) > 0 Then setup("targetN") = CLng(targetInput)

    Dim quotaChoice As String
    quotaChoice = Toolkit_CheckListMenu( _
        "표본설계 쿼터 변수", _
        "표본설계에 사용할 쿼터 변수를 선택하세요. 선택하지 않은 변수는 입력 시트에서 미사용 처리됩니다.", _
        "1" & vbTab & "산업" & vbLf & _
        "2" & vbTab & "대분류" & vbLf & _
        "3" & vbTab & "지역" & vbLf & _
        "4" & vbTab & "조사지역" & vbLf & _
        "5" & vbTab & "대조지역" & vbLf & _
        "6" & vbTab & "존" & vbLf & _
        "7" & vbTab & "성별" & vbLf & _
        "8" & vbTab & "연령", _
        "1,2,3")
    If Len(quotaChoice) > 0 Then setup("quotaEnabled") = SampleDesign_EnabledFlagsFromCsv(quotaChoice)

    Set SampleDesign_GetSetupFallback = setup
End Function

Private Function SampleDesign_ParsePipeText(ByVal pipeText As String, ByVal defaults As Variant) As Variant
    Dim parts As Variant
    parts = Split(CStr(pipeText), "|")

    Dim result() As String
    Dim i As Long
    ReDim result(LBound(defaults) To UBound(defaults))
    For i = LBound(defaults) To UBound(defaults)
        result(i) = CStr(defaults(i))
        If i - LBound(defaults) <= UBound(parts) Then
            If Len(Trim$(CStr(parts(i - LBound(defaults))))) > 0 Then result(i) = Trim$(CStr(parts(i - LBound(defaults))))
        End If
    Next i
    SampleDesign_ParsePipeText = result
End Function

Private Function SampleDesign_ParsePipeFlags(ByVal pipeText As String, ByVal defaults As Variant) As Variant
    Dim parts As Variant
    parts = Split(CStr(pipeText), "|")

    Dim result() As Boolean
    Dim i As Long
    ReDim result(LBound(defaults) To UBound(defaults))
    For i = LBound(defaults) To UBound(defaults)
        result(i) = CBool(defaults(i))
        If i - LBound(defaults) <= UBound(parts) Then
            result(i) = (CStr(parts(i - LBound(defaults))) = "1")
        End If
    Next i
    SampleDesign_ParsePipeFlags = result
End Function

Private Function SampleDesign_EnabledFlagsFromCsv(ByVal csvText As String) As Variant
    Dim defaults As Variant
    defaults = SampleDesign_DefaultQuotaEnabled()

    Dim result() As Boolean
    Dim i As Long
    ReDim result(LBound(defaults) To UBound(defaults))
    For i = LBound(result) To UBound(result)
        result(i) = False
    Next i

    Dim parts As Variant
    parts = Split(CStr(csvText), ",")
    For i = LBound(parts) To UBound(parts)
        Dim idx As Long
        idx = CLng(Val(Trim$(CStr(parts(i)))))
        If idx >= 1 And idx <= UBound(result) - LBound(result) + 1 Then result(LBound(result) + idx - 1) = True
    Next i
    SampleDesign_EnabledFlagsFromCsv = result
End Function

Private Function SampleDesign_SelectInputTemplateMode() As String
    Dim choice As String
    choice = Toolkit_ButtonMenu( _
        "표본설계 입력 형태", _
        "생성할 표본설계 입력 형태를 선택하세요.", _
        "1" & vbTab & "리스트형 입력" & vbLf & _
        "2" & vbTab & "집계표형 입력", _
        "1")

    Select Case choice
        Case "2": SampleDesign_SelectInputTemplateMode = "table"
        Case "1": SampleDesign_SelectInputTemplateMode = "list"
        Case Else: SampleDesign_SelectInputTemplateMode = ""
    End Select
End Function

Private Sub SampleDesign_WriteInputTemplate(ByVal ws As Worksheet, ByVal setup As Object)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 입력"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Dim headers As Variant
    Dim enabled As Variant
    headers = setup("quotaHeaders")
    enabled = setup("quotaEnabled")

    ws.Range("A3").Value = "기본 설정"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "전체 목표 표본수"
    ws.Range("B4").Value = CLng(setup("targetN"))
    ws.Range("A5").Value = "배분 방식"
    ws.Range("B5").Value = CStr(setup("methodName"))
    ws.Range("A6").Value = "멱배분 alpha"
    ws.Range("B6").Value = CDbl(setup("alpha"))
    ws.Range("A7").Value = "기본 최소표본"
    ws.Range("B7").Value = CLng(setup("minN"))
    ws.Range("A8").Value = "기본 최대표본"
    ws.Range("B8").Value = CLng(setup("maxN"))
    ws.Range("D3").Value = "필수 입력"
    ws.Range("D4").Value = "1. B4 전체 목표 표본수를 확인합니다."
    ws.Range("D5").Value = "2. 19행 사용여부에서 사용할 쿼터만 '사용'으로 둡니다."
    ws.Range("D6").Value = "3. 21행부터 쿼터값과 모집단수를 입력합니다."
    ws.Range("D7").Value = "4. 그룹별 목표가 있으면 아래 목표표본 그룹표를 채웁니다."
    ws.Range("D8").Value = "5. 입력 후 [할당 계산] 버튼을 누릅니다."

    Dim i As Long
    ws.Range("A11").Value = "목표표본 그룹표"
    ws.Range("A11").Font.Bold = True
    ws.Cells(12, 1).Value = "목표그룹ID"
    ws.Cells(12, 2).Value = "목표표본수"
    For i = LBound(headers) To UBound(headers)
        ws.Cells(12, 3 + i - LBound(headers)).Value = headers(i)
    Next i
    ws.Cells(12, 3 + UBound(headers) - LBound(headers) + 1).Value = "비고"
    SampleDesign_WriteTargetGroupExample ws, 13, headers, "G001", "예: 특정 산업/지역 목표"
    SampleDesign_WriteTargetGroupExample ws, 14, headers, "G002", "예: 다른 조건의 목표"
    SampleDesign_WriteTargetGroupExample ws, 15, headers, "G999", "조건을 비우면 나머지 전체 catch-all"

    ws.Range("A18").Value = "모집단 표"
    ws.Range("A18").Font.Bold = True
    ws.Cells(19, 1).Value = "사용여부"
    For i = LBound(headers) To UBound(headers)
        ws.Cells(19, 2 + i - LBound(headers)).Value = IIf(CBool(enabled(i)), "사용", "미사용")
    Next i

    Dim headerRow As Long
    headerRow = 20
    ws.Cells(headerRow, 1).Value = "쿼터ID"
    For i = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, 2 + i - LBound(headers)).Value = headers(i)
    Next i

    Dim optionStartCol As Long
    optionStartCol = 2 + UBound(headers) - LBound(headers) + 1
    ws.Cells(headerRow, optionStartCol).Value = "모집단수"
    ws.Cells(headerRow, optionStartCol + 1).Value = "우선배정여부"
    ws.Cells(headerRow, optionStartCol + 2).Value = "최소표본"
    ws.Cells(headerRow, optionStartCol + 3).Value = "최대표본"
    ws.Cells(headerRow, optionStartCol + 4).Value = "사용자가중치"
    ws.Cells(headerRow, optionStartCol + 5).Value = "비고"

    Dim exampleRow As Long
    For exampleRow = 21 To 23
        SampleDesign_WritePopulationExample ws, exampleRow, headers, enabled, optionStartCol
    Next exampleRow

    With ws.Range(ws.Cells(12, 1), ws.Cells(15, optionStartCol + 1))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(23, optionStartCol + 5))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    ws.Range(ws.Cells(12, 1), ws.Cells(12, optionStartCol + 1)).Font.Bold = True
    ws.Range(ws.Cells(12, 1), ws.Cells(12, optionStartCol + 1)).Interior.Color = RGB(220, 235, 250)
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, optionStartCol + 5)).Font.Bold = True
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, optionStartCol + 5)).Interior.Color = RGB(220, 235, 250)
    ws.Range("A4:B8").Interior.Color = RGB(255, 242, 204)
    ws.Range("D3:D8").Interior.Color = RGB(234, 244, 255)
    ws.Range("D3:D8").Font.Color = RGB(31, 78, 121)
    SampleDesign_ApplyQuotaColumnFormatting ws, headers, enabled, 2, 12, 23
    SampleDesign_ApplyInputValidation ws
    ws.Columns.AutoFit
End Sub

Private Sub SampleDesign_WriteTargetGroupExample(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal headers As Variant, _
                                                 ByVal groupId As String, ByVal noteText As String)
    ws.Cells(rowIndex, 1).Value = groupId
    ws.Cells(rowIndex, 2).Value = ""

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowIndex, 3 + i - LBound(headers)).Value = ""
    Next i
    ws.Cells(rowIndex, 3 + UBound(headers) - LBound(headers) + 1).Value = noteText
End Sub

Private Sub SampleDesign_WritePopulationExample(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal headers As Variant, _
                                                ByVal enabled As Variant, ByVal optionStartCol As Long)
    Dim exampleIndex As Long
    exampleIndex = rowIndex - 20

    ws.Cells(rowIndex, 1).Value = "Q" & Format$(exampleIndex, "000")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If CBool(enabled(i)) Then
            ws.Cells(rowIndex, 2 + i - LBound(headers)).Value = SampleDesign_ExampleQuotaValue(CStr(headers(i)), exampleIndex)
        Else
            ws.Cells(rowIndex, 2 + i - LBound(headers)).Value = ""
        End If
    Next i
    ws.Cells(rowIndex, optionStartCol).Value = Choose(exampleIndex, 1000, 500, 300)
    If exampleIndex = 1 Then ws.Cells(rowIndex, optionStartCol + 5).Value = "예시 행입니다. 실제 모집단 값으로 교체하세요."
End Sub

Private Function SampleDesign_ExampleQuotaValue(ByVal headerText As String, ByVal exampleIndex As Long) As String
    Select Case True
        Case InStr(1, headerText, "산업", vbTextCompare) > 0
            SampleDesign_ExampleQuotaValue = Choose(exampleIndex, "SW", "SW", "Non-SW")
        Case InStr(1, headerText, "대분류", vbTextCompare) > 0 Or InStr(1, headerText, "업종", vbTextCompare) > 0
            SampleDesign_ExampleQuotaValue = Choose(exampleIndex, "패키지SW", "IT서비스", "제조")
        Case InStr(1, headerText, "지역", vbTextCompare) > 0 Or InStr(1, headerText, "권역", vbTextCompare) > 0
            SampleDesign_ExampleQuotaValue = Choose(exampleIndex, "수도권", "비수도권", "수도권")
        Case InStr(1, headerText, "존", vbTextCompare) > 0
            SampleDesign_ExampleQuotaValue = Choose(exampleIndex, "A존", "B존", "C존")
        Case InStr(1, headerText, "성별", vbTextCompare) > 0
            SampleDesign_ExampleQuotaValue = Choose(exampleIndex, "남성", "여성", "전체")
        Case InStr(1, headerText, "연령", vbTextCompare) > 0 Or InStr(1, headerText, "나이", vbTextCompare) > 0
            SampleDesign_ExampleQuotaValue = Choose(exampleIndex, "20대", "30대", "40대")
        Case Else
            SampleDesign_ExampleQuotaValue = ""
    End Select
End Function

Private Sub SampleDesign_ApplyQuotaColumnFormatting(ByVal ws As Worksheet, ByVal headers As Variant, ByVal enabled As Variant, _
                                                    ByVal firstQuotaCol As Long, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        Dim colIndex As Long
        colIndex = firstQuotaCol + i - LBound(headers)
        With ws.Range(ws.Cells(firstRow, colIndex), ws.Cells(lastRow, colIndex))
            If CBool(enabled(i)) Then
                .Interior.Color = RGB(255, 255, 255)
                .Font.Color = RGB(0, 0, 0)
            Else
                .Interior.Color = RGB(242, 242, 242)
                .Font.Color = RGB(130, 130, 130)
            End If
        End With
    Next i
End Sub

Private Sub SampleDesign_ApplyInputValidation(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="단순 비례,제곱근 비례,멱배분,균등 배분,사용자 지정 가중"
    End With
    With ws.Range("B4:B8").Validation
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

Private Sub SampleDesign_WriteTableInputTemplate(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 집계표입력"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "사용 방법"
    ws.Range("A4").Value = "1. 이 시트에 모집단 구성표를 원본 형태로 붙여넣습니다. 여러 표를 한 번에 붙여넣어도 됩니다."
    ws.Range("A5").Value = "2. [구성표 리스트화] 버튼으로 탐지 결과를 미리 확인할 수 있습니다."
    ws.Range("A6").Value = "3. [할당 계산] 실행 시 표본설계_입력 시트의 모집단 표로 자동 변환을 시도합니다."
    ws.Range("A7").Value = "4. 변환 후 표본설계_입력 시트에서 쿼터명, 목표표본 그룹표, 모집단수를 검토하세요."
    ws.Range("A8").Value = "예시"
    SampleDesign_WriteArray ws, 10, 1, Array("구분", "10인 미만", "10~50인 미만", "50인 이상", "합계")
    SampleDesign_WriteArray ws, 11, 1, Array("패키지SW", 42081, 4213, 82, 46376)
    SampleDesign_WriteArray ws, 12, 1, Array("IT서비스", 20775, 1826, 58, 22659)
    SampleDesign_WriteArray ws, 13, 1, Array("인터넷SW", 7509, 474, 26, 8009)
    With ws.Range("A10:E13")
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    ws.Range("A10:E10").Font.Bold = True
    ws.Range("A10:E10").Interior.Color = RGB(220, 235, 250)
End Sub

Private Sub SampleDesign_WriteEmptyResult(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 할당결과"
    ws.Range("A3").Value = "할당 계산을 실행하면 동적 쿼터 구조의 결과가 생성됩니다."
    ws.Columns("A:C").AutoFit
End Sub

Private Function SampleDesign_FindHeaderRow(ByVal ws As Worksheet, ByVal headerText As String, ByVal maxRows As Long) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    Dim r As Long, c As Long
    For r = 1 To maxRows
        For c = 1 To lastCol + 20
            If Trim$(CStr(ws.Cells(r, c).Value)) = headerText Then
                SampleDesign_FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function SampleDesign_FindHeaderColumnInRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(headerRow, c).Value)) = headerText Then
            SampleDesign_FindHeaderColumnInRow = c
            Exit Function
        End If
    Next c
End Function

Private Function SampleDesign_ReadInputSchema(ByVal ws As Worksheet) As Object
    Dim headerRow As Long
    headerRow = SampleDesign_FindHeaderRow(ws, "모집단수", 160)
    If headerRow = 0 Then Exit Function

    Dim popCol As Long, idCol As Long
    popCol = SampleDesign_FindHeaderColumnInRow(ws, headerRow, "모집단수")
    idCol = SampleDesign_FindHeaderColumnInRow(ws, headerRow, "쿼터ID")
    If idCol = 0 Then idCol = 1
    If popCol <= idCol + 1 Then Exit Function

    Dim quotaCols As Collection
    Dim quotaHeaders As Collection
    Set quotaCols = New Collection
    Set quotaHeaders = New Collection

    Dim c As Long, headerText As String
    For c = idCol + 1 To popCol - 1
        headerText = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If Len(headerText) > 0 And SampleDesign_QuotaColumnEnabledDynamic(ws, headerRow, c) Then
            quotaCols.Add c
            quotaHeaders.Add headerText
        End If
    Next c
    If quotaCols.Count = 0 Then Exit Function

    Dim schema As Object
    Set schema = CreateObject("Scripting.Dictionary")
    schema.Add "headerRow", headerRow
    schema.Add "dataStartRow", headerRow + 1
    schema.Add "idCol", idCol
    schema.Add "popCol", popCol
    schema.Add "priorityCol", SampleDesign_FindHeaderColumnInRow(ws, headerRow, "우선배정여부")
    schema.Add "minCol", SampleDesign_FindHeaderColumnInRow(ws, headerRow, "최소표본")
    schema.Add "maxCol", SampleDesign_FindHeaderColumnInRow(ws, headerRow, "최대표본")
    schema.Add "userWeightCol", SampleDesign_FindHeaderColumnInRow(ws, headerRow, "사용자가중치")
    schema.Add "noteCol", SampleDesign_FindHeaderColumnInRow(ws, headerRow, "비고")
    schema.Add "quotaCols", quotaCols
    schema.Add "quotaHeaders", quotaHeaders
    Set SampleDesign_ReadInputSchema = schema
End Function

Private Function SampleDesign_QuotaColumnEnabledDynamic(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal colIndex As Long) As Boolean
    Dim flagText As String
    flagText = UCase$(Replace(Trim$(CStr(ws.Cells(headerRow - 1, colIndex).Value)), " ", ""))
    SampleDesign_QuotaColumnEnabledDynamic = Not (flagText = "미사용" Or flagText = "N" Or flagText = "NO" Or flagText = "FALSE" Or flagText = "0" Or flagText = "사용안함")
End Function

Private Function SampleDesign_PopulationDataRowCount(ByVal ws As Worksheet) As Long
    Dim schema As Object
    Set schema = SampleDesign_ReadInputSchema(ws)
    If schema Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, CLng(schema("popCol"))).End(xlUp).Row
    Dim r As Long
    For r = CLng(schema("dataStartRow")) To lastRow
        If SampleDesign_CellNumber(ws.Cells(r, CLng(schema("popCol"))).Value) > 0 Then SampleDesign_PopulationDataRowCount = SampleDesign_PopulationDataRowCount + 1
    Next r
End Function

Private Function SampleDesign_CellNumber(ByVal valueData As Variant) As Double
    Dim textValue As String
    textValue = Replace(Trim$(CStr(valueData)), ",", "")
    If IsNumeric(textValue) Then SampleDesign_CellNumber = CDbl(textValue)
End Function

Private Sub SampleDesign_ImportTableInputToPopulation(ByVal inputWs As Worksheet, ByVal tableWs As Worksheet, ByRef issues As Collection)
    Dim usedRange As Range
    Set usedRange = tableWs.UsedRange
    If usedRange.Rows.Count < 3 Or usedRange.Columns.Count < 3 Then Exit Sub

    Dim blocks As Collection
    Set blocks = PopList_DetectBlocks(usedRange)
    If blocks.Count = 0 Then
        issues.Add Array("집계표입력", 0, "표본설계_집계표입력 시트에서 리스트화 가능한 집계표를 찾지 못했습니다.")
        Exit Sub
    End If

    Dim quotaLabels As Collection
    Set quotaLabels = New Collection
    SampleDesign_AddUniqueQuotaLabel quotaLabels, "표제목"

    Dim i As Long, axisIndex As Long
    For i = 1 To blocks.Count
        Dim block As Object
        Set block = blocks(i)
        For axisIndex = 1 To block("rowAxisHeaders").Count
            SampleDesign_AddUniqueQuotaLabel quotaLabels, PopList_CollectionText(block("rowAxisHeaders"), axisIndex)
        Next axisIndex
    Next i
    SampleDesign_AddUniqueQuotaLabel quotaLabels, "가로구분"

    Dim headerRow As Long
    headerRow = SampleDesign_FindHeaderRow(inputWs, "모집단수", 160)
    If headerRow = 0 Then Exit Sub

    Dim popCol As Long
    popCol = SampleDesign_FindHeaderColumnInRow(inputWs, headerRow, "모집단수")
    If popCol = 0 Then Exit Sub

    Dim q As Long
    For q = 1 To popCol - 2
        If q <= quotaLabels.Count Then
            inputWs.Cells(headerRow - 1, 1 + q).Value = "사용"
            inputWs.Cells(headerRow, 1 + q).Value = quotaLabels(q)
        Else
            inputWs.Cells(headerRow - 1, 1 + q).Value = "미사용"
            inputWs.Cells(headerRow, 1 + q).Value = "쿼터" & q
        End If
    Next q

    inputWs.Range(inputWs.Cells(headerRow + 1, 1), inputWs.Cells(inputWs.Rows.Count, popCol + 5)).ClearContents

    Dim outRow As Long
    outRow = headerRow + 1
    For i = 1 To blocks.Count
        Set block = blocks(i)
        Dim lastSeen() As String
        ReDim lastSeen(1 To CLng(block("rowAxisCount")))

        Dim relRow As Long
        For relRow = CLng(block("dataStartRelRow")) To CLng(block("dataEndRelRow"))
            If Not PopList_RowIsTotalByCurrentCells(usedRange, relRow, CLng(block("rowAxisCount"))) Then
                PopList_UpdateLastSeen usedRange, relRow, CLng(block("rowAxisCount")), lastSeen

                Dim valueCols As Collection
                Set valueCols = block("valueCols")
                Dim valueIndex As Long
                For valueIndex = 1 To valueCols.Count
                    Dim relCol As Long
                    relCol = CLng(valueCols(valueIndex))
                    If PopList_IsNumberLike(usedRange.Cells(relRow, relCol).Value) Then
                        inputWs.Cells(outRow, 1).Value = "P" & Format$(outRow - headerRow, "0000")
                        inputWs.Cells(outRow, 2).Value = CStr(block("title"))
                        Dim columnAxisTargetCol As Long
                        columnAxisTargetCol = 1 + quotaLabels.Count
                        If columnAxisTargetCol > popCol - 1 Then columnAxisTargetCol = popCol - 1
                        For axisIndex = 1 To CLng(block("rowAxisCount"))
                            If 2 + axisIndex < columnAxisTargetCol Then inputWs.Cells(outRow, 2 + axisIndex).Value = lastSeen(axisIndex)
                        Next axisIndex
                        If columnAxisTargetCol >= 2 Then inputWs.Cells(outRow, columnAxisTargetCol).Value = PopList_CellText(usedRange.Cells(CLng(block("headerRelRow")), relCol))
                        inputWs.Cells(outRow, popCol).Value = PopList_ToNumber(usedRange.Cells(relRow, relCol).Value)
                        outRow = outRow + 1
                    End If
                Next valueIndex
            End If
        Next relRow
    Next i

    issues.Add Array("집계표입력", 0, blocks.Count & "개 집계표를 표본설계_입력 모집단 표로 자동 변환했습니다.")
End Sub

Private Sub SampleDesign_AddUniqueQuotaLabel(ByVal labels As Collection, ByVal labelText As String)
    labelText = Trim$(labelText)
    If Len(labelText) = 0 Then Exit Sub
    Dim item As Variant
    For Each item In labels
        If StrComp(CStr(item), labelText, vbTextCompare) = 0 Then Exit Sub
    Next item
    If labels.Count < 8 Then labels.Add labelText
End Sub

Private Function SampleDesign_ReadTargetGroups(ByVal ws As Worksheet, ByVal schema As Object, ByVal defaultTargetN As Long, ByRef issues As Collection) As Collection
    Dim groups As New Collection
    Dim groupHeaderRow As Long
    groupHeaderRow = SampleDesign_FindHeaderRow(ws, "목표그룹ID", CLng(schema("headerRow")) - 1)

    If groupHeaderRow > 0 Then
        Dim idCol As Long, targetCol As Long
        idCol = SampleDesign_FindHeaderColumnInRow(ws, groupHeaderRow, "목표그룹ID")
        targetCol = SampleDesign_FindHeaderColumnInRow(ws, groupHeaderRow, "목표표본수")
        If idCol > 0 And targetCol > 0 Then
            Dim r As Long, lastRow As Long
            lastRow = CLng(schema("headerRow")) - 2
            For r = groupHeaderRow + 1 To lastRow
                Dim targetN As Long
                targetN = CLng(SampleDesign_CellNumber(ws.Cells(r, targetCol).Value))
                If targetN > 0 Then
                    Dim group As Object
                    Set group = CreateObject("Scripting.Dictionary")
                    group.Add "id", Trim$(CStr(ws.Cells(r, idCol).Value))
                    If Len(group("id")) = 0 Then group("id") = "G" & Format$(groups.Count + 1, "000")
                    group.Add "target", targetN
                    group.Add "sourceRow", r
                    group.Add "order", groups.Count + 1
                    Dim conditions As Object
                    Set conditions = CreateObject("Scripting.Dictionary")

                    Dim q As Long
                    For q = 1 To schema("quotaHeaders").Count
                        Dim quotaName As String
                        quotaName = CStr(schema("quotaHeaders")(q))
                        Dim condCol As Long
                        condCol = SampleDesign_FindHeaderColumnInRow(ws, groupHeaderRow, quotaName)
                        If condCol > 0 Then
                            Dim condValue As String
                            condValue = Trim$(CStr(ws.Cells(r, condCol).Value))
                            If Len(condValue) > 0 Then conditions.Add quotaName, condValue
                        End If
                    Next q
                    group.Add "conditions", conditions
                    group.Add "conditionCount", conditions.Count
                    groups.Add group
                End If
            Next r
        End If
    End If

    If groups.Count = 0 And defaultTargetN > 0 Then
        Dim defaultGroup As Object
        Set defaultGroup = CreateObject("Scripting.Dictionary")
        defaultGroup.Add "id", "전체"
        defaultGroup.Add "target", defaultTargetN
        defaultGroup.Add "sourceRow", 4
        defaultGroup.Add "order", 1
        Dim emptyConditions As Object
        Set emptyConditions = CreateObject("Scripting.Dictionary")
        defaultGroup.Add "conditions", emptyConditions
        defaultGroup.Add "conditionCount", 0
        groups.Add defaultGroup
    End If

    Set SampleDesign_ReadTargetGroups = groups
End Function

Private Sub SampleDesign_ReadPopulation(ByVal ws As Worksheet, ByVal schema As Object, ByVal targetGroups As Collection, ByRef records As Collection, ByRef issues As Collection)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, CLng(schema("popCol"))).End(xlUp).Row
    If lastRow < CLng(schema("dataStartRow")) Then Exit Sub

    Dim r As Long
    For r = CLng(schema("dataStartRow")) To lastRow
        Dim popN As Double
        popN = SampleDesign_CellNumber(ws.Cells(r, CLng(schema("popCol"))).Value)
        If SampleDesign_RowIsBlank(ws, r, CLng(schema("idCol")), CLng(schema("popCol"))) Then
        ElseIf popN <= 0 Then
            issues.Add Array("제외", r, "모집단수가 0 이하이거나 숫자가 아닙니다.")
        Else
            Dim quotaValues As Object
            Set quotaValues = CreateObject("Scripting.Dictionary")
            Dim q As Long
            For q = 1 To schema("quotaHeaders").Count
                quotaValues.Add CStr(schema("quotaHeaders")(q)), Trim$(CStr(ws.Cells(r, CLng(schema("quotaCols")(q))).Value))
            Next q

            Dim group As Object
            Set group = SampleDesign_FindTargetGroupForRecord(quotaValues, targetGroups, r, issues)
            If group Is Nothing Then
                issues.Add Array("목표그룹없음", r, "어떤 목표표본 그룹에도 매칭되지 않아 할당에서 제외했습니다.")
            Else
                Dim key As String
                key = CStr(group("id")) & "|" & SampleDesign_BuildQuotaKeyDynamic(quotaValues, schema("quotaHeaders"))
                If dict.Exists(key) Then
                    Dim existing As Object
                    Set existing = dict(key)
                    existing("pop") = CDbl(existing("pop")) + popN
                    existing("sourceRows") = CStr(existing("sourceRows")) & ", " & r
                    issues.Add Array("중복합산", r, "같은 목표그룹/쿼터 조합을 기존 행에 합산했습니다.")
                Else
                    Dim rec As Object
                    Set rec = CreateObject("Scripting.Dictionary")
                    rec.Add "id", Trim$(CStr(ws.Cells(r, CLng(schema("idCol"))).Value))
                    If Len(rec("id")) = 0 Then rec("id") = "Q" & Format$(records.Count + dict.Count + 1, "000")
                    rec.Add "quotaValues", quotaValues
                    rec.Add "pop", popN
                    rec.Add "priority", SampleDesign_CellTextIfCol(ws, r, CLng(schema("priorityCol")))
                    rec.Add "min", CLng(SampleDesign_CellNumber(SampleDesign_CellTextIfCol(ws, r, CLng(schema("minCol")))))
                    rec.Add "max", CLng(SampleDesign_CellNumber(SampleDesign_CellTextIfCol(ws, r, CLng(schema("maxCol")))))
                    rec.Add "userWeight", SampleDesign_CellNumber(SampleDesign_CellTextIfCol(ws, r, CLng(schema("userWeightCol"))))
                    rec.Add "note", SampleDesign_CellTextIfCol(ws, r, CLng(schema("noteCol")))
                    rec.Add "sourceRows", CStr(r)
                    rec.Add "targetGroupId", CStr(group("id"))
                    rec.Add "groupTarget", CLng(group("target"))
                    dict.Add key, rec
                End If
            End If
        End If
    Next r

    Dim item As Variant
    For Each item In dict.Items
        records.Add item
    Next item
End Sub

Private Function SampleDesign_RowIsBlank(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long
    For c = firstCol To lastCol
        If Len(Trim$(CStr(ws.Cells(rowIndex, c).Value))) > 0 Then Exit Function
    Next c
    SampleDesign_RowIsBlank = True
End Function

Private Function SampleDesign_CellTextIfCol(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    If colIndex > 0 Then SampleDesign_CellTextIfCol = Trim$(CStr(ws.Cells(rowIndex, colIndex).Value))
End Function

Private Function SampleDesign_BuildQuotaKeyDynamic(ByVal quotaValues As Object, ByVal quotaHeaders As Collection) As String
    Dim q As Long
    For q = 1 To quotaHeaders.Count
        SampleDesign_BuildQuotaKeyDynamic = SampleDesign_BuildQuotaKeyDynamic & "|" & CStr(quotaHeaders(q)) & "=" & CStr(quotaValues(CStr(quotaHeaders(q))))
    Next q
End Function

Private Function SampleDesign_FindTargetGroupForRecord(ByVal quotaValues As Object, ByVal targetGroups As Collection, ByVal sourceRow As Long, ByRef issues As Collection) As Object
    Dim bestGroup As Object
    Dim bestSpecificity As Long
    Dim tied As Boolean
    Dim i As Long
    For i = 1 To targetGroups.Count
        Dim group As Object
        Set group = targetGroups(i)
        If SampleDesign_GroupMatchesRecord(quotaValues, group("conditions")) Then
            If bestGroup Is Nothing Or CLng(group("conditionCount")) > bestSpecificity Then
                Set bestGroup = group
                bestSpecificity = CLng(group("conditionCount"))
                tied = False
            ElseIf CLng(group("conditionCount")) = bestSpecificity Then
                tied = True
            End If
        End If
    Next i

    If tied And Not bestGroup Is Nothing Then
        issues.Add Array("목표그룹중복", sourceRow, "조건 특이도가 같은 여러 목표그룹에 매칭되어 먼저 나온 '" & CStr(bestGroup("id")) & "' 그룹을 적용했습니다.")
    End If
    Set SampleDesign_FindTargetGroupForRecord = bestGroup
End Function

Private Function SampleDesign_GroupMatchesRecord(ByVal quotaValues As Object, ByVal conditions As Object) As Boolean
    Dim key As Variant
    For Each key In conditions.Keys
        If Not quotaValues.Exists(CStr(key)) Then Exit Function
        If StrComp(Trim$(CStr(quotaValues(CStr(key)))), Trim$(CStr(conditions(CStr(key)))), vbTextCompare) <> 0 Then Exit Function
    Next key
    SampleDesign_GroupMatchesRecord = True
End Function

Private Function SampleDesign_Allocate(ByVal records As Collection, ByVal targetGroups As Collection, _
                                       ByVal methodName As String, ByVal alpha As Double, _
                                       ByVal defaultMin As Long, ByVal defaultMax As Long, _
                                       ByRef issues As Collection) As Collection
    Dim result As New Collection
    Dim groupBuckets As Object
    Set groupBuckets = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To records.Count
        Dim rec As Object
        Set rec = records(i)
        Dim gid As String
        gid = CStr(rec("targetGroupId"))
        If Not groupBuckets.Exists(gid) Then
            Dim bucket As Collection
            Set bucket = New Collection
            groupBuckets.Add gid, bucket
        End If
        Dim bucketRef As Collection
        Set bucketRef = groupBuckets(gid)
        bucketRef.Add rec
    Next i

    For i = 1 To targetGroups.Count
        Dim group As Object
        Set group = targetGroups(i)
        gid = CStr(group("id"))
        If groupBuckets.Exists(gid) Then
            SampleDesign_AllocateGroup groupBuckets(gid), group, methodName, alpha, defaultMin, defaultMax, issues, result
        Else
            issues.Add Array("목표그룹빈값", CLng(group("sourceRow")), "목표그룹 '" & gid & "'에 매칭된 모집단 행이 없습니다.")
        End If
    Next i

    Set SampleDesign_Allocate = result
End Function

Private Sub SampleDesign_AllocateGroup(ByVal records As Collection, ByVal group As Object, _
                                       ByVal methodName As String, ByVal alpha As Double, _
                                       ByVal defaultMin As Long, ByVal defaultMax As Long, _
                                       ByRef issues As Collection, ByRef result As Collection)
    Dim n As Long
    n = records.Count
    If n = 0 Then Exit Sub

    Dim weights() As Double, raws() As Double, residuals() As Double, finals() As Long
    ReDim weights(1 To n)
    ReDim raws(1 To n)
    ReDim residuals(1 To n)
    ReDim finals(1 To n)

    Dim i As Long, totalPop As Double, totalWeight As Double
    For i = 1 To n
        Dim rec As Object
        Set rec = records(i)
        totalPop = totalPop + CDbl(rec("pop"))
        weights(i) = SampleDesign_RecordWeightDynamic(rec, methodName, alpha)
        totalWeight = totalWeight + weights(i)
    Next i

    If totalWeight <= 0 Then
        issues.Add Array("오류", CLng(group("sourceRow")), "목표그룹 '" & CStr(group("id")) & "'의 배분 가중치 합계가 0입니다.")
        Exit Sub
    End If

    Dim targetN As Long
    targetN = CLng(group("target"))
    Dim assigned As Long
    For i = 1 To n
        raws(i) = targetN * weights(i) / totalWeight
        finals(i) = Fix(raws(i))
        residuals(i) = raws(i) - finals(i)
        assigned = assigned + finals(i)
    Next i

    Dim remain As Long
    remain = targetN - assigned
    Do While remain > 0
        Dim addIndex As Long
        addIndex = SampleDesign_LargestResidualIndexDynamic(residuals, finals, records, defaultMax)
        If addIndex = 0 Then Exit Do
        finals(addIndex) = finals(addIndex) + 1
        residuals(addIndex) = -1
        remain = remain - 1
    Loop

    For i = 1 To n
        Set rec = records(i)
        Dim minN As Long, maxN As Long
        minN = SampleDesign_RecordMin(rec, defaultMin)
        maxN = SampleDesign_RecordMax(rec, defaultMax)
        If minN > 0 And finals(i) < minN Then
            issues.Add Array("최소보정", CStr(rec("sourceRows")), "목표그룹 '" & CStr(group("id")) & "'에서 최소표본 " & minN & " 적용")
            finals(i) = minN
        End If
        If maxN > 0 And finals(i) > maxN Then
            issues.Add Array("최대보정", CStr(rec("sourceRows")), "목표그룹 '" & CStr(group("id")) & "'에서 최대표본 " & maxN & " 적용")
            finals(i) = maxN
        End If
    Next i

    SampleDesign_AdjustGroupToTarget finals, raws, records, targetN, defaultMin, defaultMax, CStr(group("id")), issues

    For i = 1 To n
        Set rec = records(i)
        Dim outRec As Object
        Set outRec = CreateObject("Scripting.Dictionary")
        outRec.Add "source", rec
        outRec.Add "totalPop", totalPop
        outRec.Add "weight", weights(i)
        outRec.Add "totalWeight", totalWeight
        outRec.Add "raw", raws(i)
        outRec.Add "residual", raws(i) - Fix(raws(i))
        outRec.Add "final", finals(i)
        outRec.Add "constraint", SampleDesign_ConstraintLabelDynamic(rec, finals(i), defaultMin, defaultMax)
        result.Add outRec
    Next i
End Sub

Private Function SampleDesign_RecordWeightDynamic(ByVal rec As Object, ByVal methodName As String, ByVal alpha As Double) As Double
    Dim popN As Double
    popN = CDbl(rec("pop"))
    Select Case methodName
        Case "제곱근 비례": SampleDesign_RecordWeightDynamic = Sqr(popN)
        Case "멱배분": SampleDesign_RecordWeightDynamic = popN ^ alpha
        Case "균등 배분": SampleDesign_RecordWeightDynamic = 1
        Case "사용자 지정 가중"
            If CDbl(rec("userWeight")) > 0 Then
                SampleDesign_RecordWeightDynamic = CDbl(rec("userWeight"))
            Else
                SampleDesign_RecordWeightDynamic = popN
            End If
        Case Else: SampleDesign_RecordWeightDynamic = popN
    End Select
End Function

Private Function SampleDesign_LargestResidualIndexDynamic(ByRef residuals() As Double, ByRef finalN() As Long, ByVal records As Collection, ByVal defaultMax As Long) As Long
    Dim bestIdx As Long, bestVal As Double, i As Long
    bestVal = -2
    For i = LBound(residuals) To UBound(residuals)
        Dim rec As Object
        Set rec = records(i)
        Dim maxN As Long
        maxN = SampleDesign_RecordMax(rec, defaultMax)
        If (maxN <= 0 Or finalN(i) < maxN) And residuals(i) > bestVal Then
            bestVal = residuals(i)
            bestIdx = i
        End If
    Next i
    SampleDesign_LargestResidualIndexDynamic = bestIdx
End Function

Private Sub SampleDesign_AdjustGroupToTarget(ByRef finalN() As Long, ByRef rawN() As Double, ByVal records As Collection, _
                                             ByVal targetN As Long, ByVal defaultMin As Long, ByVal defaultMax As Long, _
                                             ByVal groupId As String, ByRef issues As Collection)
    Dim currentTotal As Long
    currentTotal = SampleDesign_ArrayTotal(finalN)
    Dim guard As Long
    Do While currentTotal < targetN And guard < 100000
        Dim addIdx As Long
        addIdx = SampleDesign_AdjustAddIndexDynamic(finalN, rawN, records, defaultMax)
        If addIdx = 0 Then Exit Do
        finalN(addIdx) = finalN(addIdx) + 1
        currentTotal = currentTotal + 1
        guard = guard + 1
    Loop
    Do While currentTotal > targetN And guard < 100000
        Dim subIdx As Long
        subIdx = SampleDesign_AdjustSubtractIndexDynamic(finalN, rawN, records, defaultMin)
        If subIdx = 0 Then Exit Do
        finalN(subIdx) = finalN(subIdx) - 1
        currentTotal = currentTotal - 1
        guard = guard + 1
    Loop
    If currentTotal <> targetN Then
        issues.Add Array("합계불일치", 0, "목표그룹 '" & groupId & "'의 최종합계가 목표 표본 수와 다릅니다. 최종합계=" & currentTotal & ", 목표=" & targetN)
    End If
End Sub

Private Function SampleDesign_AdjustAddIndexDynamic(ByRef finalN() As Long, ByRef rawN() As Double, ByVal records As Collection, ByVal defaultMax As Long) As Long
    Dim bestIdx As Long, bestGap As Double, i As Long
    bestGap = -999999
    For i = LBound(finalN) To UBound(finalN)
        Dim rec As Object
        Set rec = records(i)
        Dim maxN As Long
        maxN = SampleDesign_RecordMax(rec, defaultMax)
        If maxN <= 0 Or finalN(i) < maxN Then
            If rawN(i) - finalN(i) > bestGap Then
                bestGap = rawN(i) - finalN(i)
                bestIdx = i
            End If
        End If
    Next i
    SampleDesign_AdjustAddIndexDynamic = bestIdx
End Function

Private Function SampleDesign_AdjustSubtractIndexDynamic(ByRef finalN() As Long, ByRef rawN() As Double, ByVal records As Collection, ByVal defaultMin As Long) As Long
    Dim bestIdx As Long, bestGap As Double, i As Long
    bestGap = -999999
    For i = LBound(finalN) To UBound(finalN)
        Dim rec As Object
        Set rec = records(i)
        Dim minN As Long
        minN = SampleDesign_RecordMin(rec, defaultMin)
        If finalN(i) > minN And finalN(i) > 0 Then
            If finalN(i) - rawN(i) > bestGap Then
                bestGap = finalN(i) - rawN(i)
                bestIdx = i
            End If
        End If
    Next i
    SampleDesign_AdjustSubtractIndexDynamic = bestIdx
End Function

Private Function SampleDesign_RecordMin(ByVal rec As Object, ByVal defaultMin As Long) As Long
    SampleDesign_RecordMin = CLng(rec("min"))
    If SampleDesign_RecordMin <= 0 Then SampleDesign_RecordMin = defaultMin
End Function

Private Function SampleDesign_RecordMax(ByVal rec As Object, ByVal defaultMax As Long) As Long
    SampleDesign_RecordMax = CLng(rec("max"))
    If SampleDesign_RecordMax <= 0 Then SampleDesign_RecordMax = defaultMax
End Function

Private Function SampleDesign_ConstraintLabelDynamic(ByVal rec As Object, ByVal finalN As Long, ByVal defaultMin As Long, ByVal defaultMax As Long) As String
    Dim minN As Long, maxN As Long
    minN = SampleDesign_RecordMin(rec, defaultMin)
    maxN = SampleDesign_RecordMax(rec, defaultMax)
    If minN > 0 And finalN = minN Then
        SampleDesign_ConstraintLabelDynamic = "최소 적용"
    ElseIf maxN > 0 And finalN = maxN Then
        SampleDesign_ConstraintLabelDynamic = "최대 적용"
    End If
End Function

Private Function SampleDesign_TotalTargetN(ByVal targetGroups As Collection) As Long
    Dim i As Long
    For i = 1 To targetGroups.Count
        Dim group As Object
        Set group = targetGroups(i)
        SampleDesign_TotalTargetN = SampleDesign_TotalTargetN + CLng(group("target"))
    Next i
End Function

Private Sub SampleDesign_WriteAllocationResult(ByVal ws As Worksheet, ByVal allocation As Collection, ByVal targetN As Long, _
                                               ByVal methodName As String, ByVal alpha As Double, ByVal quotaHeaders As Collection)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 할당결과"
    ws.Range("B1").Value = "배분 방식": ws.Range("C1").Value = methodName
    ws.Range("D1").Value = "목표 표본 수": ws.Range("E1").Value = targetN

    Dim headerRow As Long
    headerRow = 3
    Dim col As Long
    col = 1
    ws.Cells(headerRow, col).Value = "쿼터ID": col = col + 1
    Dim q As Long
    For q = 1 To quotaHeaders.Count
        ws.Cells(headerRow, col).Value = quotaHeaders(q)
        col = col + 1
    Next q

    Dim popCol As Long, popRateCol As Long, groupCol As Long, groupTargetCol As Long, methodCol As Long
    Dim baseCol As Long, weightCol As Long, weightSumCol As Long, rawCol As Long, residualCol As Long
    Dim finalCol As Long, sampleRateCol As Long, userWeightCol As Long, constraintCol As Long, noteCol As Long
    popCol = col: ws.Cells(headerRow, col).Value = "모집단수": col = col + 1
    popRateCol = col: ws.Cells(headerRow, col).Value = "모집단비율": col = col + 1
    groupCol = col: ws.Cells(headerRow, col).Value = "목표그룹ID": col = col + 1
    groupTargetCol = col: ws.Cells(headerRow, col).Value = "그룹목표표본수": col = col + 1
    methodCol = col: ws.Cells(headerRow, col).Value = "배분방식": col = col + 1
    baseCol = col: ws.Cells(headerRow, col).Value = "배분기준값": col = col + 1
    weightCol = col: ws.Cells(headerRow, col).Value = "배분가중치": col = col + 1
    weightSumCol = col: ws.Cells(headerRow, col).Value = "그룹가중치합": col = col + 1
    rawCol = col: ws.Cells(headerRow, col).Value = "이론할당": col = col + 1
    residualCol = col: ws.Cells(headerRow, col).Value = "잔차": col = col + 1
    finalCol = col: ws.Cells(headerRow, col).Value = "최종할당": col = col + 1
    sampleRateCol = col: ws.Cells(headerRow, col).Value = "표본비율": col = col + 1
    userWeightCol = col: ws.Cells(headerRow, col).Value = "사용자가중치": col = col + 1
    constraintCol = col: ws.Cells(headerRow, col).Value = "제약 적용": col = col + 1
    noteCol = col: ws.Cells(headerRow, col).Value = "비고"

    Dim firstDataRow As Long, lastDataRow As Long
    firstDataRow = headerRow + 1
    lastDataRow = headerRow + allocation.Count

    Dim r As Long
    For r = 1 To allocation.Count
        Dim outRec As Object, rec As Object
        Set outRec = allocation(r)
        Set rec = outRec("source")
        Dim rowNo As Long
        rowNo = headerRow + r
        col = 1
        ws.Cells(rowNo, col).Value = rec("id"): col = col + 1
        For q = 1 To quotaHeaders.Count
            ws.Cells(rowNo, col).Value = rec("quotaValues")(CStr(quotaHeaders(q)))
            col = col + 1
        Next q
        ws.Cells(rowNo, popCol).Value = CDbl(rec("pop"))
        ws.Cells(rowNo, groupCol).Value = CStr(rec("targetGroupId"))
        ws.Cells(rowNo, groupTargetCol).Value = CLng(rec("groupTarget"))
        ws.Cells(rowNo, methodCol).Value = methodName
        ws.Cells(rowNo, finalCol).Value = CLng(outRec("final"))
        ws.Cells(rowNo, userWeightCol).Value = CDbl(rec("userWeight"))
        ws.Cells(rowNo, constraintCol).Value = CStr(outRec("constraint"))
        ws.Cells(rowNo, noteCol).Value = CStr(rec("note"))
    Next r

    If allocation.Count > 0 Then
        Dim popRange As String, finalRange As String, groupRange As String, weightRange As String
        popRange = ws.Range(ws.Cells(firstDataRow, popCol), ws.Cells(lastDataRow, popCol)).Address(False, False)
        finalRange = ws.Range(ws.Cells(firstDataRow, finalCol), ws.Cells(lastDataRow, finalCol)).Address(False, False)
        groupRange = ws.Range(ws.Cells(firstDataRow, groupCol), ws.Cells(lastDataRow, groupCol)).Address(False, False)
        weightRange = ws.Range(ws.Cells(firstDataRow, weightCol), ws.Cells(lastDataRow, weightCol)).Address(False, False)

        For r = firstDataRow To lastDataRow
            Dim popCell As String, baseCell As String, weightCell As String, weightSumCell As String, groupCell As String, groupTargetCell As String, rawCell As String, finalCell As String, userWeightCell As String
            popCell = ws.Cells(r, popCol).Address(False, False)
            baseCell = ws.Cells(r, baseCol).Address(False, False)
            weightCell = ws.Cells(r, weightCol).Address(False, False)
            weightSumCell = ws.Cells(r, weightSumCol).Address(False, False)
            groupCell = ws.Cells(r, groupCol).Address(False, False)
            groupTargetCell = ws.Cells(r, groupTargetCol).Address(False, False)
            rawCell = ws.Cells(r, rawCol).Address(False, False)
            finalCell = ws.Cells(r, finalCol).Address(False, False)
            userWeightCell = ws.Cells(r, userWeightCol).Address(False, False)

            Select Case methodName
                Case "균등 배분"
                    ws.Cells(r, baseCol).Formula = "=1"
                    ws.Cells(r, weightCol).Formula = "=1"
                Case "사용자 지정 가중"
                    ws.Cells(r, baseCol).Formula = "=IF(" & userWeightCell & ">0," & userWeightCell & "," & popCell & ")"
                    ws.Cells(r, weightCol).Formula = "=" & baseCell
                Case "제곱근 비례"
                    ws.Cells(r, baseCol).Formula = "=" & popCell
                    ws.Cells(r, weightCol).Formula = "=SQRT(" & popCell & ")"
                Case "멱배분"
                    ws.Cells(r, baseCol).Formula = "=" & popCell
                    ws.Cells(r, weightCol).Formula = "=" & popCell & "^" & Replace(CStr(alpha), ",", ".")
                Case Else
                    ws.Cells(r, baseCol).Formula = "=" & popCell
                    ws.Cells(r, weightCol).Formula = "=" & popCell
            End Select
            ws.Cells(r, popRateCol).Formula = "=" & popCell & "/SUM(" & popRange & ")"
            ws.Cells(r, weightSumCol).Formula = "=SUMIF(" & groupRange & "," & groupCell & "," & weightRange & ")"
            ws.Cells(r, rawCol).Formula = "=IF(" & weightSumCell & ">0," & groupTargetCell & "*" & weightCell & "/" & weightSumCell & ",0)"
            ws.Cells(r, residualCol).Formula = "=" & rawCell & "-INT(" & rawCell & ")"
            ws.Cells(r, sampleRateCol).Formula = "=" & finalCell & "/SUM(" & finalRange & ")"
        Next r
    End If

    Dim totalRow As Long
    totalRow = lastDataRow + 1
    ws.Cells(totalRow, 1).Value = "총계"
    If allocation.Count > 0 Then
        ws.Cells(totalRow, popCol).Formula = "=SUM(" & ws.Range(ws.Cells(firstDataRow, popCol), ws.Cells(lastDataRow, popCol)).Address(False, False) & ")"
        ws.Cells(totalRow, groupTargetCol).Value = targetN
        ws.Cells(totalRow, finalCol).Formula = "=SUM(" & ws.Range(ws.Cells(firstDataRow, finalCol), ws.Cells(lastDataRow, finalCol)).Address(False, False) & ")"
    End If
    ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, noteCol)).Font.Bold = True

    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(totalRow, noteCol))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, noteCol)).Font.Bold = True
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, noteCol)).Interior.Color = RGB(220, 235, 250)
    ws.Range(ws.Cells(firstDataRow, popRateCol), ws.Cells(totalRow, popRateCol)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(firstDataRow, sampleRateCol), ws.Cells(totalRow, sampleRateCol)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(firstDataRow, baseCol), ws.Cells(totalRow, rawCol)).NumberFormat = "0.000"
    SampleDesign_ApplyDynamicResultFormatting ws, firstDataRow, lastDataRow, finalCol, constraintCol
End Sub

Private Sub SampleDesign_ApplyDynamicResultFormatting(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal finalCol As Long, ByVal constraintCol As Long)
    If lastRow < firstRow Then Exit Sub
    On Error Resume Next
    With ws.Range(ws.Cells(firstRow, finalCol), ws.Cells(lastRow, finalCol)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
        .Interior.Color = RGB(255, 230, 230)
        .Font.Color = RGB(160, 0, 0)
    End With
    With ws.Range(ws.Cells(firstRow, constraintCol), ws.Cells(lastRow, constraintCol)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=" & Chr$(34) & Chr$(34))
        .Interior.Color = RGB(255, 245, 190)
    End With
    On Error GoTo 0
End Sub

Private Sub SampleDesign_WriteReviewReport(ByVal ws As Worksheet, ByVal allocation As Collection, ByVal targetGroups As Collection, ByVal issues As Collection, _
                                           ByVal targetN As Long, ByVal methodName As String, ByVal alpha As Double)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 검토"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Dim totalPop As Double, totalFinal As Long, zeroCount As Long, i As Long
    For i = 1 To allocation.Count
        Dim outRec As Object, rec As Object
        Set outRec = allocation(i)
        Set rec = outRec("source")
        totalPop = totalPop + CDbl(rec("pop"))
        totalFinal = totalFinal + CLng(outRec("final"))
        If CLng(outRec("final")) = 0 Then zeroCount = zeroCount + 1
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

    rowNo = rowNo + 1
    ws.Cells(rowNo, 1).Value = "목표그룹 검토"
    ws.Cells(rowNo, 1).Font.Bold = True
    rowNo = rowNo + 1
    SampleDesign_WriteArray ws, rowNo, 1, Array("목표그룹ID", "목표표본수", "최종할당", "차이", "행수")
    ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 5)).Font.Bold = True
    rowNo = rowNo + 1

    For i = 1 To targetGroups.Count
        Dim group As Object
        Set group = targetGroups(i)
        Dim groupFinal As Long, groupRows As Long
        SampleDesign_GroupAllocationSummary allocation, CStr(group("id")), groupFinal, groupRows
        SampleDesign_WriteArray ws, rowNo, 1, Array(CStr(group("id")), CLng(group("target")), groupFinal, groupFinal - CLng(group("target")), groupRows)
        rowNo = rowNo + 1
    Next i

    rowNo = rowNo + 2
    ws.Cells(rowNo, 1).Value = "검토 항목"
    ws.Cells(rowNo, 1).Font.Bold = True
    rowNo = rowNo + 1
    SampleDesign_WriteArray ws, rowNo, 1, Array("유형", "원본행", "내용")
    ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 3)).Font.Bold = True
    rowNo = rowNo + 1
    For i = 1 To issues.Count
        SampleDesign_WriteArray ws, rowNo, 1, issues(i)
        rowNo = rowNo + 1
    Next i
    If issues.Count = 0 Then ws.Cells(rowNo, 1).Value = "검토 경고 없음"
    ws.Columns("A:F").AutoFit
End Sub

Private Sub SampleDesign_GroupAllocationSummary(ByVal allocation As Collection, ByVal groupId As String, ByRef finalN As Long, ByRef rowCount As Long)
    Dim i As Long
    For i = 1 To allocation.Count
        Dim outRec As Object, rec As Object
        Set outRec = allocation(i)
        Set rec = outRec("source")
        If CStr(rec("targetGroupId")) = groupId Then
            finalN = finalN + CLng(outRec("final"))
            rowCount = rowCount + 1
        End If
    Next i
End Sub

Private Function SampleDesign_TotalFinalN(ByVal allocation As Collection) As Long
    Dim i As Long
    For i = 1 To allocation.Count
        Dim outRec As Object
        Set outRec = allocation(i)
        SampleDesign_TotalFinalN = SampleDesign_TotalFinalN + CLng(outRec("final"))
    Next i
End Function

Private Sub SampleDesign_ApplyProgressFormatting(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal statusCol As Long)
    If lastRow < firstRow Then Exit Sub
    On Error Resume Next
    With ws.Range(ws.Cells(firstRow, statusCol), ws.Cells(lastRow, statusCol)).FormatConditions.Add(Type:=xlTextString, String:="미달", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 245, 190)
    End With
    With ws.Range(ws.Cells(firstRow, statusCol), ws.Cells(lastRow, statusCol)).FormatConditions.Add(Type:=xlTextString, String:="충족", TextOperator:=xlContains)
        .Interior.Color = RGB(220, 245, 220)
    End With
    With ws.Range(ws.Cells(firstRow, statusCol), ws.Cells(lastRow, statusCol)).FormatConditions.Add(Type:=xlTextString, String:="초과", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 220, 220)
    End With
    On Error GoTo 0
End Sub

Private Sub SampleDesign_WriteInputTemplate_Legacy(ByVal ws As Worksheet)
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

Private Sub SampleDesign_WriteEmptyResult_Legacy(ByVal ws As Worksheet)
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

Public Sub SampleDesign_ShowGuide()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "표본설계 가이드를 만들 통합문서를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = SampleDesign_FindSheet(wb, "표본설계_가이드")
    If ws Is Nothing Then
        Set ws = SampleDesign_PrepareOutputSheet(wb, "표본설계_가이드")
        SampleDesign_WriteGuideSheet ws, SampleDesign_DefaultSetup(), "list"
    End If
    ws.Activate
    ws.Range("A1").Select
End Sub

Private Sub SampleDesign_WriteGuideSheet(ByVal ws As Worksheet, ByVal setup As Object, ByVal templateMode As String)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 사용 가이드"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3").Value = "권장 작업 순서"
    ws.Range("A3").Font.Bold = True
    SampleDesign_WriteArray ws, 4, 1, Array("1", "설계 시트 생성", "설정 마법사에서 입력 형태, 목표 표본수, 사용할 쿼터 변수를 먼저 정합니다.")
    SampleDesign_WriteArray ws, 5, 1, Array("2", "모집단 입력", "표본설계_입력 시트의 21행부터 쿼터값과 모집단수를 입력합니다.")
    SampleDesign_WriteArray ws, 6, 1, Array("3", "그룹 목표 입력", "상위 쿼터별 목표 표본수가 다르면 목표표본 그룹표에 조건과 목표표본수를 입력합니다.")
    SampleDesign_WriteArray ws, 7, 1, Array("4", "구간 변수 생성", "매출액, 종사자수, 인구수 같은 실제 값은 구간 만들기 버튼으로 오른쪽에 구간 열을 만든 뒤 쿼터로 사용할 수 있습니다.")
    SampleDesign_WriteArray ws, 8, 1, Array("5", "할당 계산", "할당 계산 버튼을 눌러 배분 방식을 확인하고 결과/검토 시트를 생성합니다.")
    SampleDesign_WriteArray ws, 9, 1, Array("6", "진행관리", "조사 진행 중에는 진행률 관리표 버튼으로 회수/부족/초과를 관리합니다.")

    ws.Range("A11").Value = "필수 입력값"
    ws.Range("A11").Font.Bold = True
    SampleDesign_WriteArray ws, 12, 1, Array("위치", "필수 여부", "설명")
    SampleDesign_WriteArray ws, 13, 1, Array("표본설계_입력!B4", "필수", "전체 목표 표본수. 그룹표가 비어 있으면 이 값으로 전체 배분합니다.")
    SampleDesign_WriteArray ws, 14, 1, Array("표본설계_입력!B5", "필수", "배분 방식. 단순 비례, 제곱근 비례, 멱배분, 균등 배분, 사용자 지정 가중 중 선택합니다.")
    SampleDesign_WriteArray ws, 15, 1, Array("표본설계_입력!19행", "필수", "사용할 쿼터 변수는 '사용', 제외할 변수는 '미사용'으로 둡니다.")
    SampleDesign_WriteArray ws, 16, 1, Array("표본설계_입력!20행", "필수", "쿼터 변수명입니다. 산업, 지역, 성별 등 프로젝트 용어로 바꿔도 됩니다.")
    SampleDesign_WriteArray ws, 17, 1, Array("표본설계_입력!21행 이하", "필수", "쿼터ID, 쿼터값, 모집단수를 입력합니다. 모집단수 1 이상만 계산 대상입니다.")

    ws.Range("A20").Value = "목표표본 그룹표 사용법"
    ws.Range("A20").Font.Bold = True
    SampleDesign_WriteArray ws, 21, 1, Array("상황", "입력 방법")
    SampleDesign_WriteArray ws, 22, 1, Array("전체 목표만 있음", "목표표본 그룹표를 비우고 B4만 입력합니다.")
    SampleDesign_WriteArray ws, 23, 1, Array("산업별 목표가 다름", "목표표본 그룹표에 산업 조건과 목표표본수를 입력합니다.")
    SampleDesign_WriteArray ws, 24, 1, Array("산업+지역 조합별 목표가 다름", "같은 행에 산업과 지역 조건을 모두 입력합니다. 조건은 AND로 매칭됩니다.")
    SampleDesign_WriteArray ws, 25, 1, Array("나머지 전체 목표가 필요", "조건 칸을 비운 행을 만들면 catch-all 그룹으로 사용할 수 있습니다.")

    ws.Range("A28").Value = "현재 생성 설정"
    ws.Range("A28").Font.Bold = True
    SampleDesign_WriteArray ws, 29, 1, Array("입력 형태", IIf(templateMode = "table", "집계표형 입력", "리스트형 입력"))
    SampleDesign_WriteArray ws, 30, 1, Array("전체 목표 표본수", setup("targetN"))
    SampleDesign_WriteArray ws, 31, 1, Array("배분 방식", setup("methodName"))
    SampleDesign_WriteArray ws, 32, 1, Array("멱배분 alpha", setup("alpha"))
    SampleDesign_WriteArray ws, 33, 1, Array("기본 최소/최대표본", CStr(setup("minN")) & " / " & CStr(setup("maxN")))
    SampleDesign_WriteArray ws, 34, 1, Array("사용 쿼터", SampleDesign_EnabledQuotaSummary(setup("quotaHeaders"), setup("quotaEnabled")))

    ws.Range("A37").Value = "자주 나는 오류와 조치"
    ws.Range("A37").Font.Bold = True
    SampleDesign_WriteArray ws, 38, 1, Array("메시지", "조치")
    SampleDesign_WriteArray ws, 39, 1, Array("모집단수 헤더를 찾지 못함", "표본설계_입력 시트에서 '모집단수' 헤더가 삭제되었는지 확인합니다.")
    SampleDesign_WriteArray ws, 40, 1, Array("목표 표본 수를 찾지 못함", "B4 또는 목표표본 그룹표의 목표표본수를 입력합니다.")
    SampleDesign_WriteArray ws, 41, 1, Array("계산할 모집단 행을 찾지 못함", "21행 이하 모집단수에 1 이상의 숫자가 있는지 확인합니다.")
    SampleDesign_WriteArray ws, 42, 1, Array("목표그룹없음", "목표표본 그룹표의 조건이 모집단 표의 쿼터값과 정확히 일치하는지 확인합니다.")
    SampleDesign_WriteArray ws, 43, 1, Array("합계불일치", "최소/최대표본 제약 때문에 목표합계와 최종할당 합계가 다를 수 있습니다. 검토 시트의 목표그룹 검토표에서 차이를 확인합니다.")

    ws.Range("A46").Value = "계산 방식 검토"
    ws.Range("A46").Font.Bold = True
    SampleDesign_WriteArray ws, 47, 1, Array("항목", "설명")
    SampleDesign_WriteArray ws, 48, 1, Array("배분가중치", "단순 비례=모집단수, 제곱근 비례=SQRT(모집단수), 멱배분=모집단수^alpha, 균등=1, 사용자 지정=사용자가중치입니다.")
    SampleDesign_WriteArray ws, 49, 1, Array("이론할당", "그룹목표표본수 x 배분가중치 / 그룹가중치합으로 계산합니다.")
    SampleDesign_WriteArray ws, 50, 1, Array("최종할당", "이론할당을 정수화한 뒤 잔차가 큰 셀부터 보정해 목표그룹 합계를 맞춥니다.")
    SampleDesign_WriteArray ws, 51, 1, Array("제약 검토", "최소표본/최대표본을 강하게 넣으면 목표합계를 맞출 수 없는 경우가 있으므로 검토 시트의 경고를 확인합니다.")

    With ws.Range("A4:C9")
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A12:C17")
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A21:B25")
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A29:B34")
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A38:B43")
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A47:B51")
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("A3:C3,A11:C12,A20:B21,A28:B28,A37:B38,A46:B47").Interior.Color = RGB(220, 235, 250)
    ws.Columns("A:C").AutoFit
End Sub

Private Function SampleDesign_EnabledQuotaSummary(ByVal headers As Variant, ByVal enabled As Variant) As String
    Dim result As String
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If CBool(enabled(i)) Then
            If Len(result) > 0 Then result = result & ", "
            result = result & CStr(headers(i))
        End If
    Next i
    If Len(result) = 0 Then result = "(선택 없음)"
    SampleDesign_EnabledQuotaSummary = result
End Function

Private Sub SampleDesign_ReadRunSettings(ByVal inputWs As Worksheet, ByRef methodName As String, ByRef alpha As Double, _
                                         ByRef minN As Long, ByRef maxN As Long)
    If IsNumeric(inputWs.Range("B5").Value) Or Len(Trim$(CStr(inputWs.Range("B5").Value))) = 0 Then
        ' 구버전 레이아웃: B5 최소표본, B6 최대표본, B7 배분방식, B8 alpha
        minN = CLng(Val(inputWs.Range("B5").Value))
        maxN = CLng(Val(inputWs.Range("B6").Value))
        methodName = Trim$(CStr(inputWs.Range("B7").Value))
        alpha = Val(inputWs.Range("B8").Value)
    Else
        ' 1.1.16+ 레이아웃: B5 배분방식, B6 alpha, B7 최소표본, B8 최대표본
        methodName = Trim$(CStr(inputWs.Range("B5").Value))
        alpha = Val(inputWs.Range("B6").Value)
        minN = CLng(Val(inputWs.Range("B7").Value))
        maxN = CLng(Val(inputWs.Range("B8").Value))
    End If

    If Len(methodName) = 0 Then methodName = "단순 비례"
    If alpha <= 0 And methodName = "멱배분" Then alpha = 0.5
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

Private Sub SampleDesign_ReadPopulation_Legacy(ByVal ws As Worksheet, ByRef records As Collection, ByRef issues As Collection)
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

Private Function SampleDesign_Allocate_Legacy(ByVal records As Collection, ByVal targetN As Long, _
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
        Set SampleDesign_Allocate_Legacy = result
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
    Set SampleDesign_Allocate_Legacy = result
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

Private Sub SampleDesign_WriteAllocationResult_Legacy(ByVal ws As Worksheet, ByVal allocation As Collection, ByVal targetN As Long, ByVal methodName As String)
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

Private Sub SampleDesign_WriteReviewReport_Legacy(ByVal ws As Worksheet, ByVal allocation As Collection, ByVal issues As Collection, _
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

Private Function SampleDesign_TotalFinalN_Legacy(ByVal allocation As Collection) As Long
    Dim i As Long
    For i = 1 To allocation.Count
        Dim rec As Variant
        rec = allocation(i)
        SampleDesign_TotalFinalN_Legacy = SampleDesign_TotalFinalN_Legacy + CLng(rec(10))
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

Private Sub SampleDesign_ApplyProgressFormatting_Legacy(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
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
    Dim headerRowSetting As String
    Dim headerMode As String
    Dim idMode As String
    Dim keyMode As String
    Dim useMarkerRow As Boolean

    keyHeader = "ID"
    outputHeadersCsv = DBMerge_DefaultHeaderCsv(wb)
    outputSheetName = "통합데이터"
    outputPosition = "맨 뒤"
    headerRowSetting = "자동"
    headerMode = "single"
    idMode = "output_auto"
    keyMode = "header"
    useMarkerRow = False

    If Not DBMerge_ShowSetupForm(wb, selectedSheetNames, keyHeader, outputHeadersCsv, outputSheetName, outputPosition, headerRowSetting, headerMode, idMode, keyMode, useMarkerRow) Then Exit Sub

    keyHeader = Trim$(keyHeader)
    headerRowSetting = Trim$(headerRowSetting)
    keyMode = DBMerge_NormalizeKeyMode(keyMode)
    outputSheetName = DBMerge_CleanSheetName(outputSheetName, "통합데이터")
    If Len(keyHeader) = 0 And idMode <> "existing" Then keyHeader = "ID"
    If Len(keyHeader) = 0 Then
        MsgBox "메인 KEY 헤더를 입력해야 합니다.", vbExclamation, "시트 데이터 통합"
        Exit Sub
    End If
    If Len(headerRowSetting) = 0 Then headerRowSetting = "자동"

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

    Dim manualKeyColumns As Object
    If keyMode = "manual" Then
        Set manualKeyColumns = DBMerge_RequestManualKeyColumns(sheets, keyHeader)
        If manualKeyColumns Is Nothing Then Exit Sub
    End If

    Dim outputHeaders As Collection
    Set outputHeaders = DBMerge_ParseHeaderList(outputHeadersCsv, keyHeader)
    If outputHeaders.Count = 0 Or (outputHeaders.Count = 1 And Not useMarkerRow) Then
        MsgBox "통합할 헤더를 하나 이상 입력하세요.", vbExclamation, "시트 데이터 통합"
        Exit Sub
    End If

    Dim plan As Object
    Set plan = DBMerge_BuildPlan(sheets, keyHeader, outputHeaders, headerRowSetting, headerMode, idMode, keyMode, useMarkerRow, manualKeyColumns)
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

Public Sub Database_CreateMergeMarkerRows()
    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim selectedKeys As String
    selectedKeys = Toolkit_CheckListMenu( _
        "통합 표시행 만들기", _
        "표시행을 추가할 시트를 선택하세요. 헤더 바로 위에 빈 행을 만들고, 통합할 열 위에 x, 1 등 아무 값을 입력하면 됩니다.", _
        DBMerge_WorksheetOptionsText(wb), _
        DBMerge_AllWorksheetDefaultCsv(wb))
    If Len(Trim$(selectedKeys)) = 0 Then Exit Sub

    Dim headerMode As String
    headerMode = Toolkit_ButtonMenu( _
        "통합 표시행 - 헤더 구조", _
        "표시행을 넣을 시트의 헤더 구조를 선택하세요.", _
        "single" & vbTab & "1줄 헤더" & vbLf & _
        "display_next_key" & vbTab & "위=표시명, 아래=변수" & vbLf & _
        "key_prev_display" & vbTab & "이전=표시명, 선택=변수", _
        "single")
    If Len(headerMode) = 0 Then Exit Sub

    Dim headerRowSetting As String
    headerRowSetting = InputBox("헤더 행 번호를 입력하세요." & vbCrLf & _
                                "자동 탐지를 쓰려면 '자동' 또는 '0'을 입력하세요.", _
                                "통합 표시행 만들기", "자동")
    If Len(headerRowSetting) = 0 Then Exit Sub

    Dim sheets As Collection
    Set sheets = DBMerge_SelectedWorksheetsByKeys(wb, selectedKeys)
    If sheets.Count = 0 Then
        MsgBox "표시행을 추가할 시트를 선택하지 않았습니다.", vbExclamation, "통합 표시행 만들기"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim insertedCount As Long
    Dim skippedCount As Long
    Dim i As Long
    For i = 1 To sheets.Count
        Dim ws As Worksheet
        Set ws = sheets(i)

        Dim headerRow As Long
        headerRow = DBMerge_HeaderRowForSheet(ws, headerRowSetting, headerMode, False)
        If headerRow = 0 Then
            skippedCount = skippedCount + 1
        Else
            Dim insertRow As Long
            insertRow = DBMerge_TopHeaderRow(headerRow, headerMode)
            If insertRow < 1 Then
                skippedCount = skippedCount + 1
            Else
                ws.Rows(insertRow).Insert Shift:=xlDown
                With ws.Rows(insertRow)
                    .Interior.Color = RGB(255, 242, 204)
                    .Font.Color = RGB(120, 90, 0)
                    .RowHeight = 20
                End With
                insertedCount = insertedCount + 1
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox insertedCount & "개 시트에 통합 표시행을 추가했습니다." & vbCrLf & _
           skippedCount & "개 시트는 헤더 행을 찾지 못해 건너뛰었습니다.", vbInformation, "통합 표시행 만들기"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "통합 표시행을 만드는 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "통합 표시행 만들기"
End Sub

Private Function DBMerge_ShowSetupForm(ByVal wb As Workbook, ByRef selectedSheetNames As String, _
                                       ByRef keyHeader As String, ByRef outputHeadersCsv As String, _
                                       ByRef outputSheetName As String, ByRef outputPosition As String, _
                                       ByRef headerRowSetting As String, ByRef headerMode As String, _
                                       ByRef idMode As String, ByRef keyMode As String, _
                                       ByRef useMarkerRow As Boolean) As Boolean
    On Error GoTo ErrHandler

    Dim form As Object
    Set form = VBA.UserForms.Add("frmDatabaseMergeSetup")
    form.Configure DBMerge_WorksheetNamesText(wb), keyHeader, outputHeadersCsv, outputSheetName, outputPosition, headerRowSetting, headerMode, idMode, keyMode, useMarkerRow
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
    headerRowSetting = CStr(form.HeaderRowSetting)
    headerMode = CStr(form.HeaderMode)
    idMode = CStr(form.IdMode)
    keyMode = CStr(form.KeyMode)
    useMarkerRow = CBool(form.UseMarkerRow)
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

Private Function DBMerge_WorksheetOptionsText(ByVal wb As Workbook) As String
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If Len(DBMerge_WorksheetOptionsText) > 0 Then DBMerge_WorksheetOptionsText = DBMerge_WorksheetOptionsText & vbLf
        DBMerge_WorksheetOptionsText = DBMerge_WorksheetOptionsText & CStr(i) & vbTab & wb.Worksheets(i).Name
    Next i
End Function

Private Function DBMerge_AllWorksheetDefaultCsv(ByVal wb As Workbook) As String
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If Len(DBMerge_AllWorksheetDefaultCsv) > 0 Then DBMerge_AllWorksheetDefaultCsv = DBMerge_AllWorksheetDefaultCsv & ","
        DBMerge_AllWorksheetDefaultCsv = DBMerge_AllWorksheetDefaultCsv & CStr(i)
    Next i
End Function

Private Function DBMerge_SelectedWorksheetsByKeys(ByVal wb As Workbook, ByVal selectedKeys As String) As Collection
    Dim sheets As Collection
    Set sheets = New Collection

    Dim parts As Variant
    parts = Split(selectedKeys, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim keyText As String
        keyText = Trim$(CStr(parts(i)))
        If IsNumeric(keyText) Then
            Dim idx As Long
            idx = CLng(keyText)
            If idx >= 1 And idx <= wb.Worksheets.Count Then sheets.Add wb.Worksheets(idx)
        End If
    Next i

    Set DBMerge_SelectedWorksheetsByKeys = sheets
End Function

Private Function DBMerge_DefaultHeaderCsv(ByVal wb As Workbook) As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim headerRow As Long
        headerRow = DBMerge_HeaderRowForSheet(ws, "자동", "single", True)
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

Private Function DBMerge_NormalizeKeyMode(ByVal keyMode As String) As String
    If LCase$(Trim$(keyMode)) = "manual" Then
        DBMerge_NormalizeKeyMode = "manual"
    Else
        DBMerge_NormalizeKeyMode = "header"
    End If
End Function

Private Function DBMerge_RequestManualKeyColumns(ByVal sheets As Collection, ByVal keyHeader As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To sheets.Count
        Dim ws As Worksheet
        Set ws = sheets(i)

        ws.Activate

        Dim keyRange As Range
        On Error Resume Next
        Set keyRange = Application.InputBox( _
            ws.Name & " 시트에서 '" & keyHeader & "'로 사용할 KEY 열을 선택하세요." & vbCrLf & _
            "헤더 셀 하나, 열 일부, 또는 전체 열을 선택할 수 있습니다.", _
            "시트 데이터 통합 - KEY 열 직접 지정", Type:=8)
        On Error GoTo 0

        If keyRange Is Nothing Then
            Set DBMerge_RequestManualKeyColumns = Nothing
            Exit Function
        End If
        Set keyRange = keyRange.Areas(1)
        If Not keyRange.Worksheet Is ws Then
            MsgBox "KEY 열은 현재 지정 중인 '" & ws.Name & "' 시트에서 선택해야 합니다.", vbExclamation, "시트 데이터 통합"
            Set DBMerge_RequestManualKeyColumns = Nothing
            Exit Function
        End If
        If keyRange.Columns.Count <> 1 Then
            MsgBox "KEY 열은 한 개 열만 선택하세요.", vbExclamation, "시트 데이터 통합"
            Set DBMerge_RequestManualKeyColumns = Nothing
            Exit Function
        End If

        result(ws.Name) = CLng(keyRange.Column)
    Next i

    Set DBMerge_RequestManualKeyColumns = result
End Function

Private Function DBMerge_ApplyManualKeyColumnForSheet(ByVal headerMap As Object, ByVal keyHeader As String, _
                                                      ByVal manualKeyColumns As Object, ByVal sheetName As String) As Boolean
    If manualKeyColumns Is Nothing Then Exit Function
    If Not manualKeyColumns.Exists(sheetName) Then Exit Function

    DBMerge_ApplyManualKeyColumnToMap headerMap, keyHeader, CLng(manualKeyColumns(sheetName))
    DBMerge_ApplyManualKeyColumnForSheet = True
End Function

Private Sub DBMerge_ApplyManualKeyColumnToMap(ByVal headerMap As Object, ByVal keyHeader As String, ByVal keyCol As Long)
    If keyCol < 1 Then Exit Sub

    Dim keyNorm As String
    keyNorm = DBMerge_NormalizeHeader(keyHeader)
    If Len(keyNorm) = 0 Then Exit Sub

    If headerMap.Exists(keyNorm) Then
        headerMap(keyNorm) = keyCol
    Else
        headerMap.Add keyNorm, keyCol
    End If
End Sub

Public Function DBMerge_ManualKeyColumnForTest(ByVal keyHeader As String, ByVal existingCol As Long, ByVal manualCol As Long) As Long
    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")

    Dim keyNorm As String
    keyNorm = DBMerge_NormalizeHeader(keyHeader)
    If Len(keyNorm) > 0 And existingCol > 0 Then headerMap.Add keyNorm, existingCol

    DBMerge_ApplyManualKeyColumnToMap headerMap, keyHeader, manualCol
    If headerMap.Exists(keyNorm) Then DBMerge_ManualKeyColumnForTest = CLng(headerMap(keyNorm))
End Function

Private Function DBMerge_BuildPlan(ByVal sheets As Collection, ByVal keyHeader As String, ByVal outputHeaders As Collection, _
                                   ByVal headerRowSetting As String, ByVal headerMode As String, _
                                   ByVal idMode As String, ByVal keyMode As String, _
                                   ByVal useMarkerRow As Boolean, ByVal manualKeyColumns As Object) As Object
    Dim plan As Object
    Set plan = CreateObject("Scripting.Dictionary")
    plan.Add "keyHeader", keyHeader
    plan.Add "headerRowSetting", headerRowSetting
    plan.Add "headerMode", headerMode
    plan.Add "idMode", idMode
    plan.Add "keyMode", DBMerge_NormalizeKeyMode(keyMode)
    plan.Add "useMarkerRow", useMarkerRow

    Dim sheetPlans As Collection
    Set sheetPlans = New Collection
    Dim markerHeaders As Collection
    Set markerHeaders = New Collection
    Dim markerSeen As Object
    Set markerSeen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To sheets.Count
        Dim ws As Worksheet
        Set ws = sheets(i)

        Dim headerRow As Long
        headerRow = DBMerge_HeaderRowForSheet(ws, headerRowSetting, headerMode, useMarkerRow)
        If headerRow = 0 Then
            MsgBox ws.Name & " 시트에서 헤더 행을 찾지 못했습니다.", vbExclamation, "시트 데이터 통합"
            Exit Function
        End If
        If Not DBMerge_HeaderModeIsUsable(ws, headerRow, headerMode) Then
            MsgBox ws.Name & " 시트에서 선택한 헤더 구조를 적용할 수 없습니다." & vbCrLf & _
                   "헤더 행 번호와 헤더 구조를 다시 확인하세요.", vbExclamation, "시트 데이터 통합"
            Exit Function
        End If

        Dim dataStartRow As Long
        dataStartRow = DBMerge_DataStartRow(headerRow, headerMode)

        Dim headerMap As Object
        Set headerMap = DBMerge_HeaderMap(ws, headerRow, headerMode)
        If CStr(plan("keyMode")) = "manual" Then
            If Not DBMerge_ApplyManualKeyColumnForSheet(headerMap, keyHeader, manualKeyColumns, ws.Name) Then
                MsgBox ws.Name & " 시트의 수동 KEY 열 지정 정보를 찾지 못했습니다.", vbExclamation, "시트 데이터 통합"
                Exit Function
            End If
        End If
        If idMode = "existing" And Not headerMap.Exists(DBMerge_NormalizeHeader(keyHeader)) Then
            MsgBox ws.Name & " 시트에서 메인 KEY 헤더 '" & keyHeader & "'를 찾지 못했습니다." & vbCrLf & _
                   "ID가 없는 시트는 ID 처리 방식을 자동 생성으로 바꾸세요.", vbExclamation, "시트 데이터 통합"
            Exit Function
        End If

        Dim markerRow As Long
        markerRow = 0
        If useMarkerRow Then
            markerRow = DBMerge_MarkerRowForHeader(headerRow, headerMode)
            If markerRow < 1 Then
                MsgBox ws.Name & " 시트의 헤더 위에 통합 표시행이 없습니다." & vbCrLf & _
                       "먼저 [통합 표시행 만들기]를 실행하거나 헤더 행 번호를 다시 지정하세요.", vbExclamation, "시트 데이터 통합"
                Exit Function
            End If
            DBMerge_AddHeadersFromMarkerRow markerHeaders, markerSeen, ws, markerRow, headerRow, headerMode
        End If

        Dim sheetPlan As Object
        Set sheetPlan = CreateObject("Scripting.Dictionary")
        sheetPlan.Add "sheetName", ws.Name
        sheetPlan.Add "worksheet", ws
        sheetPlan.Add "headerRow", headerRow
        sheetPlan.Add "markerRow", markerRow
        sheetPlan.Add "dataStartRow", dataStartRow
        sheetPlan.Add "headerMap", headerMap
        sheetPlans.Add sheetPlan
    Next i

    Dim finalHeaders As Collection
    If useMarkerRow Then
        Set finalHeaders = DBMerge_FinalHeadersFromMarkers(keyHeader, markerHeaders)
        If finalHeaders.Count <= 1 Then
            MsgBox "통합 표시행에서 선택된 열을 찾지 못했습니다." & vbCrLf & _
                   "헤더 위 표시행에 통합할 열만 아무 값으로 표시하세요.", vbExclamation, "시트 데이터 통합"
            Exit Function
        End If
    Else
        Set finalHeaders = outputHeaders
    End If

    For i = 1 To sheetPlans.Count
        Set sheetPlan = sheetPlans(i)
        Set ws = sheetPlan("worksheet")
        Set headerMap = sheetPlan("headerMap")
        sheetPlan.Add "lastRow", DBMerge_LastDataRow(ws, CLng(sheetPlan("dataStartRow")), headerMap, finalHeaders)
        DBMerge_UpdateIdStats sheetPlan, keyHeader, idMode, finalHeaders
    Next i

    plan.Add "headers", finalHeaders
    plan.Add "sheets", sheetPlans
    Set DBMerge_BuildPlan = plan
End Function

Private Function DBMerge_HeaderRowForSheet(ByVal ws As Worksheet, ByVal headerRowSetting As String, _
                                           Optional ByVal headerMode As String = "single", _
                                           Optional ByVal useMarkerRow As Boolean = False) As Long
    Dim settingText As String
    settingText = LCase$(Trim$(headerRowSetting))
    If Len(settingText) = 0 Or settingText = "자동" Or settingText = "auto" Or settingText = "0" Then
        DBMerge_HeaderRowForSheet = DBMerge_DetectHeaderRow(ws)
        If useMarkerRow And DBMerge_RowLooksLikeMarker(ws, DBMerge_HeaderRowForSheet) Then
            DBMerge_HeaderRowForSheet = DBMerge_HeaderRowForSheet + 1
        End If
        Exit Function
    End If

    If Not IsNumeric(settingText) Then Exit Function

    Dim rowNo As Long
    rowNo = CLng(settingText)
    If rowNo < 1 Or rowNo > DBMerge_LastUsedRow(ws) Then Exit Function
    DBMerge_HeaderRowForSheet = rowNo
End Function

Private Function DBMerge_TopHeaderRow(ByVal headerRow As Long, ByVal headerMode As String) As Long
    If headerMode = "key_prev_display" Then
        DBMerge_TopHeaderRow = headerRow - 1
    Else
        DBMerge_TopHeaderRow = headerRow
    End If
End Function

Private Function DBMerge_MarkerRowForHeader(ByVal headerRow As Long, ByVal headerMode As String) As Long
    DBMerge_MarkerRowForHeader = DBMerge_TopHeaderRow(headerRow, headerMode) - 1
End Function

Private Function DBMerge_RowLooksLikeMarker(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    If rowNo < 1 Or rowNo >= DBMerge_LastUsedRow(ws) Then Exit Function

    Dim lastCol As Long
    lastCol = DBMerge_LastUsedCol(ws)
    If lastCol = 0 Then Exit Function

    If ws.Cells(rowNo, 1).Interior.Color = RGB(255, 242, 204) And DBMerge_RowTextCount(ws, rowNo + 1, lastCol) >= 2 Then
        DBMerge_RowLooksLikeMarker = True
        Exit Function
    End If

    Dim nonBlankCount As Long
    Dim shortMarkCount As Long
    Dim c As Long
    For c = 1 To lastCol
        Dim cellText As String
        cellText = Trim$(CStr(ws.Cells(rowNo, c).Value))
        If Len(cellText) > 0 Then
            nonBlankCount = nonBlankCount + 1
            If Len(cellText) <= 3 Then shortMarkCount = shortMarkCount + 1
        End If
    Next c

    If nonBlankCount = 0 Then Exit Function
    DBMerge_RowLooksLikeMarker = (nonBlankCount = shortMarkCount And DBMerge_RowTextCount(ws, rowNo + 1, lastCol) >= 2)
End Function

Private Sub DBMerge_AddHeadersFromMarkerRow(ByVal headers As Collection, ByVal seen As Object, _
                                            ByVal ws As Worksheet, ByVal markerRow As Long, _
                                            ByVal headerRow As Long, ByVal headerMode As String)
    Dim lastCol As Long
    lastCol = DBMerge_LastUsedCol(ws)

    Dim c As Long
    For c = 1 To lastCol
        If Len(Trim$(CStr(ws.Cells(markerRow, c).Value))) > 0 Then
            Dim headerText As String
            headerText = DBMerge_HeaderTextForColumn(ws, c, headerRow, headerMode)
            DBMerge_AddHeader headers, seen, headerText
        End If
    Next c
End Sub

Private Function DBMerge_HeaderTextForColumn(ByVal ws As Worksheet, ByVal colNo As Long, _
                                             ByVal headerRow As Long, ByVal headerMode As String) As String
    Select Case headerMode
        Case "display_next_key"
            DBMerge_HeaderTextForColumn = Trim$(CStr(ws.Cells(headerRow + 1, colNo).Value))
            If Len(DBMerge_HeaderTextForColumn) = 0 Then DBMerge_HeaderTextForColumn = Trim$(CStr(ws.Cells(headerRow, colNo).Value))
        Case "key_prev_display"
            DBMerge_HeaderTextForColumn = Trim$(CStr(ws.Cells(headerRow, colNo).Value))
            If Len(DBMerge_HeaderTextForColumn) = 0 Then DBMerge_HeaderTextForColumn = Trim$(CStr(ws.Cells(headerRow - 1, colNo).Value))
        Case Else
            DBMerge_HeaderTextForColumn = Trim$(CStr(ws.Cells(headerRow, colNo).Value))
    End Select
End Function

Private Function DBMerge_FinalHeadersFromMarkers(ByVal keyHeader As String, ByVal markerHeaders As Collection) As Collection
    Dim headers As Collection
    Set headers = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    DBMerge_AddHeader headers, seen, keyHeader

    Dim i As Long
    For i = 1 To markerHeaders.Count
        DBMerge_AddHeader headers, seen, CStr(markerHeaders(i))
    Next i

    Set DBMerge_FinalHeadersFromMarkers = headers
End Function

Private Function DBMerge_HeaderModeIsUsable(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerMode As String) As Boolean
    Dim lastRow As Long
    lastRow = DBMerge_LastUsedRow(ws)

    Select Case headerMode
        Case "display_next_key"
            DBMerge_HeaderModeIsUsable = (headerRow + 1 <= lastRow)
        Case "key_prev_display"
            DBMerge_HeaderModeIsUsable = (headerRow > 1)
        Case Else
            DBMerge_HeaderModeIsUsable = (headerRow <= lastRow)
    End Select
End Function

Private Function DBMerge_DataStartRow(ByVal headerRow As Long, ByVal headerMode As String) As Long
    Select Case headerMode
        Case "display_next_key"
            DBMerge_DataStartRow = headerRow + 2
        Case Else
            DBMerge_DataStartRow = headerRow + 1
    End Select
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

Private Function DBMerge_HeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerMode As String) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Select Case headerMode
        Case "display_next_key"
            DBMerge_AddHeaderMapRow ws, map, headerRow + 1
            DBMerge_AddHeaderMapRow ws, map, headerRow
        Case "key_prev_display"
            DBMerge_AddHeaderMapRow ws, map, headerRow
            DBMerge_AddHeaderMapRow ws, map, headerRow - 1
        Case Else
            DBMerge_AddHeaderMapRow ws, map, headerRow
    End Select

    Set DBMerge_HeaderMap = map
End Function

Private Sub DBMerge_AddHeaderMapRow(ByVal ws As Worksheet, ByVal map As Object, ByVal headerRow As Long)
    If headerRow < 1 Then Exit Sub

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
End Sub

Private Function DBMerge_LastDataRow(ByVal ws As Worksheet, ByVal dataStartRow As Long, ByVal headerMap As Object, ByVal outputHeaders As Collection) As Long
    Dim lastRow As Long
    lastRow = dataStartRow - 1

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

    If lastRow < dataStartRow Then lastRow = DBMerge_LastUsedRow(ws)
    DBMerge_LastDataRow = lastRow
End Function

Private Function DBMerge_PlanSummary(ByVal plan As Object) As String
    Dim result As String
    result = "메인 KEY: " & CStr(plan("keyHeader")) & vbCrLf & _
             "KEY 지정: " & DBMerge_KeyModeLabel(CStr(plan("keyMode"))) & vbCrLf & _
             "ID 처리: " & DBMerge_IdModeLabel(CStr(plan("idMode"))) & vbCrLf & _
             "열 선택: " & IIf(CBool(plan("useMarkerRow")), "통합 표시행 기준", "헤더 목록 직접 지정") & vbCrLf & _
             "헤더 구조: " & DBMerge_HeaderModeLabel(CStr(plan("headerMode"))) & vbCrLf & _
             "헤더 행 설정: " & CStr(plan("headerRowSetting")) & vbCrLf & _
             "통합 헤더 수: " & CStr(plan("headers").Count) & vbCrLf & _
             "대상 시트 수: " & CStr(plan("sheets").Count) & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To plan("sheets").Count
        Dim sheetPlan As Object
        Set sheetPlan = plan("sheets")(i)
        result = result & i & ". " & CStr(sheetPlan("sheetName")) & _
                 " / 헤더 행: " & CStr(sheetPlan("headerRow")) & _
                 " / 데이터 시작: " & CStr(sheetPlan("dataStartRow")) & _
                 " / 마지막 행: " & CStr(sheetPlan("lastRow")) & _
                 DBMerge_KeyColumnSummary(CStr(plan("keyMode")), CStr(plan("keyHeader")), sheetPlan("headerMap")) & _
                 " / 기존ID: " & CStr(sheetPlan("idExistingCount")) & _
                 " / 자동ID: " & CStr(sheetPlan("idGeneratedCount")) & vbCrLf
    Next i
    If CStr(plan("idMode")) = "source_auto" Then
        result = result & vbCrLf & "주의: 실행 시 원본 시트에 ID 열 또는 빈 ID 값이 생성됩니다." & vbCrLf
    End If
    DBMerge_PlanSummary = result
End Function

Private Function DBMerge_KeyModeLabel(ByVal keyMode As String) As String
    If DBMerge_NormalizeKeyMode(keyMode) = "manual" Then
        DBMerge_KeyModeLabel = "시트별 KEY 열 직접 지정"
    Else
        DBMerge_KeyModeLabel = "메인 KEY 헤더명으로 자동 매칭"
    End If
End Function

Private Function DBMerge_KeyColumnSummary(ByVal keyMode As String, ByVal keyHeader As String, ByVal headerMap As Object) As String
    If DBMerge_NormalizeKeyMode(keyMode) <> "manual" Then Exit Function

    Dim keyNorm As String
    keyNorm = DBMerge_NormalizeHeader(keyHeader)
    If headerMap.Exists(keyNorm) Then
        DBMerge_KeyColumnSummary = " / KEY 열: " & DBMerge_ColumnLetter(CLng(headerMap(keyNorm)))
    Else
        DBMerge_KeyColumnSummary = " / KEY 열: 없음"
    End If
End Function

Private Function DBMerge_IdModeLabel(ByVal idMode As String) As String
    Select Case idMode
        Case "existing"
            DBMerge_IdModeLabel = "기존 KEY 헤더 사용"
        Case "source_auto"
            DBMerge_IdModeLabel = "ID 없으면 원본 시트에 ID열 생성"
        Case Else
            DBMerge_IdModeLabel = "ID 없으면 결과 시트에만 자동 생성"
    End Select
End Function

Private Function DBMerge_HeaderModeLabel(ByVal headerMode As String) As String
    Select Case headerMode
        Case "display_next_key"
            DBMerge_HeaderModeLabel = "선택 행=헤더변수명, 다음 행=실제 헤더변수"
        Case "key_prev_display"
            DBMerge_HeaderModeLabel = "선택 행=실제 헤더변수, 이전 행=헤더변수명"
        Case Else
            DBMerge_HeaderModeLabel = "선택 행 1줄"
    End Select
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
    If CStr(plan("idMode")) = "source_auto" Then DBMerge_ApplySourceIds plan

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
        For r = CLng(sheetPlan("dataStartRow")) To CLng(sheetPlan("lastRow"))
            If DBMerge_RowHasSelectedData(ws, r, headerMap, plan("headers"), CStr(plan("keyHeader"))) Then
                outWs.Cells(outRow, 1).Value = ws.Name
                outWs.Cells(outRow, 2).Value = r
                outWs.Cells(outRow, 3).Value = DBMerge_KeyValueForRow(ws, r, headerMap, CStr(plan("keyHeader")))

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

Private Function DBMerge_RowHasSelectedData(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerMap As Object, _
                                            ByVal headers As Collection, Optional ByVal keyHeader As String = "") As Boolean
    Dim i As Long
    Dim keyNorm As String
    keyNorm = DBMerge_NormalizeHeader(keyHeader)
    For i = 1 To headers.Count
        Dim norm As String
        norm = DBMerge_NormalizeHeader(CStr(headers(i)))
        If Len(keyNorm) > 0 And norm = keyNorm Then
            ' KEY만 있는 행을 데이터 행으로 오인하지 않도록 데이터 판단에서는 제외한다.
        ElseIf headerMap.Exists(norm) Then
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

Private Function DBMerge_KeyValueForRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerMap As Object, ByVal keyHeader As String) As String
    Dim keyValue As String
    keyValue = Trim$(CStr(DBMerge_CellByHeader(ws, rowNo, headerMap, keyHeader)))
    If Len(keyValue) > 0 Then
        DBMerge_KeyValueForRow = keyValue
    Else
        DBMerge_KeyValueForRow = DBMerge_AutoId(ws.Name, rowNo)
    End If
End Function

Private Function DBMerge_AutoId(ByVal sheetName As String, ByVal rowNo As Long) As String
    DBMerge_AutoId = DBMerge_CleanIdPart(sheetName) & "_R" & CStr(rowNo)
End Function

Private Function DBMerge_CleanIdPart(ByVal textValue As String) As String
    Dim result As String
    result = Trim$(textValue)
    result = Replace(result, " ", "_")
    result = Replace(result, "'", "")
    result = Replace(result, "!", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    If Len(result) = 0 Then result = "Sheet"
    DBMerge_CleanIdPart = result
End Function

Private Sub DBMerge_UpdateIdStats(ByVal sheetPlan As Object, ByVal keyHeader As String, _
                                  ByVal idMode As String, ByVal headers As Collection)
    Dim ws As Worksheet
    Set ws = sheetPlan("worksheet")
    Dim headerMap As Object
    Set headerMap = sheetPlan("headerMap")

    Dim existingCount As Long
    Dim generatedCount As Long
    Dim blankFilledCount As Long
    Dim r As Long
    For r = CLng(sheetPlan("dataStartRow")) To CLng(sheetPlan("lastRow"))
        If DBMerge_RowHasSelectedData(ws, r, headerMap, headers, keyHeader) Then
            Dim keyValue As String
            keyValue = Trim$(CStr(DBMerge_CellByHeader(ws, r, headerMap, keyHeader)))
            If Len(keyValue) > 0 Then
                existingCount = existingCount + 1
            Else
                generatedCount = generatedCount + 1
                If headerMap.Exists(DBMerge_NormalizeHeader(keyHeader)) Then blankFilledCount = blankFilledCount + 1
            End If
        End If
    Next r

    sheetPlan.Add "idExistingCount", existingCount
    sheetPlan.Add "idGeneratedCount", generatedCount
    sheetPlan.Add "idBlankFilledCount", blankFilledCount
End Sub

Private Sub DBMerge_ApplySourceIds(ByVal plan As Object)
    Dim i As Long
    For i = 1 To plan("sheets").Count
        Dim sheetPlan As Object
        Set sheetPlan = plan("sheets")(i)

        Dim ws As Worksheet
        Set ws = sheetPlan("worksheet")
        Dim headerMap As Object
        Set headerMap = sheetPlan("headerMap")
        Dim keyHeader As String
        keyHeader = CStr(plan("keyHeader"))
        Dim keyNorm As String
        keyNorm = DBMerge_NormalizeHeader(keyHeader)

        Dim keyCol As Long
        If headerMap.Exists(keyNorm) Then
            keyCol = CLng(headerMap(keyNorm))
        Else
            keyCol = 1
            ws.Columns(keyCol).Insert Shift:=xlToRight
            DBMerge_WriteSourceIdHeader ws, CLng(sheetPlan("headerRow")), CStr(plan("headerMode")), keyHeader
            Set headerMap = DBMerge_HeaderMap(ws, CLng(sheetPlan("headerRow")), CStr(plan("headerMode")))
            sheetPlan.Remove "headerMap"
            sheetPlan.Add "headerMap", headerMap
        End If

        Dim r As Long
        For r = CLng(sheetPlan("dataStartRow")) To CLng(sheetPlan("lastRow"))
            If DBMerge_RowHasSelectedData(ws, r, headerMap, plan("headers"), keyHeader) Then
                If Len(Trim$(CStr(ws.Cells(r, keyCol).Value))) = 0 Then
                    ws.Cells(r, keyCol).NumberFormat = "@"
                    ws.Cells(r, keyCol).Value = DBMerge_AutoId(ws.Name, r)
                End If
            End If
        Next r
    Next i
End Sub

Private Sub DBMerge_WriteSourceIdHeader(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                        ByVal headerMode As String, ByVal keyHeader As String)
    Select Case headerMode
        Case "display_next_key"
            ws.Cells(headerRow, 1).Value = keyHeader
            ws.Cells(headerRow + 1, 1).Value = keyHeader
        Case "key_prev_display"
            ws.Cells(headerRow - 1, 1).Value = keyHeader
            ws.Cells(headerRow, 1).Value = keyHeader
        Case Else
            ws.Cells(headerRow, 1).Value = keyHeader
    End Select
End Sub

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
        "9. 법정동 마스터 양식 생성" & vbCrLf & _
        "10. 네이버 공식몰 링크 검색", _
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
        Case "10": NaverOfficialMall_FindLinksForSelection
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "주소/검색", "1~10 중 하나를 입력하세요."
    End Select
End Sub

Public Sub Toolkit_LaunchRosterQuality()
    Dim choice As String
    choice = Toolkit_LauncherPrompt( _
        "명부 품질", _
        "실행할 명부 품질 기능을 선택하세요." & vbCrLf & vbCrLf & _
        "1. 명부 검증 시트 만들기" & vbCrLf & _
        "2. 품질 진단" & vbCrLf & _
        "3. 중복 탐지" & vbCrLf & _
        "4. 업체명 매칭" & vbCrLf & _
        "5. 전화번호 검사" & vbCrLf & _
        "6. 사업자등록 상태조회", _
        "1")

    Select Case choice
        Case "1": Roster_CreateValidationWorkspace
        Case "2": Roster_QualityReport
        Case "3": Roster_DetectDuplicates
        Case "4": Roster_MatchCompanyNames
        Case "5": Phone_ValidateSelection
        Case "6": BizNo_QueryStatusForSelection
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "명부 품질", "1~6 중 하나를 입력하세요."
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
        "5. 모집단 구성표 리스트화(자동/직접 지정)" & vbCrLf & _
        "6. 리스트→집계표" & vbCrLf & _
        "7. 표본설계 가이드 열기" & vbCrLf & _
        "8. 숫자값 구간 만들기", _
        "1")

    Select Case choice
        Case "1": SampleDesign_CreateTemplate
        Case "2": SampleDesign_RunAllocation
        Case "3": SampleDesign_CreateProgressTracker
        Case "4": SampleDesign_ReviewQuota
        Case "5": SampleDesign_ListifyPopulationTable
        Case "6": SampleDesign_PivotPopulationList
        Case "7": SampleDesign_ShowGuide
        Case "8": SampleDesign_CreateBandColumn
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "표본설계", "1~8 중 하나를 입력하세요."
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

Public Sub Ribbon_NaverOfficialMallFindLinks(control As IRibbonControl)
    NaverOfficialMall_FindLinksForSelection
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

Public Sub Ribbon_RosterValidationSheet(control As IRibbonControl)
    Roster_CreateValidationWorkspace
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

Public Sub Ribbon_SampleDesignPivotPopulationList(control As IRibbonControl)
    SampleDesign_PivotPopulationList
End Sub

Public Sub Ribbon_SampleDesignRunAllocation(control As IRibbonControl)
    SampleDesign_RunAllocation
End Sub

Public Sub Ribbon_SampleDesignCreateProgressTracker(control As IRibbonControl)
    SampleDesign_CreateProgressTracker
End Sub

Public Sub Ribbon_SampleDesignShowGuide(control As IRibbonControl)
    SampleDesign_ShowGuide
End Sub

Public Sub Ribbon_SampleDesignCreateBandColumn(control As IRibbonControl)
    SampleDesign_CreateBandColumn
End Sub

Public Sub Ribbon_DatabaseMergeSheets(control As IRibbonControl)
    Database_MergeSheetsByHeaders
End Sub

Public Sub Ribbon_DatabaseCreateMergeMarkerRows(control As IRibbonControl)
    Database_CreateMergeMarkerRows
End Sub

Public Sub Ribbon_DatabaseSplitMergedSheet(control As IRibbonControl)
    Database_SplitMergedSheetBySource
End Sub
