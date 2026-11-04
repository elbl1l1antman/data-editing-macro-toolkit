Attribute VB_Name = "MacroToolkit"
' ============================================================
' 모듈명  : MacroToolkit
' 설  명  : 엑셀 데이터 편집을 위한 매크로 도구 모음
'           - 텍스트 형식 숫자를 실제 숫자로 일괄 변환
'           - 혼용된 한국 전화번호 형식을 표준 형식으로 정규화
' ============================================================
Option Explicit

' 현재 툴킷 버전 (외부 모듈에서도 참조 가능하도록 Public 선언)
Public Const TOOLKIT_VERSION As String = "1.2.6"

' 버전 정보를 기록할 숨김 시트 이름 (언더스코어 접두어로 내부 시트임을 표시)
Public Const VERSION_SHEET As String = "_MacroToolkit"
Public Const LOG_SHEET As String = "_MacroToolkit_Log"
Public Const BIZ_API_BASE_URL As String = "https://api.odcloud.kr/api/nts-businessman/v1/status"
Public Const BIZ_APP_NAME As String = "DataEditingMacroToolkit"
Public Const BIZ_APP_SECTION As String = "BusinessApi"
Public Const BIZ_APP_KEY As String = "ServiceKey"
Public Const NAVER_API_BASE_URL As String = "https://openapi.naver.com/v1/search/local.json"
Public Const NAVER_WEB_API_BASE_URL As String = "https://openapi.naver.com/v1/search/webkr.json"
Public Const NAVER_APP_NAME As String = "DataEditingMacroToolkit"
Public Const NAVER_APP_SECTION As String = "NaverLocalSearch"
Public Const NAVER_CLIENT_ID_KEY As String = "ClientId"
Public Const NAVER_CLIENT_SECRET_KEY As String = "ClientSecret"
Public Const NAVER_BUILD_INDUSTRY_KEY As String = "BuildIndustry"
Public Const NAVER_BUILD_DISPLAY_KEY As String = "BuildDisplayCount"
Public Const NAVER_BUILD_DELAY_KEY As String = "BuildDelaySeconds"
Public Const NAVER_BUILD_SORT_KEY As String = "BuildSortMode"
Public Const NAVER_BUILD_BASIS_KEY As String = "BuildDongBasis"
Public Const NAVER_BUILD_SIDO_KEY As String = "BuildSidoFilter"
Public Const NAVER_BUILD_SIGUNGU_KEY As String = "BuildSigunguFilter"
Public Const NAVER_MALL_DISPLAY_KEY As String = "OfficialMallDisplayCount"
Public Const NAVER_MALL_DELAY_KEY As String = "OfficialMallDelaySeconds"
Public Const NAVER_BUSINESS_KEYWORD_KEY As String = "BusinessSearchKeyword"
Public Const NAVER_DONG_MASTER_SHEET As String = "_NaverLocalDongMaster"
Public Const NAVER_USER_DONG_MASTER_SHEET As String = "행정동_마스터"
Public Const NAVER_LEGAL_DONG_MASTER_SHEET As String = "_NaverLocalLegalDongMaster"
Public Const NAVER_USER_LEGAL_DONG_MASTER_SHEET As String = "법정동_마스터"
Public Const API_RESOLVE_TIMEOUT_MS As Long = 5000
Public Const API_CONNECT_TIMEOUT_MS As Long = 10000
Public Const API_SEND_TIMEOUT_MS As Long = 15000
Public Const API_RECEIVE_TIMEOUT_MS As Long = 30000
Public Const TOOLKIT_PROGRESS_ROW_THRESHOLD As Long = 1000
Public Const TOOLKIT_PROGRESS_BATCH_ROWS As Long = 500
Public Const TOOLKIT_PROGRESS_PAIR_BATCH As Long = 10000
Public Const TOOLKIT_PROGRESS_REFRESH_SECONDS As Double = 0.1
Public Const DBMERGE_OUTPUT_BATCH_ROWS As Long = 500
Public Const TOOLKIT_CANCEL_ERROR_NUMBER As Long = vbObjectError + 9130

Private mToolkitOperationDepth As Long
Private mToolkitPrevScreenUpdating As Boolean
Private mToolkitPrevEnableEvents As Boolean
Private mToolkitPrevDisplayAlerts As Boolean
Private mToolkitPrevCalculation As XlCalculation
Private mToolkitPrevStatusBar As Variant

Private gProgressForm As Object
Private mToolkitCancelRequested As Boolean
Private mToolkitLastProgressRefresh As Double

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

Public Function Toolkit_CompileSmokeTest() As String
    Toolkit_CompileSmokeTest = TOOLKIT_VERSION & "|forms=14"
End Function

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
    ws.Range("A53").Value = "NaverBusinessSearch_SearchSelection"
    ws.Range("A54").Value = "DataCheck_CreateCharacterCheckFormula"
    ws.Range("A55").Value = "Address_ConvertProvinceNotation"
    ws.Range("A56").Value = "DataCleanup_ClearResidualBlankValues"
    ws.Range("A57").Value = "DataCleanup_ClearErrorAndZeroValues"

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

Public Sub Toolkit_ProgressShow(ByVal title As String, ByVal message As String, ByVal totalCount As Long)
    On Error Resume Next
    mToolkitCancelRequested = False
    Set gProgressForm = VBA.UserForms.Add("frmProgressStatus")
    If Not gProgressForm Is Nothing Then
        gProgressForm.Configure title, totalCount, message
        gProgressForm.Show vbModeless
    End If
    Application.StatusBar = title & " - " & message
    mToolkitLastProgressRefresh = Timer
    DoEvents
    On Error GoTo 0
End Sub

Public Sub Toolkit_ProgressUpdate(ByVal currentCount As Long, ByVal totalCount As Long, ByVal message As String)
    Dim cancelRequested As Boolean

    If currentCount < totalCount Then
        If Not Toolkit_ShouldRefreshProgress() Then Exit Sub
    End If

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

Private Function Toolkit_ShouldRefreshProgress() As Boolean
    Dim currentTick As Double
    currentTick = Timer
    If currentTick < mToolkitLastProgressRefresh Or _
       currentTick - mToolkitLastProgressRefresh >= TOOLKIT_PROGRESS_REFRESH_SECONDS Then
        mToolkitLastProgressRefresh = currentTick
        Toolkit_ShouldRefreshProgress = True
    End If
End Function

Public Sub Toolkit_ProgressClose()
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

Public Function Toolkit_ShouldShowRowProgress(ByVal totalRows As Long) As Boolean
    Toolkit_ShouldShowRowProgress = (totalRows > 0)
End Function

Public Sub Toolkit_RowProgressStart(ByVal title As String, ByVal totalRows As Long, ByVal message As String)
    If Toolkit_ShouldShowRowProgress(totalRows) Then Toolkit_ProgressShow title, message, totalRows
End Sub

Public Sub Toolkit_RowProgressUpdate(ByVal processedRows As Long, ByVal totalRows As Long, ByVal message As String)
    If Not Toolkit_ShouldShowRowProgress(totalRows) Then Exit Sub
    Toolkit_ProgressUpdate processedRows, totalRows, message
End Sub

Public Sub Toolkit_RowProgressClose(ByVal totalRows As Long)
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

Public Function Toolkit_IsCancelError(ByVal errorNumber As Long) As Boolean
    Toolkit_IsCancelError = (errorNumber = TOOLKIT_CANCEL_ERROR_NUMBER)
End Function

Public Function Toolkit_ShowCancelMessage(ByVal title As String) As Boolean
    If Not Toolkit_IsCancelError(Err.Number) Then Exit Function
    MsgBox "작업이 취소되었습니다.", vbInformation, title
    Toolkit_ShowCancelMessage = True
End Function

Public Sub Toolkit_RaiseIfCancelled()
    If Toolkit_ProgressCancelRequested() Then
        Err.Raise TOOLKIT_CANCEL_ERROR_NUMBER, "DataEditingMacroToolkit", "사용자가 작업을 취소했습니다."
    End If
End Sub

Public Sub Toolkit_SelectA1OnAllSheets()
    Toolkit_ResetAllSheetsToTop
End Sub

Public Function Toolkit_GetRangeFromUserForm(ByVal formTitle As String, _
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

Public Sub Toolkit_BeginOperation(Optional ByVal statusMessage As String = "")
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

Public Sub Toolkit_EndOperation()
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

Public Function Toolkit_CreateRangeSnapshot(ByVal targetRange As Range) As Collection
    Dim snapshots As New Collection
    Dim area As Range

    For Each area In targetRange.Areas
        Dim item As Object
        Set item = CreateObject("Scripting.Dictionary")
        item.Add "worksheet", area.Worksheet
        item.Add "address", area.Address
        item.Add "formula2", area.Formula2
        snapshots.Add item
    Next area

    Set Toolkit_CreateRangeSnapshot = snapshots
End Function

Public Sub Toolkit_RestoreRangeSnapshot(ByVal snapshots As Collection)
    If snapshots Is Nothing Then Exit Sub

    Dim item As Variant
    For Each item In snapshots
        Dim ws As Worksheet
        Set ws = item("worksheet")
        ws.Range(CStr(item("address"))).Formula2 = item("formula2")
    Next item
End Sub

Public Function Toolkit_ValidateWorksheet(ByVal ws As Worksheet, ByVal purpose As String, _
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

Public Function Toolkit_CanInsertColumns(ByVal ws As Worksheet, ByVal startCol As Long, _
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

Public Function Toolkit_CanDeleteColumns(ByVal ws As Worksheet, ByVal startCol As Long, _
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

Public Function Toolkit_CanInsertRows(ByVal ws As Worksheet, ByVal startRow As Long, _
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

Public Function Toolkit_CanDeleteRows(ByVal ws As Worksheet, ByVal startRow As Long, _
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

Public Sub Toolkit_LogApiEvent(ByVal wb As Workbook, ByVal sourceName As String, ByVal actionName As String, _
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

Public Function Toolkit_GetActualDataRange(ByVal ws As Worksheet) As Range
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

Public Function Toolkit_TrimRangeToActualData(ByVal target As Range) As Range
    Dim actualData As Range
    Set actualData = Toolkit_GetActualDataRange(target.Worksheet)
    If actualData Is Nothing Then Exit Function
    Set Toolkit_TrimRangeToActualData = Intersect(target, actualData)
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
        "7. 기업/사업체명 정리" & vbCrLf & _
        "8. 문자 검진 함수 생성" & vbCrLf & _
        "9. 수식 빈값 정리" & vbCrLf & _
        "10. 오류값/0 지우기", _
        "1")

    Select Case choice
        Case "1": ConvertTextNumbersWithForm
        Case "2": ConvertTextNumbersInCurrentSheetDataRange
        Case "3": NormalizePhonesWithForm
        Case "4": RemovePhoneDashesToCopiedColumns
        Case "5": RemoveBizNoDashesToCopiedColumns
        Case "6": BizNo_MergeThreeColumns
        Case "7": CompanyName_CleanWithOptions
        Case "8": DataCheck_CreateCharacterCheckFormula
        Case "9": DataCleanup_ClearResidualBlankValues
        Case "10": DataCleanup_ClearErrorAndZeroValues
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "값 정리", "1~10 중 하나를 선택하세요."
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
        "10. 네이버 공식몰 링크 검색" & vbCrLf & _
        "11. 네이버 업체 검색" & vbCrLf & _
        "12. 광역시도 표기 변환", _
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
        Case "11": NaverBusinessSearch_SearchSelection
        Case "12": Address_ConvertProvinceNotation
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "주소/검색", "1~12 중 하나를 입력하세요."
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
        "8. 숫자값 구간 만들기" & vbCrLf & _
        "9. 표본 추출", _
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
        Case "9": SampleExtract_Run
        Case Else
            If Len(choice) > 0 Then Toolkit_ShowLauncherError "표본설계", "1~9 중 하나를 선택하세요."
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

Public Function Toolkit_ButtonMenu(ByVal title As String, ByVal guideText As String, _
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

Public Function Toolkit_OptionsFromNumberedPrompt(ByVal promptText As String) As String
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

Public Function Toolkit_CheckListMenu(ByVal title As String, ByVal guideText As String, _
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

Public Sub Ribbon_DataCheckCreateCharacterFormula(control As IRibbonControl)
    DataCheck_CreateCharacterCheckFormula
End Sub

Public Sub Ribbon_DataCleanupClearResidualBlanks(control As IRibbonControl)
    DataCleanup_ClearResidualBlankValues
End Sub

Public Sub Ribbon_DataCleanupClearErrorsAndZeros(control As IRibbonControl)
    DataCleanup_ClearErrorAndZeroValues
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

Public Sub Ribbon_AddressConvertProvinceNotation(control As IRibbonControl)
    Address_ConvertProvinceNotation
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

Public Sub Ribbon_NaverBusinessSearch(control As IRibbonControl)
    NaverBusinessSearch_SearchSelection
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

Public Sub Ribbon_SampleExtractRun(control As IRibbonControl)
    SampleExtract_Run
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
