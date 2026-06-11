Attribute VB_Name = "ReportAutomationAddin"
' ============================================================
' 모듈명  : ReportAutomationAddin
' 설  명  : 집계표 기반 보고서 산출 시트 생성 전용 추가기능
' ============================================================
Option Explicit

' 추가기능 버전. 산출 메타 시트와 로그 시트에 함께 기록해 결과 파일의 생성 기준을 추적한다.
Public Const REPORT_AUTOMATION_VERSION As String = "0.1.0"

' 사용자 통합문서 안에 생성하는 내부 로그 시트명. 일반 사용자에게는 보이지 않도록 VeryHidden 처리한다.
Private Const REPORT_LOG_SHEET As String = "_ReportAutomation_Log"

' 단순 분포표(좁은 표)와 배너 교차표를 구분하는 열 수 기준값.
Private Const SIMPLE_TABLE_MAX_COL As Long = 6

' 단순 분포표에서 % 열 위치를 헤더에서 못 찾을 때 쓰는 기본 열 번호.
Private Const DEFAULT_PCT_COL As Long = 5

' 분석문 본문에 포함하는 상위 항목 수 (1위 제외 보조 문장용).
Private Const NARRATIVE_TOP_ITEMS As Long = 4

' 수치 요약 문자열에 포함하는 상위 항목 수.
Private Const SUMMARY_TOP_ITEMS As Long = 5

' 차트 데이터 시트에서 "포함" 기본값으로 처리하는 최대 항목 순위.
Private Const CHART_INCLUDE_TOP_N As Long = 8

' tableRec 배열 인덱스 상수.
Private Const IDX_TABLE_KEY  As Long = 0
Private Const IDX_TABLE_NO   As Long = 1
Private Const IDX_TITLE      As Long = 2
Private Const IDX_BASIS      As Long = 3
Private Const IDX_BASE_LABEL As Long = 4
Private Const IDX_START_ROW  As Long = 5
Private Const IDX_END_ROW    As Long = 6
Private Const IDX_LAST_COL   As Long = 7
Private Const IDX_TABLE_TYPE As Long = 8
Private Const IDX_WARNING    As Long = 9

' Application 상태를 일시 변경하기 전 원래 값을 저장하는 모듈 전역 변수들.
' 오류가 발생해도 ReportAutomation_EndOperation에서 최대한 원복한다.
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mPrevDisplayAlerts As Boolean
Private mPrevCalculation As XlCalculation
Private mPrevStatusBar As Variant

' UserForm에서 전달한 1회성 실행 옵션. 설정 시트를 생성한 뒤 이 값으로 덮어쓴다.
Private mHasOptionOverrides As Boolean
Private mOverrideBannerSetting As String
Private mOverrideTitlePrefixes As String

' ============================================================
' 프로시저 : ReportAutomation_About
' 설  명  : 추가기능 버전과 현재 제공 기능을 사용자에게 안내한다.
' 호출처  : 리본 About 버튼 또는 매크로 목록
' ============================================================
Public Sub ReportAutomation_About()
    MsgBox "보고서 자동화 추가기능" & vbCrLf & _
           "버전: " & REPORT_AUTOMATION_VERSION & vbCrLf & vbCrLf & _
           "기능:" & vbCrLf & _
           "- 집계표 제목([표 ...]) 자동 탐지" & vbCrLf & _
           "- 보고서용 표 목록, 분석문, 차트데이터, 삽입표 생성" & vbCrLf & _
           "- QA/출처/수정이력/메타 시트 생성", _
           vbInformation, "보고서 자동화"
End Sub

' ============================================================
' 프로시저 : ReportAutomation_GenerateExcelOutputs
' 설  명  : 사용자가 직접 실행하는 공개 매크로.
'           활성 통합문서의 집계표를 읽어 보고서 산출 시트를 생성한다.
' 주  의  : 추가기능 파일이 아니라 데이터 통합문서가 활성화되어 있어야 한다.
' ============================================================
Public Sub ReportAutomation_GenerateExcelOutputs()
    ReportAutomation_GenerateExcelOutputsCore False
End Sub

' ============================================================
' 프로시저 : ReportAutomation_ShowOptions
' 설  명  : 사용자 설정 폼을 열어 옵션을 받은 뒤 보고서 산출을 실행한다.
' 주  의  : 폼이 빌드에 포함되지 않은 환경에서는 기본 실행으로 fallback한다.
' ============================================================
Public Sub ReportAutomation_ShowOptions()
    On Error GoTo Fallback
    ReportAutomationOptionsForm.Show
    Exit Sub

Fallback:
    ReportAutomation_GenerateExcelOutputs
End Sub

' ============================================================
' 함수명 : ReportAutomation_RunWithOptions
' 설  명 : UserForm에서 받은 옵션을 1회성 override로 저장하고 산출 생성을 실행한다.
' 반환값 : 실행 요청을 정상 접수했으면 True
' ============================================================
Public Function ReportAutomation_RunWithOptions(ByVal bannerSetting As String, ByVal titlePrefixes As String) As Boolean
    mHasOptionOverrides = True
    mOverrideBannerSetting = Trim$(bannerSetting)
    mOverrideTitlePrefixes = Trim$(titlePrefixes)

    If Len(mOverrideBannerSetting) = 0 Then mOverrideBannerSetting = "전체"
    If Len(mOverrideTitlePrefixes) = 0 Then mOverrideTitlePrefixes = "진흥지구별,사업체 특성별,산업별"

    ReportAutomation_GenerateExcelOutputsCore False
    ReportAutomation_RunWithOptions = True
End Function

' ============================================================
' 함수명 : ReportAutomation_DefaultBannerSetting
' 설  명 : UserForm 초기값으로 사용할 배너 설정을 반환한다.
' ============================================================
Public Function ReportAutomation_DefaultBannerSetting() As String
    Dim ws As Worksheet
    If Not ActiveWorkbook Is Nothing Then Set ws = ReportAutomation_FindLatestOutputSheet(ActiveWorkbook, "보고서_설정")
    ReportAutomation_DefaultBannerSetting = CStr(ReportAutomation_SettingValue("추출 배너 목록", "전체", ws))
End Function

' ============================================================
' 함수명 : ReportAutomation_DefaultTitlePrefixes
' 설  명 : UserForm 초기값으로 사용할 제목 제거 접두어 목록을 반환한다.
' ============================================================
Public Function ReportAutomation_DefaultTitlePrefixes() As String
    Dim ws As Worksheet
    If Not ActiveWorkbook Is Nothing Then Set ws = ReportAutomation_FindLatestOutputSheet(ActiveWorkbook, "보고서_설정")
    ReportAutomation_DefaultTitlePrefixes = CStr(ReportAutomation_SettingValue("제목 제거 접두어", "진흥지구별,사업체 특성별,산업별", ws))
End Function

' ============================================================
' 프로시저 : ReportAutomation_GenerateExcelOutputsSilent
' 설  명  : 자동 테스트/외부 호출용 무인 실행 매크로.
'           MsgBox 대신 Err.Raise로 실패 사유를 호출자에게 전달한다.
' ============================================================
Public Sub ReportAutomation_GenerateExcelOutputsSilent()
    ReportAutomation_GenerateExcelOutputsCore True
End Sub

' ============================================================
' 프로시저 : Ribbon_ReportAutomationGenerate
' 설  명  : 리본 버튼에서 산출 시트 생성을 호출하는 콜백.
' ============================================================
Public Sub Ribbon_ReportAutomationGenerate(control As IRibbonControl)
    ReportAutomation_ShowOptions
End Sub

' ============================================================
' 프로시저 : Ribbon_ReportAutomationAbout
' 설  명  : 리본 버튼에서 추가기능 정보를 표시하는 콜백.
' ============================================================
Public Sub Ribbon_ReportAutomationAbout(control As IRibbonControl)
    ReportAutomation_About
End Sub

' ============================================================
' 프로시저 : ReportAutomation_GenerateExcelOutputsCore
' 설  명  : 보고서 자동화의 메인 실행 루틴.
' 흐  름  : 대상 통합문서 확인 → 집계표 탐지 → 산출 시트 생성
'           → 분석문/차트데이터/삽입표/QA/메타 작성 → 로그 기록
' 인  자  : silent=True이면 UI 메시지 없이 오류를 호출자에게 다시 발생시킨다.
' ============================================================
Private Sub ReportAutomation_GenerateExcelOutputsCore(ByVal silent As Boolean)
    On Error GoTo ErrHandler

    ' 실행 대상은 반드시 사용자가 연 데이터 통합문서여야 한다.
    ' ThisWorkbook은 추가기능 자신이므로 여기에 산출 시트를 만들면 안 된다.
    If ActiveWorkbook Is Nothing Then Exit Sub
    If ActiveWorkbook Is ThisWorkbook Then
        If silent Then
            Err.Raise vbObjectError + 7200, "ReportAutomation_GenerateExcelOutputs", "보고서 산출 대상 통합문서를 먼저 활성화하세요."
        Else
            MsgBox "보고서 산출 대상 통합문서를 먼저 활성화하세요.", vbExclamation, "보고서 자동화"
        End If
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' 우선 활성 시트, 그 다음 Sheet1, 마지막으로 모든 시트를 훑어 집계표 원본 시트를 찾는다.
    Dim dataWs As Worksheet
    Set dataWs = ReportAutomation_ResolveDataSheet(wb)
    If dataWs Is Nothing Then
        If silent Then
            Err.Raise vbObjectError + 7201, "ReportAutomation_GenerateExcelOutputs", "집계표 제목([표 ...])이 있는 시트를 찾지 못했습니다."
        Else
            MsgBox "집계표 제목([표 ...])이 있는 시트를 찾지 못했습니다.", vbExclamation, "보고서 자동화"
        End If
        Exit Sub
    End If

    ' 원본 시트에서 [표 ...] 제목행을 기준으로 표 블록 메타데이터를 수집한다.
    Dim tables As Collection
    Set tables = ReportAutomation_CollectTables(dataWs)
    If tables.Count = 0 Then
        If silent Then
            Err.Raise vbObjectError + 7202, "ReportAutomation_GenerateExcelOutputs", "집계표 제목([표 ...])을 찾지 못했습니다."
        Else
            MsgBox "집계표 제목([표 ...])을 찾지 못했습니다.", vbExclamation, "보고서 자동화"
        End If
        Exit Sub
    End If

    ReportAutomation_BeginOperation "보고서 자동화 산출 시트를 생성하는 중입니다..."

    ' 산출물은 한 번의 실행마다 timestamp가 붙은 새 시트로 생성한다.
    ' 기존 산출 시트를 덮어쓰지 않아 사용자가 이전 결과와 비교할 수 있다.
    Dim wsSettings As Worksheet, wsList As Worksheet, wsNarr As Worksheet
    Dim wsChart As Worksheet, wsInsert As Worksheet, wsQA As Worksheet
    Dim wsSource As Worksheet, wsRevision As Worksheet, wsMeta As Worksheet
    Dim wsPriorSettings As Worksheet

    ' 사용자가 이전 산출 설정 시트를 편집한 경우, 새 실행에서도 그 값을 승계한다.
    Set wsPriorSettings = ReportAutomation_FindLatestOutputSheet(wb, "보고서_설정")

    Set wsSettings = ReportAutomation_AddOutputSheet(wb, "보고서_설정")
    Set wsList = ReportAutomation_AddOutputSheet(wb, "보고서_표목록")
    Set wsNarr = ReportAutomation_AddOutputSheet(wb, "보고서_분석문")
    Set wsChart = ReportAutomation_AddOutputSheet(wb, "보고서_차트데이터")
    Set wsInsert = ReportAutomation_AddOutputSheet(wb, "보고서_삽입표")
    Set wsQA = ReportAutomation_AddOutputSheet(wb, "보고서_QA")
    Set wsSource = ReportAutomation_AddOutputSheet(wb, "보고서_출처")
    Set wsRevision = ReportAutomation_AddOutputSheet(wb, "보고서_수정이력")
    Set wsMeta = ReportAutomation_AddOutputSheet(wb, "_ReportMeta")

    ' 각 산출 시트 작성은 역할별 전용 함수로 분리해 후속 리뷰와 테스트 범위를 좁힌다.
    ReportAutomation_WriteSettingsSheet wsSettings, wb, dataWs, wsPriorSettings
    If mHasOptionOverrides Then
        ReportAutomation_SetSettingValue wsSettings, "추출 배너 목록", mOverrideBannerSetting
        ReportAutomation_SetSettingValue wsSettings, "제목 제거 접두어", mOverrideTitlePrefixes
    End If
    ReportAutomation_WriteTableList wsList, tables, dataWs
    ReportAutomation_WriteNarratives wsNarr, wsChart, wsInsert, wsQA, tables, dataWs, wsSettings
    ReportAutomation_WriteSourceSheet wsSource
    ReportAutomation_WriteRevisionSheet wsRevision
    ReportAutomation_WriteMetaSheet wsMeta, wb, dataWs, tables, wsSettings, wsList, wsNarr, wsChart, wsInsert

    ' 메타 시트는 Python/HWP/PPT 후속 자동화가 참조하는 내부 정보이므로 숨김 처리한다.
    wsMeta.Visible = xlSheetVeryHidden

    ReportAutomation_LogEvent wb, "generate_excel_outputs", dataWs.Name, "OK", tables.Count & "개 표 탐지"
    ReportAutomation_EndOperation

    If Not silent Then
        MsgBox "보고서 자동화 산출 시트를 생성했습니다." & vbCrLf & _
               "원본 시트: " & dataWs.Name & vbCrLf & _
               "탐지 표 수: " & tables.Count, vbInformation, "보고서 자동화"
    End If
    ReportAutomation_ClearOptionOverrides
    Exit Sub

ErrHandler:
    ' EndOperation 내부의 On Error Resume Next가 Err를 소거하기 전에 저장한다.
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    ReportAutomation_EndOperation
    On Error Resume Next
    ReportAutomation_LogEvent wb, "generate_excel_outputs", "", "ERROR", errDesc
    On Error GoTo 0
    If silent Then
        Err.Raise errNum, errSrc, errDesc
    Else
        MsgBox "보고서 자동화 산출 시트 생성 중 오류가 발생했습니다." & vbCrLf & errDesc, vbExclamation, "보고서 자동화"
    End If
    ReportAutomation_ClearOptionOverrides
End Sub

' ============================================================
' 함수명 : ReportAutomation_ResolveDataSheet
' 설  명 : 집계표 원본 시트를 결정한다.
' 우선순위: 현재 활성 시트 → Sheet1 → 통합문서 내 첫 번째 집계표 포함 시트
' 반환값 : [표 ...] 제목행이 있는 Worksheet, 없으면 Nothing
' ============================================================
Private Function ReportAutomation_ResolveDataSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    ' 사용자가 보고 있던 시트를 가장 먼저 신뢰한다.
    If TypeName(ActiveSheet) = "Worksheet" Then
        Set ws = ActiveSheet
        If Not ws.Parent Is Nothing Then
            If ws.Parent Is wb And ReportAutomation_CountTableTitleRows(ws) > 0 Then
                Set ReportAutomation_ResolveDataSheet = ws
                Exit Function
            End If
        End If
    End If

    ' SPSS/조사 집계표 샘플은 Sheet1에 본표가 있는 경우가 많아 두 번째 후보로 본다.
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Sheet1")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If ReportAutomation_CountTableTitleRows(ws) > 0 Then
            Set ReportAutomation_ResolveDataSheet = ws
            Exit Function
        End If
    End If

    ' 시트명이 바뀐 파일도 처리하기 위해 모든 워크시트를 마지막 후보로 순회한다.
    For Each ws In wb.Worksheets
        If ReportAutomation_CountTableTitleRows(ws) > 0 Then
            Set ReportAutomation_ResolveDataSheet = ws
            Exit Function
        End If
    Next ws
End Function

' ============================================================
' 함수명 : ReportAutomation_CountTableTitleRows
' 설  명 : A열에서 [표 ...] 형식의 제목행 개수를 센다.
' 반환값 : 해당 시트가 집계표 원본인지 판단하는 근거 수치
' ============================================================
Private Function ReportAutomation_CountTableTitleRows(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 1 To lastRow
        If ReportAutomation_IsTableTitle(ws.Cells(r, 1).Value) Then
            ReportAutomation_CountTableTitleRows = ReportAutomation_CountTableTitleRows + 1
        End If
    Next r
End Function

' ============================================================
' 함수명 : ReportAutomation_CollectTables
' 설  명 : [표 ...] 제목행을 기준으로 표 블록 범위와 기본 메타데이터를 수집한다.
' 반환값 : Collection(Array(table_key, table_no, title, basis, base_label,
'                          start_row, end_row, last_col, table_type, warning))
' 주  의 : 현재 MVP는 A열 제목행 기준의 연속 블록 구조를 전제로 한다.
' ============================================================
Private Function ReportAutomation_CollectTables(ByVal ws As Worksheet) As Collection
    Dim titleRows As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 1 To lastRow
        If ReportAutomation_IsTableTitle(ws.Cells(r, 1).Value) Then titleRows.Add r
    Next r

    ' 각 제목행부터 다음 제목행 직전까지를 하나의 표 블록으로 본다.
    Dim tables As New Collection
    Dim i As Long
    For i = 1 To titleRows.Count
        Dim startRow As Long, endRow As Long, lastCol As Long
        startRow = CLng(titleRows(i))
        If i < titleRows.Count Then
            endRow = CLng(titleRows(i + 1)) - 1
        Else
            endRow = ReportAutomation_LastNonEmptyRowInBlock(ws, startRow, lastRow)
        End If
        lastCol = ReportAutomation_LastNonEmptyColInBlock(ws, startRow, endRow)

        ' 제목 문자열에서 표번호, 분석 제목, 기준 문구를 분리한다.
        Dim titleText As String, tableNo As String, cleanTitle As String, basis As String
        titleText = CStr(ws.Cells(startRow, 1).Value)
        tableNo = ReportAutomation_ParseTableNo(titleText)
        cleanTitle = ReportAutomation_ParseTableTitle(titleText)
        basis = ReportAutomation_ParseBasis(titleText)

        Dim baseLabel As String
        If startRow + 1 <= endRow Then baseLabel = CStr(ws.Cells(startRow + 1, 1).Value)

        ' 배열 기반 레코드로 저장한다. 후속 단계에서 rec(index)로 빠르게 접근한다.
        tables.Add Array( _
            "T" & Format$(i, "0000"), tableNo, cleanTitle, basis, baseLabel, _
            startRow, endRow, lastCol, _
            ReportAutomation_ClassifyTable(ws, startRow, endRow, lastCol), _
            ReportAutomation_TableWarning(ws, startRow, endRow, lastCol))
    Next i

    Set ReportAutomation_CollectTables = tables
End Function

' ============================================================
' 함수명 : ReportAutomation_IsTableTitle
' 설  명 : 셀 값이 집계표 제목행인지 판정한다.
' 기준   : 왼쪽 공백 제거 후 "[표 "로 시작
' ============================================================
Private Function ReportAutomation_IsTableTitle(ByVal valueData As Variant) As Boolean
    Dim text As String
    text = Trim$(CStr(valueData))
    ReportAutomation_IsTableTitle = (Left$(text, 3) = "[표 ")
End Function

' ============================================================
' 함수명 : ReportAutomation_LastNonEmptyRowInBlock
' 설  명 : 표 블록 안에서 마지막 값이 있는 행을 뒤에서부터 찾는다.
' 용  도 : 마지막 표는 다음 제목행이 없으므로 실제 끝 행을 추정해야 한다.
' ============================================================
Private Function ReportAutomation_LastNonEmptyRowInBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal maxRow As Long) As Long
    Dim r As Long, c As Long
    For r = maxRow To startRow Step -1
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
                ReportAutomation_LastNonEmptyRowInBlock = r
                Exit Function
            End If
        Next c
    Next r
    ReportAutomation_LastNonEmptyRowInBlock = startRow
End Function

' ============================================================
' 함수명 : ReportAutomation_LastNonEmptyColInBlock
' 설  명 : 표 블록 내 사용된 마지막 열 번호를 계산한다.
' 반환값 : 최소 1 이상의 열 번호
' ============================================================
Private Function ReportAutomation_LastNonEmptyColInBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long, c As Long, lastCol As Long
    For r = startRow To endRow
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If c > lastCol Then lastCol = c
    Next r
    If lastCol < 1 Then lastCol = 1
    ReportAutomation_LastNonEmptyColInBlock = lastCol
End Function

' ============================================================
' 함수명 : ReportAutomation_ParseTableNo
' 설  명 : "[표 MQT2 > ...]" 같은 제목에서 표 번호(MQT2)를 추출한다.
' 반환값 : 추출 실패 시 빈 문자열
' ============================================================
Private Function ReportAutomation_ParseTableNo(ByVal titleText As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, titleText, "[표 ", vbTextCompare)
    endPos = InStr(1, titleText, ">", vbTextCompare)
    If startPos > 0 And endPos > startPos Then
        ReportAutomation_ParseTableNo = Trim$(Mid$(titleText, startPos + 3, endPos - startPos - 3))
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_ParseTableTitle
' 설  명 : 표 제목 문자열에서 보고서용 제목 부분만 추출한다.
' 예  시 : "[표 MDBR-6 > [리스트 기준] 대분류...]" → "[리스트 기준] 대분류..."
' ============================================================
Private Function ReportAutomation_ParseTableTitle(ByVal titleText As String) As String
    Dim endPos As Long
    endPos = InStr(1, titleText, ">", vbTextCompare)
    If endPos > 0 Then
        Dim parsedTitle As String
        parsedTitle = Trim$(Mid$(titleText, endPos + 1))
        If Right$(parsedTitle, 1) = "]" Then parsedTitle = Left$(parsedTitle, Len(parsedTitle) - 1)
        ReportAutomation_ParseTableTitle = Trim$(parsedTitle)
    Else
        ReportAutomation_ParseTableTitle = titleText
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_ParseBasis
' 설  명 : 제목 안의 기준 문구([리스트 기준], [응답 기준] 등)를 추출한다.
' 주  의 : 현재는 "> [" 다음 첫 닫는 대괄호까지를 기준 문구로 본다.
' ============================================================
Private Function ReportAutomation_ParseBasis(ByVal titleText As String) As String
    Dim openPos As Long, closePos As Long
    openPos = InStr(1, titleText, "> [", vbTextCompare)
    If openPos > 0 Then
        closePos = InStr(openPos + 3, titleText, "]", vbTextCompare)
        If closePos > openPos Then
            ReportAutomation_ParseBasis = Mid$(titleText, openPos + 2, closePos - openPos - 1)
        End If
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_ClassifyTable
' 설  명 : 표 폭과 키워드로 표 유형을 대략 분류한다.
' 반환값 : 단일 전체 분포표 / 평균 포함 표 / N/% 쌍 표 / 배너 교차표
' ============================================================
Private Function ReportAutomation_ClassifyTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As String
    If lastCol <= SIMPLE_TABLE_MAX_COL Then
        ReportAutomation_ClassifyTable = "단일 전체 분포표"
    ElseIf ReportAutomation_BlockContainsText(ws, startRow, endRow, lastCol, "▶평균◀") Then
        ReportAutomation_ClassifyTable = "평균 포함 표"
    ElseIf ReportAutomation_BlockContainsText(ws, startRow, startRow + 3, lastCol, "N") Then
        ReportAutomation_ClassifyTable = "N/% 쌍 표"
    Else
        ReportAutomation_ClassifyTable = "배너 교차표"
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_BlockContainsText
' 설  명 : 표 블록 범위 안에 특정 문자열이 포함되어 있는지 검사한다.
' 용  도 : 표 유형 분류, 평균 포함 여부, N/% 구조 판정
' ============================================================
Private Function ReportAutomation_BlockContainsText(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal needle As String) As Boolean
    Dim r As Long, c As Long
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count
    For r = startRow To endRow
        For c = 1 To lastCol
            If InStr(1, CStr(ws.Cells(r, c).Value), needle, vbTextCompare) > 0 Then
                ReportAutomation_BlockContainsText = True
                Exit Function
            End If
        Next c
    Next r
End Function

' ============================================================
' 함수명 : ReportAutomation_TableWarning
' 설  명 : 표 구조상 검토가 필요한 항목을 warning 문자열로 만든다.
' 현재검사: BASE 행 존재 여부, 사용 열 존재 여부
' ============================================================
Private Function ReportAutomation_TableWarning(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long) As String
    Dim warningText As String
    If startRow + 1 > endRow Or InStr(1, CStr(ws.Cells(startRow + 1, 1).Value), "BASE:", vbTextCompare) = 0 Then
        warningText = "BASE 행 확인 필요"
    End If
    If lastCol <= 1 Then
        If Len(warningText) > 0 Then warningText = warningText & "; "
        warningText = warningText & "사용 열 확인 필요"
    End If
    ReportAutomation_TableWarning = warningText
End Function

' ============================================================
' 함수명 : ReportAutomation_AddOutputSheet
' 설  명 : 산출 시트를 통합문서 맨 뒤에 추가하고 중복 없는 이름을 부여한다.
' ============================================================
Private Function ReportAutomation_AddOutputSheet(ByVal wb As Workbook, ByVal baseName As String) As Worksheet
    Set ReportAutomation_AddOutputSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ReportAutomation_AddOutputSheet.Name = ReportAutomation_UniqueSheetName(wb, baseName)
End Function

' ============================================================
' 프로시저 : ReportAutomation_WriteSettingsSheet
' 설  명  : 보고서 생성 기준과 후속 자동화 옵션을 설정 시트에 기록한다.
' 용  도  : 사용자가 문체/수치 표기/LLM 사용 여부를 나중에 편집할 수 있게 한다.
' ============================================================
Private Sub ReportAutomation_WriteSettingsSheet(ByVal ws As Worksheet, ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal priorSettingsWs As Worksheet)
    ws.Range("A1").Value = "보고서 자동화 설정"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' 현재 단계에서는 기본값을 먼저 깔아두고, 후속 UI에서 사용자가 수정하는 구조를 염두에 둔다.
    Dim labels As Variant, values As Variant
    labels = Array("프로젝트ID", "프로젝트명", "보고서명", "프로젝트 유형", "산출물 유형", _
                   "문체 프로필", "수치 표기", "N 표기 여부", "비교 기준", "QA 실행 여부", _
                   "출처 관리 여부", "수정 이력 관리 여부", "LLM 문장 고도화", "LLM 제공자", _
                   "LLM API 키 입력 방식", "LLM 모델명", "원본 통합문서", "원본 집계표 시트", _
                   "추출 배너 목록", "제목 제거 접두어")
    values = Array("PRJ_" & Format(Now, "yyyymmdd_hhnn"), "", "", "데이터분석", "Excel 산출 시트", _
                   "공공조사 보고서형", "소수점 1자리", "가중 사례수와 실제 사례수 보존", _
                   "전체 대비", "수치 QA / 문체 QA / 표 서식 QA", "사용 안 함", "사용", _
                   "사용 안 함", "사용 안 함", "실행 시 입력", "", wb.Name, dataWs.Name, _
                   "전체", "진흥지구별,사업체 특성별,산업별")

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        ws.Cells(i + 3, 1).Value = labels(i)
        If labels(i) = "원본 통합문서" Then
            ws.Cells(i + 3, 2).Value = wb.Name
        ElseIf labels(i) = "원본 집계표 시트" Then
            ws.Cells(i + 3, 2).Value = dataWs.Name
        Else
            ws.Cells(i + 3, 2).Value = ReportAutomation_SettingValue(CStr(labels(i)), values(i), priorSettingsWs)
        End If
    Next i

    ws.Columns("A:B").AutoFit
    ws.Range("A3:A" & UBound(labels) + 3).Font.Bold = True
    ReportAutomation_StyleHeader ws.Range("A1:B1")
End Sub

' ============================================================
' 프로시저 : ReportAutomation_WriteTableList
' 설  명  : 탐지된 표 목록과 원본 범위를 한 행당 한 표로 기록한다.
' 용  도  : 사용자가 보고서에 사용할 표를 선택/제외하고, 원본 위치를 추적한다.
' ============================================================
Private Sub ReportAutomation_WriteTableList(ByVal ws As Worksheet, ByVal tables As Collection, ByVal dataWs As Worksheet)
    ws.Range("A1:L1").Value = Array("선택", "table_key", "table_no", "title", "basis", "base_label", _
                                    "source_sheet", "source_range", "table_type", "row_count", "col_count", "warning")
    ReportAutomation_StyleHeader ws.Range("A1:L1")

    Dim i As Long, rec As Variant
    For i = 1 To tables.Count
        rec = tables(i)
        ' 기본값은 전체 표 사용(Y). 사용자는 이 열을 N으로 바꿔 후속 생성 대상에서 제외할 수 있다.
        ws.Cells(i + 1, 1).Value = "Y"
        ws.Cells(i + 1, 2).Value = rec(IDX_TABLE_KEY)
        ws.Cells(i + 1, 3).Value = rec(IDX_TABLE_NO)
        ws.Cells(i + 1, 4).Value = rec(IDX_TITLE)
        ws.Cells(i + 1, 5).Value = rec(IDX_BASIS)
        ws.Cells(i + 1, 6).Value = rec(IDX_BASE_LABEL)
        ws.Cells(i + 1, 7).Value = dataWs.Name
        ws.Cells(i + 1, 8).Value = dataWs.Range(dataWs.Cells(CLng(rec(IDX_START_ROW)), 1), dataWs.Cells(CLng(rec(IDX_END_ROW)), CLng(rec(IDX_LAST_COL)))).Address(False, False)
        ws.Cells(i + 1, 9).Value = rec(IDX_TABLE_TYPE)
        ws.Cells(i + 1, 10).Value = CLng(rec(IDX_END_ROW)) - CLng(rec(IDX_START_ROW)) + 1
        ws.Cells(i + 1, 11).Value = CLng(rec(IDX_LAST_COL))
        ws.Cells(i + 1, 12).Value = rec(IDX_WARNING)
    Next i

    ws.Columns("A:L").AutoFit
    ws.Rows(1).AutoFilter
End Sub

' ============================================================
' 프로시저 : ReportAutomation_WriteNarratives
' 설  명  : 분석문, 차트 데이터, 삽입용 표, QA 시트를 동시에 작성한다.
' 이유    : 네 산출물은 같은 핵심 포인트(points)를 공유하므로 한 번의 순회로 생성한다.
' ============================================================
Private Sub ReportAutomation_WriteNarratives(ByVal wsNarr As Worksheet, ByVal wsChart As Worksheet, ByVal wsInsert As Worksheet, _
                                             ByVal wsQA As Worksheet, ByVal tables As Collection, ByVal dataWs As Worksheet, _
                                             ByVal wsSettings As Worksheet)
    wsNarr.Range("A1:I1").Value = Array("table_key", "문항/표 제목", "분석문_기본", "분석문_세부", "분석문_LLM", _
                                        "주요 수치 요약", "검토 상태", "사용자 수정문", "최종 사용문")
    wsChart.Range("A1:H1").Value = Array("table_key", "series_group", "category", "measure", "value", "display_value", "sort_order", "include_chart")
    wsInsert.Range("A1:H1").Value = Array("table_key", "title", "category", "weighted_n", "raw_n", "percent", "unit", "source_cell")
    wsQA.Range("A1:F1").Value = Array("table_key", "qa_type", "severity", "message", "source_range", "checked_at")

    ReportAutomation_StyleHeader wsNarr.Range("A1:I1")
    ReportAutomation_StyleHeader wsChart.Range("A1:H1")
    ReportAutomation_StyleHeader wsInsert.Range("A1:H1")
    ReportAutomation_StyleHeader wsQA.Range("A1:F1")

    ' 사용자가 설정 시트에서 선택한 배너 목록과 제목 제거 접두어를 읽는다.
    Dim bannerSetting As String
    bannerSetting = ReportAutomation_ReadBannerSetting(wsSettings)

    Dim titlePrefixes() As String
    titlePrefixes = ReportAutomation_ReadTitlePrefixes(wsSettings)

    Dim narrRow As Long: narrRow = 2
    Dim chartRow As Long: chartRow = 2
    Dim insertRow As Long: insertRow = 2
    Dim qaRow As Long: qaRow = 2

    Dim i As Long, rec As Variant
    For i = 1 To tables.Count
        rec = tables(i)

        ' points는 분석문/차트/삽입표에 공통으로 쓰는 핵심 수치 목록이다.
        ' 각 원소는 Array(category, percent, weighted_n, raw_n, source_cell) 구조를 따른다.
        Dim points As Collection
        Set points = ReportAutomation_ExtractKeyPoints(dataWs, rec)

        ' 보고서 본문용 문장과 리뷰용 수치 요약을 각각 생성한다.
        Dim narrative As String, summary As String
        narrative = ReportAutomation_BuildNarrative(CStr(rec(IDX_TITLE)), CStr(rec(IDX_BASIS)), points, titlePrefixes)
        summary = ReportAutomation_FormatPointSummary(points)

        wsNarr.Cells(narrRow, 1).Value = rec(IDX_TABLE_KEY)
        wsNarr.Cells(narrRow, 2).Value = rec(IDX_TITLE)
        wsNarr.Cells(narrRow, 3).Value = narrative
        wsNarr.Cells(narrRow, 6).Value = summary
        wsNarr.Cells(narrRow, 7).Value = IIf(points.Count = 0, "확인 필요", "자동 생성")
        wsNarr.Cells(narrRow, 9).Value = narrative
        narrRow = narrRow + 1

        ' 전체 기준 수치를 차트 데이터와 삽입용 표 데이터로 펼친다.
        Dim p As Long, pointRec As Variant
        For p = 1 To points.Count
            pointRec = points(p)

            wsChart.Cells(chartRow, 1).Value = rec(IDX_TABLE_KEY)
            wsChart.Cells(chartRow, 2).Value = "전체"
            wsChart.Cells(chartRow, 3).Value = pointRec(0)
            wsChart.Cells(chartRow, 4).Value = "%"
            wsChart.Cells(chartRow, 5).Value = pointRec(1)
            wsChart.Cells(chartRow, 6).Value = ReportAutomation_FormatPercent(CDbl(pointRec(1)))
            wsChart.Cells(chartRow, 7).Value = p
            wsChart.Cells(chartRow, 8).Value = IIf(p <= CHART_INCLUDE_TOP_N, "Y", "N")
            chartRow = chartRow + 1

            wsInsert.Cells(insertRow, 1).Value = rec(IDX_TABLE_KEY)
            wsInsert.Cells(insertRow, 2).Value = rec(IDX_TITLE)
            wsInsert.Cells(insertRow, 3).Value = pointRec(0)
            wsInsert.Cells(insertRow, 4).Value = pointRec(2)
            wsInsert.Cells(insertRow, 5).Value = pointRec(3)
            wsInsert.Cells(insertRow, 6).Value = pointRec(1)
            wsInsert.Cells(insertRow, 7).Value = "%"
            wsInsert.Cells(insertRow, 8).Value = pointRec(4)
            insertRow = insertRow + 1
        Next p

        ' 사용자가 선택한 배너가 있고 배너 교차표인 경우에만 배너별 수치를 차트 데이터에 추가한다.
        If CLng(rec(IDX_LAST_COL)) > SIMPLE_TABLE_MAX_COL Then
            Dim bannerGroups As Collection
            Set bannerGroups = ReportAutomation_FindBannerGroups(dataWs, CLng(rec(IDX_START_ROW)), _
                                                                  ReportAutomation_FindTotalRow(dataWs, CLng(rec(IDX_START_ROW)), CLng(rec(IDX_END_ROW))), _
                                                                  CLng(rec(IDX_LAST_COL)))
            ReportAutomation_AppendBannerChartRows wsChart, chartRow, dataWs, rec, bannerGroups, bannerSetting
        End If

        ' 분석문 생성 실패 또는 구조 경고가 있는 표는 QA 시트에 남겨 수동 검토 대상으로 만든다.
        If points.Count = 0 Or Len(CStr(rec(IDX_WARNING))) > 0 Then
            wsQA.Cells(qaRow, 1).Value = rec(IDX_TABLE_KEY)
            wsQA.Cells(qaRow, 2).Value = IIf(points.Count = 0, "수치 QA", "표 구조 QA")
            wsQA.Cells(qaRow, 3).Value = "warning"
            wsQA.Cells(qaRow, 4).Value = IIf(points.Count = 0, "분석문 생성에 사용할 전체 기준 비율 값을 찾지 못했습니다.", rec(IDX_WARNING))
            wsQA.Cells(qaRow, 5).Value = dataWs.Range(dataWs.Cells(CLng(rec(IDX_START_ROW)), 1), dataWs.Cells(CLng(rec(IDX_END_ROW)), CLng(rec(IDX_LAST_COL)))).Address(False, False)
            wsQA.Cells(qaRow, 6).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
            qaRow = qaRow + 1
        End If
    Next i

    wsNarr.Columns("A:I").AutoFit
    wsNarr.Columns("C:I").ColumnWidth = 36
    wsNarr.Columns("C:I").WrapText = True
    wsChart.Columns("A:H").AutoFit
    wsInsert.Columns("A:H").AutoFit
    wsQA.Columns("A:F").AutoFit
    wsNarr.Rows(1).AutoFilter
    wsChart.Rows(1).AutoFilter
    wsInsert.Rows(1).AutoFilter
    wsQA.Rows(1).AutoFilter
End Sub

' ============================================================
' 함수명 : ReportAutomation_ExtractKeyPoints
' 설  명 : 표 유형에 맞춰 핵심 응답 수치 목록을 추출하고 내림차순 정렬한다.
' 반환값 : Collection(Array(category, percent, weighted_n, raw_n, source_cell))
' ============================================================
Private Function ReportAutomation_ExtractKeyPoints(ByVal ws As Worksheet, ByVal tableRec As Variant) As Collection
    Dim points As New Collection
    Dim startRow As Long, endRow As Long, lastCol As Long
    startRow = CLng(tableRec(IDX_START_ROW))
    endRow = CLng(tableRec(IDX_END_ROW))
    lastCol = CLng(tableRec(IDX_LAST_COL))

    ' 열 수가 적은 표는 일반 분포표, 넓은 표는 배너 교차표로 보고 다른 파서를 적용한다.
    If lastCol <= SIMPLE_TABLE_MAX_COL Then
        ReportAutomation_ExtractSimplePoints ws, startRow, endRow, lastCol, points
    Else
        ReportAutomation_ExtractWidePoints ws, startRow, endRow, lastCol, points
    End If

    ReportAutomation_SortPointsDescending points
    Set ReportAutomation_ExtractKeyPoints = points
End Function

' ============================================================
' 프로시저 : ReportAutomation_ExtractSimplePoints
' 설  명  : 단일 전체 분포표에서 항목별 비율을 추출한다.
' 전  제  : 항목명은 보통 A/B열, 비율은 % 헤더 열 또는 기본 5열에 있다.
' ============================================================
Private Sub ReportAutomation_ExtractSimplePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection)
    Dim pctCol As Long
    pctCol = ReportAutomation_FindHeaderColumn(ws, startRow, endRow, lastCol, "%")
    If pctCol = 0 Then pctCol = DEFAULT_PCT_COL

    Dim r As Long
    For r = startRow + 3 To endRow
        Dim category As String
        ' SPSS 표 구조에 따라 항목명이 B열 또는 A열에 올 수 있어 B열 우선, A열 보조로 읽는다.
        category = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(category) = 0 Then category = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(category) > 0 And Left$(category, 1) <> "■" And Left$(category, 1) <> "▣" Then
            If IsNumeric(ws.Cells(r, pctCol).Value) Then
                points.Add Array(category, CDbl(ws.Cells(r, pctCol).Value), ws.Cells(r, 3).Value, ws.Cells(r, 4).Value, ws.Cells(r, pctCol).Address(False, False))
            End If
        End If
    Next r
End Sub

' ============================================================
' 프로시저 : ReportAutomation_ExtractWidePoints
' 설  명  : 배너 교차표에서 전체 행(■로 시작)의 비율 값을 추출한다.
' 전  제  : 전체 행의 각 % 열이 보고서 전체 요약의 후보 수치가 된다.
' ============================================================
Private Sub ReportAutomation_ExtractWidePoints(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal points As Collection)
    Dim totalRow As Long
    totalRow = ReportAutomation_FindTotalRow(ws, startRow, endRow)
    If totalRow = 0 Then Exit Sub

    ' 전체 행의 N/% 쌍 중 % 열만 읽어 응답 범주별 비율 후보로 만든다.
    Dim c As Long
    For c = 4 To lastCol
        If ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) And IsNumeric(ws.Cells(totalRow, c).Value) Then
            Dim category As String
            category = ReportAutomation_HeaderLabel(ws, startRow + 1, totalRow - 1, c)
            If Len(category) > 0 And category <> "계" Then
                points.Add Array(category, CDbl(ws.Cells(totalRow, c).Value), "", ws.Cells(totalRow, 3).Value, ws.Cells(totalRow, c).Address(False, False))
            End If
        End If
    Next c
End Sub

' ============================================================
' 함수명 : ReportAutomation_FindHeaderColumn
' 설  명 : 표 상단 일부 행에서 특정 헤더 텍스트가 있는 열을 찾는다.
' 반환값 : 열 번호, 찾지 못하면 0
' ============================================================
Private Function ReportAutomation_FindHeaderColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal lastCol As Long, ByVal headerText As String) As Long
    Dim r As Long, c As Long
    For r = startRow To Application.WorksheetFunction.Min(startRow + 5, endRow)
        For c = 1 To lastCol
            If Trim$(CStr(ws.Cells(r, c).Value)) = headerText Then
                ReportAutomation_FindHeaderColumn = c
                Exit Function
            End If
        Next c
    Next r
End Function

' ============================================================
' 함수명 : ReportAutomation_FindTotalRow
' 설  명 : 배너 교차표에서 전체/합계 역할을 하는 행을 찾는다.
' 기준   : A열 값이 "■"로 시작하는 행
' ============================================================
Private Function ReportAutomation_FindTotalRow(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    For r = startRow To endRow
        If Left$(Trim$(CStr(ws.Cells(r, 1).Value)), 1) = "■" Then
            ReportAutomation_FindTotalRow = r
            Exit Function
        End If
    Next r
End Function

' ============================================================
' 함수명 : ReportAutomation_IsPercentMeasureColumn
' 설  명 : 특정 열이 N/% 쌍 중 비율(%) 열인지 판정한다.
' 기준   : 전체 행 위쪽 헤더 영역에 "%" 표시가 있는지 확인
' ============================================================
Private Function ReportAutomation_IsPercentMeasureColumn(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal colIndex As Long) As Boolean
    Dim r As Long
    For r = startRow + 1 To totalRow - 1
        If Trim$(CStr(ws.Cells(r, colIndex).Value)) = "%" Then
            ReportAutomation_IsPercentMeasureColumn = True
            Exit Function
        End If
    Next r
End Function

' ============================================================
' 함수명 : ReportAutomation_HeaderLabel
' 설  명 : 병합/다단 헤더 구조에서 특정 열의 응답 범주명을 추정한다.
' 방식   : 현재 열에서 왼쪽으로 거슬러 올라가며 의미 있는 텍스트를 찾는다.
' ============================================================
Private Function ReportAutomation_HeaderLabel(ByVal ws As Worksheet, ByVal firstHeaderRow As Long, ByVal lastHeaderRow As Long, ByVal colIndex As Long) As String
    Dim r As Long, c As Long, text As String
    For r = lastHeaderRow To firstHeaderRow Step -1
        For c = colIndex To 1 Step -1
            text = Trim$(CStr(ws.Cells(r, c).Value))
            If Len(text) > 0 And text <> "%" And UCase$(text) <> "N" And text <> "사례수" Then
                ReportAutomation_HeaderLabel = text
                Exit Function
            End If
        Next c
    Next r
End Function

' ============================================================
' 프로시저 : ReportAutomation_SortPointsDescending
' 설  명  : 핵심 수치 목록을 percent 값 기준 내림차순으로 정렬한다.
' 주  의  : VBA Collection은 직접 정렬이 어려워 배열로 복사 후 다시 채운다.
' ============================================================
Private Sub ReportAutomation_SortPointsDescending(ByVal points As Collection)
    If points.Count <= 1 Then Exit Sub

    Dim arr() As Variant
    ReDim arr(1 To points.Count)

    Dim i As Long, j As Long
    For i = 1 To points.Count
        arr(i) = points(i)
    Next i

    ' 데이터 수가 표별로 크지 않으므로 단순 비교 정렬로 충분하다.
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CDbl(arr(j)(1)) > CDbl(arr(i)(1)) Then
                Dim temp As Variant
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    Do While points.Count > 0
        points.Remove 1
    Loop
    For i = LBound(arr) To UBound(arr)
        points.Add arr(i)
    Next i
End Sub

' ============================================================
' 함수명 : ReportAutomation_BuildNarrative
' 설  명 : 핵심 수치 목록을 보고서 본문 문장으로 변환한다.
' 문체   : 첨부 Python 작성기와 맞춰 "가장 높게 나타남 / 다음으로 ... 순" 구조 사용
' 반환값 : 분석문_기본 및 최종 사용문에 들어갈 텍스트
' ============================================================
Private Function ReportAutomation_BuildNarrative(ByVal titleText As String, ByVal basisText As String, ByVal points As Collection, ByRef titlePrefixes() As String) As String
    Dim normalizedTitle As String
    normalizedTitle = ReportAutomation_NormalizeAnalysisTitle(titleText, titlePrefixes)
    If Len(normalizedTitle) = 0 Then normalizedTitle = titleText

    If points.Count = 0 Then
        ReportAutomation_BuildNarrative = normalizedTitle & " 문항의 수치형 응답을 찾지 못했습니다."
        Exit Function
    End If

    Dim p1 As Variant
    p1 = points(1)

    Dim text As String
    ' 복수응답 문항은 해석 기준을 명시해 단일응답 문항과 구분한다.
    If ReportAutomation_IsMultiResponse(titleText, basisText) Then
        text = normalizedTitle & ReportAutomation_TopicParticle(normalizedTitle) & " 복수응답 기준으로 " & _
               ReportAutomation_Quoted(CStr(p1(0))) & ReportAutomation_SubjectParticle(CStr(p1(0))) & " " & _
               ReportAutomation_FormatPercent(CDbl(p1(1))) & "로 가장 높게 나타남"
    Else
        text = normalizedTitle & ReportAutomation_TopicParticle(normalizedTitle) & " " & _
               ReportAutomation_Quoted(CStr(p1(0))) & ReportAutomation_SubjectParticle(CStr(p1(0))) & " " & _
               ReportAutomation_FormatPercent(CDbl(p1(1))) & "로 가장 높게 나타남"
    End If

    ' 2~4위 항목은 별도 줄로 구성해 본문 편집 시 첫 문장과 보조 문장을 나누기 쉽게 한다.
    Dim othersText As String, i As Long, maxItems As Long, pointRec As Variant
    maxItems = points.Count
    If maxItems > NARRATIVE_TOP_ITEMS Then maxItems = NARRATIVE_TOP_ITEMS
    For i = 2 To maxItems
        pointRec = points(i)
        If Len(othersText) > 0 Then othersText = othersText & ", "
        othersText = othersText & ReportAutomation_Quoted(CStr(pointRec(0))) & _
                     "(" & ReportAutomation_FormatPercent(CDbl(pointRec(1))) & ")"
    Next i
    If Len(othersText) > 0 Then
        text = text & vbLf & "다음으로 " & othersText & " 순으로 나타남"
    End If

    ReportAutomation_BuildNarrative = text
End Function

' ============================================================
' 함수명 : ReportAutomation_FormatPointSummary
' 설  명 : 리뷰용으로 핵심 수치와 원본 셀 주소를 압축 표기한다.
' 예  시 : 도소매업=43.1%(E12)
' ============================================================
Private Function ReportAutomation_FormatPointSummary(ByVal points As Collection) As String
    Dim i As Long, maxItems As Long, text As String, rec As Variant
    maxItems = points.Count
    If maxItems > SUMMARY_TOP_ITEMS Then maxItems = SUMMARY_TOP_ITEMS

    For i = 1 To maxItems
        rec = points(i)
        If Len(text) > 0 Then text = text & "; "
        text = text & CStr(rec(0)) & "=" & ReportAutomation_FormatPercent(CDbl(rec(1))) & "(" & CStr(rec(4)) & ")"
    Next i
    ReportAutomation_FormatPointSummary = text
End Function

' ============================================================
' 함수명 : ReportAutomation_NormalizeAnalysisTitle
' 설  명 : 표 제목에서 보고서 문장에 불필요한 기준/메타 문구를 제거한다.
' 예  시 : "[리스트 기준] 대분류..." → "대분류..."
' ============================================================
Private Function ReportAutomation_NormalizeAnalysisTitle(ByVal titleText As String, ByRef userPrefixes() As String) As String
    Dim text As String
    Dim closePos As Long
    Dim pIdx As Long
    Dim pfx As String
    Dim marker As Variant
    Dim pos As Long

    text = ReportAutomation_CleanText(titleText)

    If Left$(text, 1) = "[" Then
        closePos = InStr(1, text, "]", vbTextCompare)
        ' 제목 앞의 [리스트 기준], [응답 기준] 등은 본문 문장에서는 제거한다.
        If closePos > 0 And closePos < Len(text) Then text = ReportAutomation_CleanText(Mid$(text, closePos + 1))
    End If

    ' 설정 시트 "제목 제거 접두어"에서 읽어온 목록을 순회해 제거한다.
    ' 사용자가 설정 시트 B열에 쉼표 구분으로 접두어를 추가하면 이 배열에 반영된다.
    For pIdx = LBound(userPrefixes) To UBound(userPrefixes)
        pfx = Trim$(userPrefixes(pIdx))
        If Len(pfx) > 0 Then
            If Left$(text, Len(pfx)) = pfx Then
                text = ReportAutomation_CleanText(Mid$(text, Len(pfx) + 1))
                Exit For
            End If
            ' 뒤에 공백이 붙지 않은 형태도 처리한다 (예: "성별" vs "성별 ").
            If Left$(text, Len(pfx & " ")) = pfx & " " Then
                text = ReportAutomation_CleanText(Mid$(text, Len(pfx) + 2))
                Exit For
            End If
        End If
    Next pIdx

    ' 표 머리글이 제목에 붙어 들어온 경우, 분석 제목 뒤쪽의 메타 컬럼명을 잘라낸다.
    For Each marker In Array("사례수", "전체", "진흥지구명", "사업체 특성", "단위", "Base :")
        pos = InStr(1, text, CStr(marker), vbTextCompare)
        If pos > 1 Then
            text = ReportAutomation_CleanText(Left$(text, pos - 1))
            Exit For
        End If
    Next marker

    ReportAutomation_NormalizeAnalysisTitle = text
End Function

' ============================================================
' 함수명 : ReportAutomation_CleanText
' 설  명 : 줄바꿈/비분리 공백/중복 공백을 정리한 단일 줄 텍스트를 반환한다.
' ============================================================
Private Function ReportAutomation_CleanText(ByVal text As String) As String
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    Do While InStr(1, text, "  ", vbBinaryCompare) > 0
        text = Replace(text, "  ", " ")
    Loop
    ReportAutomation_CleanText = Trim$(text)
End Function

' ============================================================
' 함수명 : ReportAutomation_IsMultiResponse
' 설  명 : 제목/기준 문구에 복수응답을 시사하는 표현이 있는지 판정한다.
' ============================================================
Private Function ReportAutomation_IsMultiResponse(ByVal titleText As String, ByVal basisText As String) As Boolean
    Dim text As String
    text = titleText & " " & basisText
    ReportAutomation_IsMultiResponse = (InStr(1, text, "복수", vbTextCompare) > 0 Or _
                                        InStr(1, text, "중복", vbTextCompare) > 0 Or _
                                        InStr(1, text, "multiple", vbTextCompare) > 0)
End Function

' ============================================================
' 함수명 : ReportAutomation_FormatPercent
' 설  명 : 비율 값을 보고서 기본 표기인 소수점 1자리 퍼센트로 변환한다.
' ============================================================
Private Function ReportAutomation_FormatPercent(ByVal valueData As Double) As String
    ReportAutomation_FormatPercent = Format(valueData, "0.0") & "%"
End Function

' ============================================================
' 함수명 : ReportAutomation_Quoted
' 설  명 : 응답 항목명을 보고서 문장용 작은따옴표 표기로 감싼다.
' ============================================================
Private Function ReportAutomation_Quoted(ByVal text As String) As String
    ReportAutomation_Quoted = "'" & text & "'"
End Function

' ============================================================
' 함수명 : ReportAutomation_TopicParticle
' 설  명 : 앞 단어 받침 여부에 따라 보조사 "은/는"을 반환한다.
' ============================================================
Private Function ReportAutomation_TopicParticle(ByVal text As String) As String
    ReportAutomation_TopicParticle = IIf(ReportAutomation_HasBatchim(text), "은", "는")
End Function

' ============================================================
' 함수명 : ReportAutomation_SubjectParticle
' 설  명 : 앞 단어 받침 여부에 따라 주격 조사 "이/가"를 반환한다.
' ============================================================
Private Function ReportAutomation_SubjectParticle(ByVal text As String) As String
    ReportAutomation_SubjectParticle = IIf(ReportAutomation_HasBatchim(text), "이", "가")
End Function

' ============================================================
' 함수명 : ReportAutomation_HasBatchim
' 설  명 : 문자열 마지막 의미 문자에 한글 받침이 있는지 판정한다.
' 보  조 : 숫자/영문 끝 글자도 자연스러운 조사 선택을 위해 간단히 처리한다.
' ============================================================
Private Function ReportAutomation_HasBatchim(ByVal text As String) As Boolean
    Dim cleaned As String
    cleaned = ReportAutomation_CleanText(text)
    Do While Len(cleaned) > 0 And InStr(1, "'""”)]} ", Right$(cleaned, 1), vbBinaryCompare) > 0
        cleaned = Left$(cleaned, Len(cleaned) - 1)
    Loop
    If Len(cleaned) = 0 Then Exit Function

    Dim ch As String
    ch = Right$(cleaned, 1)

    ' 숫자는 한국어로 읽었을 때 받침이 있는 숫자만 받침 있음으로 본다.
    If ch Like "#" Then
        ReportAutomation_HasBatchim = (InStr(1, "013678", ch, vbBinaryCompare) > 0)
        Exit Function
    End If

    ' AscW가 음수로 반환되는 환경을 고려해 0~65535 범위로 보정한다.
    Dim code As Long
    code = AscW(ch)
    If code < 0 Then code = code + 65536

    ' 한글 완성형 음절은 (코드-가) Mod 28 값으로 받침 유무를 판정한다.
    If code >= 44032 And code <= 55203 Then
        ReportAutomation_HasBatchim = (((code - 44032) Mod 28) <> 0)
    Else
        ch = LCase$(ch)
        ReportAutomation_HasBatchim = (InStr(1, "bcklmnpst", ch, vbBinaryCompare) > 0)
    End If
End Function

' ============================================================
' 프로시저 : ReportAutomation_WriteSourceSheet
' 설  명  : 보고서에 사용할 출처 관리용 기본 시트를 생성한다.
' 현재값  : 원본 집계표를 primary Excel source로 1건 등록한다.
' ============================================================
Private Sub ReportAutomation_WriteSourceSheet(ByVal ws As Worksheet)
    ws.Range("A1:J1").Value = Array("source_id", "title", "author_org", "year", "url", "accessed_at", "source_type", "reliability", "used_section", "note")
    ReportAutomation_StyleHeader ws.Range("A1:J1")
    ws.Cells(2, 1).Value = "SRC_0001"
    ws.Cells(2, 2).Value = "원본 집계표"
    ws.Cells(2, 7).Value = "Excel"
    ws.Cells(2, 8).Value = "primary"
    ws.Columns("A:J").AutoFit
End Sub

' ============================================================
' 프로시저 : ReportAutomation_WriteRevisionSheet
' 설  명  : 사용자의 문장 수정 이력을 기록할 빈 시트 구조를 생성한다.
' 용  도  : 후속 LLM 재작성/검수 단계에서 before/after를 추적한다.
' ============================================================
Private Sub ReportAutomation_WriteRevisionSheet(ByVal ws As Worksheet)
    ws.Range("A1:G1").Value = Array("revision_id", "target", "user_feedback", "before", "after", "created_at", "applied_status")
    ReportAutomation_StyleHeader ws.Range("A1:G1")
    ws.Columns("A:G").AutoFit
End Sub

' ============================================================
' 프로시저 : ReportAutomation_WriteMetaSheet
' 설  명  : 후속 자동화가 참조할 실행 메타데이터를 key/value 형태로 기록한다.
' 주  의  : 생성 후 xlSheetVeryHidden으로 숨긴다.
' ============================================================
Private Sub ReportAutomation_WriteMetaSheet(ByVal ws As Worksheet, ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal tables As Collection, _
                                            ByVal wsSettings As Worksheet, ByVal wsList As Worksheet, ByVal wsNarr As Worksheet, _
                                            ByVal wsChart As Worksheet, ByVal wsInsert As Worksheet)
    ws.Range("A1:B1").Value = Array("key", "value")
    ReportAutomation_StyleHeader ws.Range("A1:B1")

    Dim rowNo As Long: rowNo = 2
    ReportAutomation_WriteMetaKV ws, rowNo, "generated_at", Format(Now, "yyyy-mm-dd hh:nn:ss")
    ReportAutomation_WriteMetaKV ws, rowNo, "source_workbook", wb.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "source_sheet", dataWs.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "table_count", tables.Count
    ' 산출 시트명은 timestamp가 붙으므로 메타에 정확한 이름을 남겨 후속 스크립트가 찾을 수 있게 한다.
    ReportAutomation_WriteMetaKV ws, rowNo, "settings_sheet", wsSettings.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "table_list_sheet", wsList.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "narrative_sheet", wsNarr.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "chart_data_sheet", wsChart.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "insert_table_sheet", wsInsert.Name
    ReportAutomation_WriteMetaKV ws, rowNo, "style_profile", "formal_korean_research"
    ReportAutomation_WriteMetaKV ws, rowNo, "number_format", "percent_1_decimal"
    ReportAutomation_WriteMetaKV ws, rowNo, "hwp_chart_paste_mode", "metafile_preferred"
    ws.Columns("A:B").AutoFit
End Sub

' ============================================================
' 프로시저 : ReportAutomation_WriteMetaKV
' 설  명  : 메타 시트에 key/value 한 줄을 쓰고 다음 행 번호로 이동한다.
' ============================================================
Private Sub ReportAutomation_WriteMetaKV(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal keyText As String, ByVal valueText As Variant)
    ws.Cells(rowNo, 1).Value = keyText
    ws.Cells(rowNo, 2).Value = valueText
    rowNo = rowNo + 1
End Sub

' ============================================================
' 프로시저 : ReportAutomation_StyleHeader
' 설  명  : 산출 시트 헤더 행에 공통 서식을 적용한다.
' 범  위  : 굵게, 연한 파랑 배경, 기본 테두리
' ============================================================
Private Sub ReportAutomation_StyleHeader(ByVal targetRange As Range)
    With targetRange
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
    End With
End Sub

' ============================================================
' 함수명 : ReportAutomation_UniqueSheetName
' 설  명 : Excel의 31자 제한을 지키면서 중복 없는 산출 시트명을 만든다.
' 방식   : baseName_mmdd_hhnn 형식, 중복 시 뒤에 _2, _3 추가
' ============================================================
Private Function ReportAutomation_UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim name As String
    name = Left$(baseName & "_" & Format(Now, "mmdd_hhnn"), 31)
    Dim candidate As String
    candidate = name

    Dim n As Long
    n = 1
    Do While ReportAutomation_WorksheetExists(wb, candidate)
        n = n + 1
        candidate = Left$(name, 28) & "_" & n
    Loop
    ReportAutomation_UniqueSheetName = candidate
End Function

' ============================================================
' 함수명 : ReportAutomation_FindLatestOutputSheet
' 설  명 : 특정 접두어로 시작하는 산출 시트 중 통합문서에서 가장 뒤에 있는 시트를 찾는다.
' 용  도 : 이전 실행의 설정 시트를 읽어 사용자가 편집한 옵션을 다음 실행에 승계한다.
' ============================================================
Private Function ReportAutomation_FindLatestOutputSheet(ByVal wb As Workbook, ByVal baseName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Left$(ws.Name, Len(baseName)) = baseName Then
            Set ReportAutomation_FindLatestOutputSheet = ws
        End If
    Next ws
End Function

' ============================================================
' 함수명 : ReportAutomation_SettingValue
' 설  명 : 이전 설정 시트에서 동일 라벨 값을 찾아 반환하고, 없으면 기본값을 반환한다.
' 주  의 : 원본 통합문서/시트명처럼 실행마다 달라지는 값은 호출부에서 별도로 고정한다.
' ============================================================
Private Function ReportAutomation_SettingValue(ByVal labelText As String, ByVal defaultValue As Variant, ByVal priorSettingsWs As Worksheet) As Variant
    If priorSettingsWs Is Nothing Then
        ReportAutomation_SettingValue = defaultValue
        Exit Function
    End If

    Dim r As Long
    For r = 3 To priorSettingsWs.Cells(priorSettingsWs.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(priorSettingsWs.Cells(r, 1).Value)) = labelText Then
            ReportAutomation_SettingValue = priorSettingsWs.Cells(r, 2).Value
            Exit Function
        End If
    Next r

    ReportAutomation_SettingValue = defaultValue
End Function

' ============================================================
' 프로시저 : ReportAutomation_SetSettingValue
' 설  명  : 설정 시트에서 라벨을 찾아 값을 갱신한다.
' 용  도  : UserForm에서 입력받은 값을 생성 직후 설정 시트에 반영한다.
' ============================================================
Private Sub ReportAutomation_SetSettingValue(ByVal ws As Worksheet, ByVal labelText As String, ByVal valueText As String)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 1).Value)) = labelText Then
            ws.Cells(r, 2).Value = valueText
            Exit Sub
        End If
    Next r
End Sub

' ============================================================
' 프로시저 : ReportAutomation_ClearOptionOverrides
' 설  명  : UserForm에서 넘긴 1회성 override 값을 초기화한다.
' ============================================================
Private Sub ReportAutomation_ClearOptionOverrides()
    mHasOptionOverrides = False
    mOverrideBannerSetting = ""
    mOverrideTitlePrefixes = ""
End Sub

' ============================================================
' 함수명 : ReportAutomation_CellText
' 설  명 : 셀의 표시 텍스트를 읽되, 병합셀인 경우 병합 영역의 좌상단 값을 반환한다.
' 용  도 : 배너 교차표의 병합 헤더를 모든 소속 열에서 동일하게 인식한다.
' ============================================================
Private Function ReportAutomation_CellText(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    Dim cell As Range
    Set cell = ws.Cells(rowIndex, colIndex)

    If cell.MergeCells Then
        ReportAutomation_CellText = Trim$(CStr(cell.MergeArea.Cells(1, 1).Value))
    Else
        ReportAutomation_CellText = Trim$(CStr(cell.Value))
    End If
End Function

' ============================================================
' 함수명 : ReportAutomation_WorksheetExists
' 설  명 : 통합문서에 특정 이름의 워크시트가 있는지 확인한다.
' ============================================================
Private Function ReportAutomation_WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    ReportAutomation_WorksheetExists = Not ws Is Nothing
End Function

' ============================================================
' 프로시저 : ReportAutomation_BeginOperation
' 설  명  : 긴 작업 전 Excel 상태를 저장하고 화면갱신/이벤트/계산을 잠시 끈다.
' 효과    : 대형 집계표 처리 속도 개선 및 중간 이벤트 충돌 방지
' ============================================================
Private Sub ReportAutomation_BeginOperation(ByVal statusText As String)
    On Error Resume Next
    mPrevScreenUpdating = Application.ScreenUpdating
    mPrevEnableEvents = Application.EnableEvents
    mPrevDisplayAlerts = Application.DisplayAlerts
    mPrevCalculation = Application.Calculation
    mPrevStatusBar = Application.StatusBar

    ' 수천 행 집계표에서 시트 생성/서식 적용을 반복하므로 성능 옵션을 임시로 낮춘다.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = statusText
    On Error GoTo 0
End Sub

' ============================================================
' 프로시저 : ReportAutomation_EndOperation
' 설  명  : BeginOperation에서 바꾼 Excel Application 상태를 원래대로 복원한다.
' 주  의  : 오류 처리 경로에서도 호출되므로 On Error Resume Next로 최대한 복구한다.
' ============================================================
Private Sub ReportAutomation_EndOperation()
    On Error Resume Next
    Application.ScreenUpdating = mPrevScreenUpdating
    Application.EnableEvents = mPrevEnableEvents
    Application.DisplayAlerts = mPrevDisplayAlerts
    Application.Calculation = mPrevCalculation
    Application.StatusBar = mPrevStatusBar
    On Error GoTo 0
End Sub

' ============================================================
' 프로시저 : ReportAutomation_LogEvent
' 설  명  : 사용자 통합문서 내부 로그 시트에 실행 결과를 남긴다.
' 주  의  : 로그 시트는 VeryHidden 처리해 사용자의 보고서 시트 목록을 방해하지 않는다.
' ============================================================
Private Sub ReportAutomation_LogEvent(ByVal wb As Workbook, ByVal actionName As String, ByVal targetName As String, ByVal statusText As String, ByVal detailText As String)
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(REPORT_LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        ' 최초 실행 시에만 로그 시트를 만들고 헤더를 구성한다.
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = REPORT_LOG_SHEET
        ws.Range("A1:F1").Value = Array("logged_at", "version", "action", "target", "status", "detail")
        ReportAutomation_StyleHeader ws.Range("A1:F1")
        ws.Visible = xlSheetVeryHidden
    End If

    ' 로그는 append-only로 누적해 반복 실행/오류 이력을 보존한다.
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = REPORT_AUTOMATION_VERSION
    ws.Cells(nextRow, 3).Value = actionName
    ws.Cells(nextRow, 4).Value = targetName
    ws.Cells(nextRow, 5).Value = statusText
    ws.Cells(nextRow, 6).Value = detailText
End Sub

' ============================================================
' 함수명 : ReportAutomation_ReadBannerSetting
' 설  명 : 설정 시트의 "추출 배너 목록" 항목을 읽는다.
' 반환값 : 쉼표 구분 배너 목록 문자열. 항목이 없으면 "전체" 반환.
' 예  시 : "전체" / "전체,성별" / "전체,성별,연령대"
' ============================================================
Private Function ReportAutomation_ReadBannerSetting(ByVal ws As Worksheet) As String
    If ws Is Nothing Then
        ReportAutomation_ReadBannerSetting = "전체"
        Exit Function
    End If

    Dim r As Long
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "추출 배너 목록" Then
            Dim val As String
            val = Trim$(CStr(ws.Cells(r, 2).Value))
            ReportAutomation_ReadBannerSetting = IIf(Len(val) > 0, val, "전체")
            Exit Function
        End If
    Next r
    ReportAutomation_ReadBannerSetting = "전체"
End Function

' ============================================================
' 함수명 : ReportAutomation_ReadTitlePrefixes
' 설  명 : 설정 시트의 "제목 제거 접두어" 항목을 읽어 배열로 반환한다.
' 반환값 : 쉼표로 분리된 접두어 문자열 배열. 항목 없으면 빈 배열.
' 사용법 : 설정 시트 B열에 쉼표 구분으로 입력 (예: 진흥지구별,사업체 특성별)
' ============================================================
Private Function ReportAutomation_ReadTitlePrefixes(ByVal ws As Worksheet) As String()
    Dim defaultArr() As String
    ReDim defaultArr(0 To 0)
    defaultArr(0) = ""

    If ws Is Nothing Then
        ReportAutomation_ReadTitlePrefixes = defaultArr
        Exit Function
    End If

    Dim r As Long
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "제목 제거 접두어" Then
            Dim val As String
            val = Trim$(CStr(ws.Cells(r, 2).Value))
            If Len(val) = 0 Then
                ReportAutomation_ReadTitlePrefixes = defaultArr
            Else
                ReportAutomation_ReadTitlePrefixes = Split(val, ",")
            End If
            Exit Function
        End If
    Next r
    ReportAutomation_ReadTitlePrefixes = defaultArr
End Function

' ============================================================
' 함수명 : ReportAutomation_FindBannerGroups
' 설  명 : 배너 교차표 헤더 영역을 스캔해 배너 그룹명과 소속 % 열을 매핑한다.
' 반환값 : Collection
'   각 원소: Array(bannerName As String, pctCols As Variant)
'     - bannerName : 배너 그룹명 (예: "전체", "성별", "연령대")
'     - pctCols    : 해당 그룹에 속한 % 열 번호 Long 배열
' 동작 원리:
'   헤더 행에서 N/% 가 아닌 텍스트를 배너 그룹명 후보로 보고,
'   그 열부터 다음 그룹 시작 전까지의 % 열들을 묶는다.
' ============================================================
Private Function ReportAutomation_FindBannerGroups(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalRow As Long, ByVal lastCol As Long) As Collection
    Dim groups As New Collection
    Dim headerStart As Long, headerEnd As Long
    Dim c As Long, r As Long
    Dim cellText As String
    Dim colGroupName() As String
    Dim curName As String
    Dim colList As Collection
    Dim lastSeenGroupName As String

    If totalRow = 0 Then
        Set ReportAutomation_FindBannerGroups = groups
        Exit Function
    End If

    ' 헤더 행 범위: 제목 다음 행 ~ 전체 행 위 행
    headerStart = startRow + 1
    headerEnd = totalRow - 1
    If headerStart > headerEnd Then
        Set ReportAutomation_FindBannerGroups = groups
        Exit Function
    End If

    ' 열별로 가장 위에 있는 배너 그룹명을 수집한다.
    ReDim colGroupName(1 To lastCol)
    For c = 1 To lastCol
        For r = headerStart To headerEnd
            cellText = ReportAutomation_CellText(ws, r, c)
            If Len(cellText) > 0 And cellText <> "%" And UCase$(cellText) <> "N" And cellText <> "사례수" Then
                colGroupName(c) = cellText
                lastSeenGroupName = cellText
                Exit For
            End If
        Next r
        If Len(colGroupName(c)) = 0 And ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) Then
            colGroupName(c) = lastSeenGroupName
        End If
    Next c

    ' 연속된 같은 그룹명 구간을 하나의 배너 그룹으로 묶는다.
    curName = ""
    Set colList = New Collection
    For c = 1 To lastCol
        If colGroupName(c) <> curName Then
            If Len(curName) > 0 And colList.Count > 0 Then
                groups.Add Array(curName, ReportAutomation_CollectionToArray(colList))
            End If
            curName = colGroupName(c)
            Set colList = New Collection
        End If
        If ReportAutomation_IsPercentMeasureColumn(ws, startRow, totalRow, c) Then
            colList.Add c
        End If
    Next c
    If Len(curName) > 0 And colList.Count > 0 Then
        groups.Add Array(curName, ReportAutomation_CollectionToArray(colList))
    End If

    Set ReportAutomation_FindBannerGroups = groups
End Function

' ============================================================
' 함수명 : ReportAutomation_CollectionToArray
' 설  명 : Long 값만 들어있는 Collection을 Variant 배열로 변환한다.
' ============================================================
Private Function ReportAutomation_CollectionToArray(ByVal col As Collection) As Variant
    If col.Count = 0 Then
        ReportAutomation_CollectionToArray = Array()
        Exit Function
    End If
    Dim arr() As Variant
    ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    ReportAutomation_CollectionToArray = arr
End Function

' ============================================================
' 프로시저 : ReportAutomation_SortArrayDescending
' 설  명  : Array 안의 Array 레코드를 특정 숫자 인덱스 기준으로 내림차순 정렬한다.
' 용  도  : 배너별 차트 데이터를 원본 행 순서가 아니라 값이 큰 순서로 기록한다.
' ============================================================
Private Sub ReportAutomation_SortArrayDescending(ByRef arr As Variant, ByVal valueIndex As Long)
    Dim i As Long, j As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CDbl(arr(j)(valueIndex)) > CDbl(arr(i)(valueIndex)) Then
                Dim temp As Variant
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' ============================================================
' 프로시저 : ReportAutomation_AppendBannerChartRows
' 설  명  : 사용자가 선택한 배너 그룹의 데이터를 차트 시트에 추가 기록한다.
' 인  자  : chartRow — 참조 전달(ByRef), 기록 후 다음 빈 행으로 이동한다.
'           bannerSetting — 쉼표 구분 배너 목록 문자열 (예: "전체,성별")
' 주  의  : "전체"는 WriteNarratives에서 이미 기록하므로 이 함수에서는 건너뛴다.
' ============================================================
Private Sub ReportAutomation_AppendBannerChartRows(ByVal wsChart As Worksheet, ByRef chartRow As Long, _
                                                    ByVal ws As Worksheet, ByVal tableRec As Variant, _
                                                    ByVal bannerGroups As Collection, ByVal bannerSetting As String)
    Dim startRow As Long, endRow As Long, totalRow As Long
    Dim bg As Variant
    Dim bgName As String
    Dim pctCols As Variant
    Dim colIdx As Long, r As Long
    Dim pctCol As Long
    Dim subLabel As String
    Dim sortOrder As Long
    Dim category As String

    startRow = CLng(tableRec(IDX_START_ROW))
    endRow   = CLng(tableRec(IDX_END_ROW))
    totalRow = ReportAutomation_FindTotalRow(ws, startRow, endRow)
    If totalRow = 0 Then Exit Sub

    For Each bg In bannerGroups
        bgName = CStr(bg(0))

        ' "전체"는 이미 기록됐으므로 건너뛴다.
        If bgName = "전체" Then GoTo NextGroup

        ' 사용자가 선택하지 않은 배너는 건너뛴다.
        If Not ReportAutomation_BannerRequested(bgName, bannerSetting) Then GoTo NextGroup

        pctCols = bg(1)
        If Not IsArray(pctCols) Then GoTo NextGroup

        ' 각 % 열마다 응답 항목 행을 순회해 기록한다.
        For colIdx = LBound(pctCols) To UBound(pctCols)
            pctCol = CLng(pctCols(colIdx))

            ' 이 % 열의 하위 범주 레이블을 헤더에서 읽는다.
            subLabel = ReportAutomation_HeaderLabel(ws, startRow + 1, totalRow - 1, pctCol)
            If Len(subLabel) = 0 Then subLabel = bgName

            ' 항목 행을 먼저 수집한 뒤 값 기준 내림차순으로 정렬한다.
            Dim rowCandidates As Collection
            Set rowCandidates = New Collection
            For r = startRow + 3 To endRow
                If r = totalRow Then GoTo NextRow

                category = Trim$(CStr(ws.Cells(r, 2).Value))
                If Len(category) = 0 Then category = Trim$(CStr(ws.Cells(r, 1).Value))
                If Len(category) = 0 Then GoTo NextRow
                If Left$(category, 1) = "■" Or Left$(category, 1) = "▣" Then GoTo NextRow

                If IsNumeric(ws.Cells(r, pctCol).Value) Then
                    rowCandidates.Add Array(category, CDbl(ws.Cells(r, pctCol).Value))
                End If
NextRow:
            Next r

            If rowCandidates.Count > 0 Then
                Dim rowsArr As Variant, rowIdx As Long, rowRec As Variant
                rowsArr = ReportAutomation_CollectionToArray(rowCandidates)
                ReportAutomation_SortArrayDescending rowsArr, 1

                For rowIdx = LBound(rowsArr) To UBound(rowsArr)
                    rowRec = rowsArr(rowIdx)
                    sortOrder = rowIdx - LBound(rowsArr) + 1
                    wsChart.Cells(chartRow, 1).Value = tableRec(IDX_TABLE_KEY)
                    wsChart.Cells(chartRow, 2).Value = subLabel
                    wsChart.Cells(chartRow, 3).Value = rowRec(0)
                    wsChart.Cells(chartRow, 4).Value = "%"
                    wsChart.Cells(chartRow, 5).Value = CDbl(rowRec(1))
                    wsChart.Cells(chartRow, 6).Value = ReportAutomation_FormatPercent(CDbl(rowRec(1)))
                    wsChart.Cells(chartRow, 7).Value = sortOrder
                    wsChart.Cells(chartRow, 8).Value = IIf(sortOrder <= CHART_INCLUDE_TOP_N, "Y", "N")
                    chartRow = chartRow + 1
                Next rowIdx
            End If
        Next colIdx
NextGroup:
    Next bg
End Sub

' ============================================================
' 함수명 : ReportAutomation_BannerRequested
' 설  명 : 배너 설정 문자열에 특정 배너명이 포함되어 있는지 확인한다.
' ============================================================
Private Function ReportAutomation_BannerRequested(ByVal bannerName As String, ByVal setting As String) As Boolean
    Dim tokens() As String
    tokens = Split(setting, ",")
    Dim t As Long
    For t = LBound(tokens) To UBound(tokens)
        If Trim$(tokens(t)) = bannerName Then
            ReportAutomation_BannerRequested = True
            Exit Function
        End If
    Next t
End Function
