Attribute VB_Name = "KosisWizard"
Option Explicit

Private Const KOSIS_LAST_SECTION As String = "KOSIS_Last"

Public Sub Ribbon_KosisStart(ByVal control As IRibbonControl)
    Kosis_Start
End Sub

Public Sub Ribbon_KosisSetKey(ByVal control As IRibbonControl)
    Kosis_SetApiKey
End Sub

Public Sub Ribbon_KosisRefresh(ByVal control As IRibbonControl)
    Kosis_RefreshCache
End Sub

Public Sub Kosis_RefreshCache()
    Dim tableId As String, rows As Collection, started As Boolean, failure As Long
    On Error GoTo Failed
    tableId = Kosis_Button("KOSIS 분류정보 갱신", "갱신할 통계표를 선택하세요.", _
        "DT_1B04006" & vbTab & "주민등록인구" & vbLf & _
        "DT_1K52F01" & vbTab & "사업체구분별 (11차 산업분류)" & vbLf & _
        "DT_1K52F03" & vbTab & "종사자규모별 (11차 산업분류)" & vbLf & _
        "DT_1K52D01" & vbTab & "사업체구분별 (10차 산업분류)" & vbLf & _
        "DT_1K52D03" & vbTab & "종사자규모별 (10차 산업분류)", "")
    If Len(tableId) = 0 Then Exit Sub
    Toolkit_ProgressShow "KOSIS 분류정보 갱신", "분류와 시점 정보를 내려받고 있습니다.", 2: started = True
    Set rows = Kosis_ReadMetadata(tableId, True)
    Toolkit_ProgressUpdate 1, 2, "시점 정보를 내려받고 있습니다."
    Set rows = Kosis_ReadPeriods(tableId, True)
    Toolkit_ProgressClose
    started = False
    MsgBox "분류와 시점 정보를 갱신했습니다.", vbInformation, "KOSIS 분류정보"
    Exit Sub
Failed:
    failure = Err.Number
    If started Then Toolkit_ProgressClose
    If failure <> TOOLKIT_CANCEL_ERROR_NUMBER Then MsgBox "갱신을 완료하지 못했습니다. 인증키와 네트워크를 확인하세요. 이미 갱신된 정보가 있을 수 있습니다.", vbExclamation, "KOSIS 분류정보"
End Sub

Public Sub Kosis_Start()
    Dim wb As Workbook, choice As String, edition As String, tableId As String, title As String
    Dim axis2Id As String, axis3Id As String, axes As Variant, metadata As Collection, periodsMeta As Collection
    Dim periods As Collection, items As Collection, regions As Collection, axis2 As Collection, axis3 As Collection
    Dim rowAxes As Collection, bands As Collection, groups As New Collection, outputs As Collection
    Dim raw As Collection, pivot As Collection, missing As Collection, review As Collection, overlaps As Collection
    Dim mode As String, spec As String, defaults As Boolean, started As Boolean, topAge As Long, uncovered As Long
    Dim periodCode As Variant, item As Variant, group As Object, id As Variant, msg As String, rowText As String
    Dim expected As Double, perGroup As Long, progress As Long, period As String, periodType As String, startedFetch As Boolean
    Dim failure As Long, details As String, tableChoice As String, tableDefault As String, rowDefault As String, modeDefault As String, saved As Boolean, cacheTime As String
    Dim picked As Collection, goBack As Boolean, previousTable As String
    On Error GoTo Failed
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise 5, , "결과를 저장할 통합문서를 먼저 여세요."
    If wb Is ThisWorkbook Or wb.ReadOnly Or wb.ProtectStructure Then Err.Raise 5, , "편집할 수 있는 일반 통합문서를 활성화하세요."
    If Not Kosis_ValidKey(GetSetting(BIZ_APP_NAME, "KOSIS", "ApiKey", "")) Then
        MsgBox "먼저 KOSIS 인증키 설정을 실행하세요.", vbExclamation, "KOSIS 모집단"
        Exit Sub
    End If
    If Len(GetSetting(BIZ_APP_NAME, KOSIS_LAST_SECTION, "Table", "")) > 0 Then
        choice = CStr(MsgBox("직전 조회 조건을 불러오시겠습니까? 시점은 조회 전에 다시 확인합니다.", vbYesNoCancel + vbQuestion, "KOSIS 모집단"))
        If CLng(choice) = vbCancel Then Exit Sub
        defaults = (CLng(choice) = vbYes)
    End If
    If defaults Then
        tableDefault = GetSetting(BIZ_APP_NAME, KOSIS_LAST_SECTION, "Table", "")
        If tableDefault = "DT_1B04006" Then
            tableDefault = "population"
        ElseIf Right$(tableDefault, 2) = "01" Then
            tableDefault = "type"
        ElseIf Right$(tableDefault, 2) = "03" Then
            tableDefault = "size"
        End If
    End If
SelectTable:
    tableChoice = Kosis_Button("KOSIS 모집단", "가져올 통계자료를 선택하세요.", _
        "population" & vbTab & "주민등록인구: 시군구 / 1세별" & vbLf & _
        "type" & vbTab & "전국사업체조사: 사업체구분별" & vbLf & _
        "size" & vbTab & "전국사업체조사: 종사자규모별", tableDefault)
    If Len(tableChoice) = 0 Then Exit Sub
    If tableChoice = "population" Then
        tableId = "DT_1B04006": title = "행정구역(시군구)별/1세별 주민등록인구"
        axis2Id = "YRE": axis3Id = "": axes = Array("A", "YRE")
    Else
        edition = "F"
        If defaults Then
            If Mid$(GetSetting(BIZ_APP_NAME, KOSIS_LAST_SECTION, "Table", ""), 8, 1) = "D" Then edition = "D"
        End If
        edition = Kosis_Button("산업분류 개정판", "기준 시점에 제공되는 산업분류를 선택하세요.", _
            "F" & vbTab & "11차 산업분류" & vbLf & "D" & vbTab & "10차 산업분류" & vbLf & "__BACK__" & vbTab & "이전 단계", edition)
        If Len(edition) = 0 Then Exit Sub
        If edition = "__BACK__" Then GoTo SelectTable
        If tableChoice = "type" Then
            tableId = "DT_1K52" & edition & "01": title = "시도·산업·사업체구분별 사업체수, 종사자수"
            axis2Id = "B": axis3Id = "C": axes = Array("A", "B", "C")
        ElseIf tableChoice = "size" Then
            tableId = "DT_1K52" & edition & "03": title = "시도·산업·종사자규모별 사업체수, 종사자수"
            axis2Id = "B": axis3Id = "E": axes = Array("A", "B", "E")
        Else
            Err.Raise 5, , "자료 선택을 확인하세요."
        End If
    End If
    If Len(previousTable) > 0 And previousTable <> tableId Then
        Set periods = Nothing: Set items = Nothing: Set regions = Nothing
        Set axis2 = Nothing: Set axis3 = Nothing: Set rowAxes = Nothing
        Set bands = Nothing: spec = "": mode = "": defaults = False
    End If
    previousTable = tableId
    If defaults Then defaults = (GetSetting(BIZ_APP_NAME, KOSIS_LAST_SECTION, "Table", "") = tableId)
    Toolkit_ProgressShow "KOSIS 모집단", "분류와 시점 정보를 확인하고 있습니다.", 1: started = True
    Set metadata = Kosis_ReadMetadata(tableId)
    Set periodsMeta = Kosis_ReadPeriods(tableId)
    cacheTime = Kosis_CacheTimestamp(tableId)
    Toolkit_ProgressClose
    started = False
SelectPeriods:
    Set picked = Kosis_PickItems("KOSIS 시점 선택", "조회할 연도 또는 월을 직접 선택하세요. 분류정보 갱신: " & cacheTime, Kosis_PeriodChoices(periodsMeta), "period", Kosis_Defaults(periods, Kosis_Last(defaults, "Periods")), goBack)
    If goBack Then GoTo SelectTable
    If picked Is Nothing Then Exit Sub
    Set periods = picked
SelectItems:
    Set picked = Kosis_PickItems("KOSIS 항목 선택", "가져올 통계항목을 선택하세요.", metadata, "ITEM", Kosis_Defaults(items, Kosis_Last(defaults, "Items")), goBack)
    If goBack Then GoTo SelectPeriods
    If picked Is Nothing Then Exit Sub
    Set items = picked
SelectRegions:
    msg = Kosis_Defaults(regions, Kosis_Last(defaults, "Regions"))
    If Len(msg) = 0 And tableChoice <> "population" Then msg = "00"
    Set picked = Kosis_PickItems("KOSIS 지역 선택", "조회할 지역을 선택하세요.", metadata, "A", msg, goBack)
    If goBack Then GoTo SelectItems
    If picked Is Nothing Then Exit Sub
    Set regions = picked
SelectAxis2:
    If tableChoice = "population" Then
        modeDefault = "single"
        If Len(Kosis_Last(defaults, "AgeSpec")) > 0 Then modeDefault = "bands"
        mode = Kosis_Button("KOSIS 연령", "연령을 표시할 방식을 선택하세요.", _
            "single" & vbTab & "1세별 선택" & vbLf & "bands" & vbTab & "직접 연령구간 만들기" & vbLf & "__BACK__" & vbTab & "이전 단계", modeDefault)
        If Len(mode) = 0 Then Exit Sub
        If mode = "__BACK__" Then GoTo SelectRegions
        If mode = "bands" Then
            topAge = Kosis_FindTopAge(metadata)
            If Len(spec) = 0 Then spec = Kosis_Last(defaults, "AgeSpec")
            If Not Kosis_ChooseAgeBands(topAge, spec, bands, spec, uncovered) Then Exit Sub
            Set axis2 = Kosis_AgeSourceCodes(metadata, bands, topAge)
        Else
            Set bands = Nothing: spec = "": uncovered = 0
            Set picked = Kosis_PickItems("KOSIS 연령 선택", "조회할 연령을 선택하세요.", metadata, axis2Id, Kosis_Defaults(axis2, Kosis_Last(defaults, "Axis2")), goBack)
            If goBack Then GoTo SelectAxis2
            If picked Is Nothing Then Exit Sub
            Set axis2 = picked
        End If
    Else
        Set picked = Kosis_PickItems("KOSIS 산업 선택", "산업을 검색하거나 계층을 펼쳐 선택하세요.", metadata, axis2Id, Kosis_Defaults(axis2, Kosis_Last(defaults, "Axis2")), goBack)
        If goBack Then GoTo SelectRegions
        If picked Is Nothing Then Exit Sub
        Set axis2 = picked
    End If
SelectAxis3:
    If axis3 Is Nothing Then Set axis3 = New Collection
    If Len(axis3Id) > 0 Then
        Set picked = Kosis_PickItems("KOSIS 분류 선택", "사업체 구분 또는 종사자 규모를 선택하세요.", metadata, axis3Id, Kosis_Defaults(axis3, Kosis_Last(defaults, "Axis3")), goBack)
        If goBack Then GoTo SelectAxis2
        If picked Is Nothing Then Exit Sub
        Set axis3 = picked
    End If
SelectLayout:
    choice = Kosis_Button("KOSIS 행렬 구성", "행과 열 구분을 설정합니다.", "next" & vbTab & "행과 열 설정" & vbLf & "__BACK__" & vbTab & "이전 단계", "next")
    If Len(choice) = 0 Then Exit Sub
    If choice = "__BACK__" Then
        If Len(axis3Id) > 0 Then GoTo SelectAxis3 Else GoTo SelectAxis2
    End If
    rowText = "A" & vbTab & Kosis_AxisName(metadata, "A") & vbLf & axis2Id & vbTab & Kosis_AxisName(metadata, axis2Id)
    If Len(axis3Id) > 0 Then rowText = rowText & vbLf & axis3Id & vbTab & Kosis_AxisName(metadata, axis3Id)
    rowDefault = Kosis_Defaults(rowAxes, Kosis_Last(defaults, "RowAxes"), ","): If Len(rowDefault) = 0 Then rowDefault = "A"
    choice = Toolkit_CheckListMenu("KOSIS 행렬 구성", "행에 둘 구분을 선택하세요. 선택하지 않은 구분은 열에 배치됩니다.", rowText, rowDefault)
    If Len(choice) = 0 Then Exit Sub
    Set rowAxes = Kosis_CollectionFromDelimited(choice, ",")
    If rowAxes.Count >= UBound(axes) + 1 Then
        MsgBox "열에 둘 구분을 하나 이상 남겨주세요.", vbExclamation, "KOSIS 행렬 구성"
        GoTo SelectLayout
    End If
    Set overlaps = Kosis_FindOverlaps(regions, metadata, "A")
    Kosis_AppendOverlaps overlaps, Kosis_FindOverlaps(axis2, metadata, axis2Id)
    If Len(axis3Id) > 0 Then Kosis_AppendOverlaps overlaps, Kosis_FindOverlaps(axis3, metadata, axis3Id)
    If tableChoice = "population" Then
        If Kosis_Contains(items, "T2") And (Kosis_Contains(items, "T3") Or Kosis_Contains(items, "T4")) Then overlaps.Add "총인구와 남녀 인구는 서로 포함됩니다."
    End If
    perGroup = CLng(CDbl(regions.Count) * CDbl(axis2.Count) * IIf(axis3.Count > 0, CDbl(axis3.Count), 1#))
    expected = CDbl(perGroup) * CDbl(periods.Count) * CDbl(items.Count)
    If expected > 100000# Then Err.Raise 5, , "선택한 조합이 10만 개를 초과합니다. 조회 조건을 나누어 주세요."
    msg = "통계표: " & title & " (" & tableId & ")" & vbCrLf & _
          "선택 시점: " & Kosis_JoinCollection(periods, ", ") & vbCrLf & _
          "분류정보 갱신: " & cacheTime & vbCrLf & _
          "지역 " & regions.Count & "개, 분류 " & axis2.Count & "개, 항목 " & items.Count & "개" & vbCrLf & _
          "예상 조합: " & Format$(expected, "#,##0") & "개"
    msg = msg & vbCrLf & "항목: " & Kosis_SelectionSummary(metadata, "ITEM", items)
    msg = msg & vbCrLf & "지역: " & Kosis_SelectionSummary(metadata, "A", regions)
    If bands Is Nothing Then
        msg = msg & vbCrLf & "분류: " & Kosis_SelectionSummary(metadata, axis2Id, axis2)
    Else
        msg = msg & vbCrLf & "연령구간: " & CStr(bands.Count) & "개"
    End If
    If axis3.Count > 0 Then msg = msg & vbCrLf & "상세 구분: " & Kosis_SelectionSummary(metadata, axis3Id, axis3)
    If overlaps.Count > 0 Then msg = msg & vbCrLf & vbCrLf & "상위·하위 선택이 겹칩니다. 값은 별도로 표시하며 중복 합산하지 않습니다. 검토 시트를 확인하세요."
    If uncovered > 0 Then msg = msg & vbCrLf & "지정하지 않은 연령이 있습니다: " & uncovered & "개"
    choice = CStr(MsgBox(msg & vbCrLf & vbCrLf & "예: 조회 / 아니요: 이전 단계 / 취소: 종료", vbYesNoCancel + vbQuestion + vbDefaultButton2, "KOSIS 조회 확인"))
    If CLng(choice) = vbNo Then GoTo SelectLayout
    If CLng(choice) <> vbYes Then Exit Sub
    Toolkit_ProgressShow "KOSIS 모집단", "통계자료를 조회하고 있습니다.", CLng(expected): started = True: startedFetch = True
    progress = 0
    For Each periodCode In periods
        periodType = Left$(CStr(periodCode), 1): period = Mid$(CStr(periodCode), 2)
        For Each item In items
            Set raw = Kosis_FetchBatches(tableId, period, periodType, CStr(item), regions, axis2, axis3, "A", axis2Id, axis3Id, missing, progress, CLng(expected))
            Set review = New Collection
            Kosis_AppendOverlaps review, overlaps
            If uncovered > 0 Then review.Add "지정하지 않은 연령 범주: " & CStr(uncovered) & "개"
            Set pivot = raw
            If Not bands Is Nothing Then
                Dim ageReview As Collection
                Set pivot = Kosis_AggregateAges(raw, metadata, bands, regions, period, CStr(item), topAge, ageReview)
                Kosis_AppendOverlaps review, ageReview
            End If
            Set group = CreateObject("Scripting.Dictionary")
            group.Add "records", pivot: group.Add "rawRecords", raw: group.Add "missing", missing
            group.Add "period", period: group.Add "item", CStr(item): group.Add "review", review: group.Add "tableId", tableId
            groups.Add group
            progress = progress + perGroup
            Toolkit_ProgressUpdate progress, CLng(expected), "KOSIS 통계자료 조회 중"
        Next item
    Next periodCode
    Toolkit_ProgressClose
    started = False
    Set outputs = Kosis_WriteGroups(wb, groups, metadata, rowAxes, axes, title & " (" & tableId & ")")
    saved = Kosis_SaveLast(tableId, periods, items, regions, axis2, axis3, rowAxes, spec)
    outputs(2).Activate
    msg = "조회 결과와 검토 시트를 만들었습니다. 검토 사항과 자료 출처를 확인하세요."
    If Not saved Then msg = msg & vbCrLf & "직전 조회 조건은 저장하지 못했습니다."
    If MsgBox(msg & vbCrLf & vbCrLf & "기존 표본설계 기능을 실행하시겠습니까? 범위는 직접 선택합니다.", vbYesNo + vbInformation + vbDefaultButton2, "KOSIS 모집단") = vbYes Then SampleCandidate_Start
    Exit Sub
Failed:
    failure = Err.Number: details = Err.Description
    If started Then Toolkit_ProgressClose
    If failure = TOOLKIT_CANCEL_ERROR_NUMBER Then
        MsgBox "조회가 취소되었습니다. 원본 시트는 변경되지 않았습니다.", vbInformation, "KOSIS 모집단"
    Else
        If startedFetch Then Kosis_FailureReport wb, title, period, CStr(item), details
        MsgBox "KOSIS 작업을 완료하지 못했습니다. 원본 시트는 변경되지 않았습니다." & vbCrLf & details, vbExclamation, "KOSIS 모집단"
    End If
End Sub

Private Function Kosis_Button(ByVal title As String, ByVal guide As String, ByVal options As String, ByVal defaultKey As String) As String
    Kosis_Button = Toolkit_ButtonMenu(title, guide, options, defaultKey)
End Function

Private Function Kosis_Last(ByVal enabled As Boolean, ByVal name As String) As String
    If enabled Then Kosis_Last = GetSetting(BIZ_APP_NAME, KOSIS_LAST_SECTION, name, "")
End Function

Private Function Kosis_AxisName(ByVal metadata As Collection, ByVal axis As String) As String
    Dim entry As Object
    For Each entry In metadata
        If CStr(entry("axis")) = axis Then Kosis_AxisName = CStr(entry("axisName")): Exit Function
    Next entry
    Err.Raise 5, , "통계표 분류 정보를 확인할 수 없습니다."
End Function

Private Function Kosis_CollectionFromDelimited(ByVal text As String, ByVal delimiter As String) As Collection
    Dim output As New Collection, value As Variant
    For Each value In Split(text, delimiter)
        If Len(CStr(value)) > 0 Then output.Add CStr(value)
    Next value
    Set Kosis_CollectionFromDelimited = output
End Function

Private Function Kosis_JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim value As Variant, result As String
    For Each value In values
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(value)
    Next value
    Kosis_JoinCollection = result
End Function

Private Function Kosis_Defaults(ByVal values As Collection, ByVal fallback As String, Optional ByVal delimiter As String = vbLf) As String
    If values Is Nothing Then Kosis_Defaults = fallback: Exit Function
    If values.Count = 0 Then Kosis_Defaults = fallback Else Kosis_Defaults = Kosis_JoinCollection(values, delimiter)
End Function

Private Function Kosis_SelectionSummary(ByVal metadata As Collection, ByVal axis As String, ByVal values As Collection) As String
    Dim value As Variant, entry As Object, result As String, n As Long, label As String
    For Each value In values
        If n = 5 Then result = result & " 외 " & CStr(values.Count - n) & "개": Exit For
        label = CStr(value)
        For Each entry In metadata
            If CStr(entry("axis")) = axis And CStr(entry("code")) = CStr(value) Then
                label = label & " " & CStr(entry("name"))
                Exit For
            End If
        Next entry
        If Len(result) > 0 Then result = result & ", "
        result = result & label: n = n + 1
    Next value
    Kosis_SelectionSummary = result
End Function

Private Function Kosis_Contains(ByVal values As Collection, ByVal key As String) As Boolean
    Dim value As Variant
    For Each value In values
        If CStr(value) = key Then Kosis_Contains = True: Exit Function
    Next value
End Function

Private Sub Kosis_AppendOverlaps(ByVal target As Collection, ByVal source As Collection)
    Dim value As Variant
    For Each value In source: target.Add CStr(value): Next value
End Sub

Private Function Kosis_SaveLast(ByVal tableId As String, ByVal periods As Collection, ByVal items As Collection, _
    ByVal regions As Collection, ByVal axis2 As Collection, ByVal axis3 As Collection, ByVal rowAxes As Collection, ByVal spec As String) As Boolean
    On Error GoTo Failed
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "Table", ""
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "Periods", Kosis_JoinCollection(periods, vbLf)
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "Items", Kosis_JoinCollection(items, vbLf)
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "Regions", Kosis_JoinCollection(regions, vbLf)
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "Axis2", Kosis_JoinCollection(axis2, vbLf)
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "Axis3", Kosis_JoinCollection(axis3, vbLf)
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "RowAxes", Kosis_JoinCollection(rowAxes, ",")
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "AgeSpec", spec
    SaveSetting BIZ_APP_NAME, KOSIS_LAST_SECTION, "Table", tableId
    Kosis_SaveLast = True
Failed:
End Function

Private Function Kosis_ChooseAgeBands(ByVal topAge As Long, ByVal previous As String, ByRef bands As Collection, _
    ByRef spec As String, ByRef uncovered As Long) As Boolean
    Dim form As Object, initial As String
    initial = previous
    If Len(initial) = 0 Then initial = "1" & vbTab & "0" & vbTab & "이상" & vbTab & "20" & vbTab & "미만" & vbTab & "0~19세" & vbLf & _
        "2" & vbTab & "20" & vbTab & "이상" & vbTab & "40" & vbTab & "미만" & vbTab & "20~39세" & vbLf & _
        "3" & vbTab & "40" & vbTab & "이상" & vbTab & "60" & vbTab & "미만" & vbTab & "40~59세" & vbLf & _
        "4" & vbTab & "60" & vbTab & "이상" & vbTab & "" & vbTab & "미만" & vbTab & "60세 이상"
    On Error GoTo Failed
    Set form = VBA.UserForms.Add("frmSampleDesignBandBuilder")
    form.Configure initial, "연령구간"
    form.Caption = "KOSIS 연령구간 설정"
    form.Controls("lblGuide").Caption = "연령의 시작·종료 조건과 이름을 입력하세요. " & topAge & "세 이상은 하나의 묶음입니다."
    form.Controls("cmdEmployeePreset").Visible = False: form.Controls("cmdRevenuePreset").Visible = False
    form.Controls("cmdOK").Caption = "구간 확인"
    form.Show
    If CBool(form.Cancelled) Then Unload form: Exit Function
    spec = CStr(form.BandSpecText): Unload form: Set form = Nothing
    Set bands = Kosis_ParseAgeBands(spec, topAge, uncovered)
    If uncovered > 0 Then
        If MsgBox("선택하지 않은 연령 범주가 " & uncovered & "개 있습니다. 계속하시겠습니까?", vbYesNo + vbExclamation + vbDefaultButton2, "KOSIS 연령구간") <> vbYes Then Exit Function
    End If
    Kosis_ChooseAgeBands = True
    Exit Function
Failed:
    Dim message As String: message = Err.Description
    If Not form Is Nothing Then Unload form
    MsgBox "연령구간을 확인하세요." & vbCrLf & message, vbExclamation, "KOSIS 연령구간"
End Function

Private Sub Kosis_FailureReport(ByVal wb As Workbook, ByVal title As String, ByVal period As String, ByVal item As String, ByVal message As String)
    Dim ws As Worksheet, name As String, i As Long, probe As Object
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Or wb.ProtectStructure Then Exit Sub
    name = "KOSIS_검토": i = 1
    Do
        Set probe = Nothing: Set probe = wb.Sheets(name)
        If probe Is Nothing Then Exit Do
        i = i + 1: name = "KOSIS_검토_" & CStr(i)
    Loop
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    ws.Range("A1:A4").NumberFormat = "@"
    ws.Range("A1").Value2 = "KOSIS 조회 실패: 완성된 결과 없음"
    ws.Range("A2").Value2 = title
    ws.Range("A3").Value2 = "실패 시점: " & period & " / 항목: " & item
    ws.Range("A4").Value2 = message
    ws.Columns("A").AutoFit
End Sub
