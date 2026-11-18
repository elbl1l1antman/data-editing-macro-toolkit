Attribute VB_Name = "SampleDesignCandidateUI"
Option Explicit

Public Sub SampleCandidate_Start()
    On Error GoTo Failed
    Dim outputWs As Worksheet, oldScreenUpdating As Boolean, oldDisplayAlerts As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    Dim rowAxes As Range, columnAxes As Range, populations As Range, tableRange As Range, inputMode As String
    inputMode = Toolkit_ButtonMenu("표본설계 1/5 - 입력 방식", "모집단 집계표의 구성을 어떻게 지정할까요?", _
        "1" & vbTab & "표 전체를 선택해 자동 인식" & vbLf & "2" & vbTab & "구분과 값 범위를 직접 지정", "1")
    If Len(inputMode) = 0 Then Exit Sub
    If inputMode = "1" Then
        If Not SampleCandidate_PickRange("표본설계 - 집계표 선택", _
            "머리글, 행 구분, 모집단 숫자를 포함한 집계표 하나를 선택하세요.", tableRange) Then Exit Sub
        If Not SampleCandidate_DetectTable(tableRange, rowAxes, columnAxes, populations) Then
            If MsgBox("선택한 표의 구조를 확실하게 인식하지 못했습니다." & vbCrLf & _
                "행 구분과 머리글, 숫자 영역을 직접 선택할까요?", vbQuestion + vbYesNo + vbDefaultButton2, "표본설계") <> vbYes Then Exit Sub
            inputMode = "2"
        End If
    ElseIf inputMode <> "2" Then
        Err.Raise 5, , "입력 방식을 다시 선택하세요."
    End If
    If inputMode = "2" Then
        If Not SampleCandidate_PickRange("표본설계 1/5", _
            "표의 세로 구분을 선택하세요. 예: 지역 또는 산업명 데이터 셀. 헤더는 제외합니다.", rowAxes) Then Exit Sub
        If Not SampleCandidate_PickRange("표본설계 2/5", _
            "표의 가로 머리글을 모두 선택하세요. 성별과 연령처럼 여러 행이면 함께 선택합니다.", columnAxes) Then Exit Sub
        If Not SampleCandidate_PickRange("표본설계 3/5", _
            "모집단 숫자가 들어 있는 본문을 선택하세요. 행과 열은 앞에서 선택한 구분과 맞아야 합니다.", populations) Then Exit Sub
    End If
    If populations.Worksheet.Parent.IsAddin Then Err.Raise 5, , "작업할 통합문서를 선택하세요."

    Dim records As Collection, rec As Variant, sumPopulation As Double, i As Long
    Set records = SampleCandidate_ReadTable(rowAxes, columnAxes, populations)
    If records.Count = 0 Then Err.Raise 5, , "선택한 범위에 계산할 모집단이 없습니다."
    For i = 1 To records.Count
        rec = records(i): sumPopulation = sumPopulation + CDbl(rec(UBound(rec) - 2))
    Next i
    Dim outputName As String
    outputName = "표본설계_새결과"
    If Not SampleCandidate_Preview(records, rowAxes, columnAxes, populations, sumPopulation, outputName) Then Exit Sub

    Dim choices As String, axisCount As Long, j As Long, selection As String, parts As Variant
    axisCount = rowAxes.Columns.Count + columnAxes.Rows.Count
    choices = ""
    For j = 1 To axisCount
        If Len(choices) > 0 Then choices = choices & vbLf
        If j <= rowAxes.Columns.Count Then
            choices = choices & CStr(j) & vbTab & "세로 구분 " & j & " (예: " & CStr(records(1)(j - 1)) & ")"
        Else
            choices = choices & CStr(j) & vbTab & "가로 머리글 " & (j - rowAxes.Columns.Count) & " (예: " & CStr(records(1)(j - 1)) & ")"
        End If
    Next j
    selection = Toolkit_ButtonMenu("표본설계 4/5 - 목표 구분", "목표 표본수를 어떻게 지정할까요?", _
        "1" & vbTab & "전체 목표 하나 입력" & vbLf & "2" & vbTab & "지역·산업 등 그룹별로 입력", "1")
    If Len(selection) = 0 Then Exit Sub
    Dim indexes As Variant
    If selection = "1" Then
        indexes = Array()
    ElseIf selection = "2" Then
        If axisCount > 30 Then Err.Raise 5, , "그룹 구분은 최대 30개입니다. 필요한 구분만 선택해 다시 실행하세요."
        selection = Toolkit_CheckListMenu("그룹 구분 선택", _
            "목표를 따로 지정할 구분을 선택하세요. 여러 구분을 선택하면 해당 조합마다 목표를 입력합니다.", choices, "1")
        If Len(selection) = 0 Then Exit Sub
        parts = Split(selection, ",")
        ReDim indexes(0 To UBound(parts))
        For j = 0 To UBound(parts)
            If Not IsNumeric(parts(j)) Then Err.Raise 5, , "목표 구분을 다시 선택하세요."
            If CLng(parts(j)) < 1 Or CLng(parts(j)) > axisCount Then Err.Raise 5, , "목표 구분이 올바르지 않습니다."
            indexes(j) = CLng(parts(j)) - 1
        Next j
    Else
        Err.Raise 5, , "목표 지정 방식을 다시 선택하세요."
    End If

    Dim grouped As Collection, targetCounts As Object, groupNames As Object, groupPops As Object
    Set grouped = SampleCandidate_GroupRecords(records, indexes)
    Set targetCounts = CreateObject("Scripting.Dictionary")
    Set groupNames = CreateObject("Scripting.Dictionary")
    Set groupPops = CreateObject("Scripting.Dictionary")
    Dim groupKey As String, groupLabel As String, index As Variant
    For i = 1 To records.Count
        rec = records(i): groupKey = SampleCandidate_GroupKey(rec, indexes)
        If Not groupPops.Exists(groupKey) Then
            groupLabel = ""
            For Each index In indexes
                If Len(groupLabel) > 0 Then groupLabel = groupLabel & " / "
                groupLabel = groupLabel & CStr(rec(CLng(index)))
            Next index
            If Len(groupLabel) = 0 Then groupLabel = "전체"
            groupNames.Add groupKey, groupLabel
            groupPops.Add groupKey, 0#
        End If
        groupPops(groupKey) = CDbl(groupPops(groupKey)) + CDbl(rec(UBound(rec) - 2))
    Next i
    Dim key As Variant, answer As Variant
    For Each key In groupPops.Keys
        answer = Application.InputBox("목표 그룹: " & CStr(groupNames(key)) & vbCrLf & _
            "모집단: " & Format$(CDbl(groupPops(key)), "#,##0") & vbCrLf & _
            "이 그룹의 목표 표본수를 입력하세요.", "표본설계 5/5 - 목표 입력", 0, Type:=1)
        If VarType(answer) = vbBoolean Then Exit Sub
        If CDbl(answer) < 0 Or CDbl(answer) <> Fix(CDbl(answer)) Or CDbl(answer) > 2147483647# Then Err.Raise 5, , "목표 표본수는 0 이상의 정수여야 합니다."
        If CDbl(answer) > CDbl(groupPops(key)) Then Err.Raise 5, , "목표 그룹 '" & CStr(groupNames(key)) & "'의 표본수가 모집단보다 큽니다."
        targetCounts.Add CStr(key), CLng(answer)
    Next key

    Dim method As String, alpha As Double, stageMode As String, stageAxis As Long
    Dim firstMethod As String, firstAlpha As Double, stageChoices As String, isUsed As Boolean
    stageMode = Toolkit_ButtonMenu("배분 단계", "목표 그룹 안에서 바로 배분하거나, 구분별로 먼저 나눈 뒤 배분할 수 있습니다.", _
        "1" & vbTab & "한 번에 배분" & vbLf & "2" & vbTab & "두 단계로 배분", "1")
    If Len(stageMode) = 0 Then Exit Sub
    If stageMode = "2" Then
        If axisCount > 20 Then Err.Raise 5, , "단계별 배분은 최대 20개 구분에서 선택할 수 있습니다."
        For j = 1 To axisCount
            isUsed = False
            For Each index In indexes
                If CLng(index) = j - 1 Then isUsed = True
            Next index
            If Not isUsed Then
                If Len(stageChoices) > 0 Then stageChoices = stageChoices & vbLf
                stageChoices = stageChoices & CStr(j) & vbTab & "구분 " & j & " (예: " & CStr(records(1)(j - 1)) & ")"
            End If
        Next j
        If Len(stageChoices) = 0 Then Err.Raise 5, , "1단계에서 배분할 구분이 남아 있지 않습니다."
        selection = Toolkit_ButtonMenu("1단계 구분", "목표 그룹 안에서 먼저 배분할 구분을 선택하세요.", stageChoices, "")
        If Len(selection) = 0 Then Exit Sub
        If Not IsNumeric(selection) Then Err.Raise 5, , "1단계 구분을 확인하세요."
        stageAxis = CLng(selection) - 1
        If Not SampleCandidate_ChooseMethod("1단계 배분 방식", firstMethod, firstAlpha) Then Exit Sub
    ElseIf stageMode <> "1" Then
        Err.Raise 5, , "배분 단계를 확인하세요."
    End If
    If Not SampleCandidate_ChooseMethod(IIf(stageMode = "2", "2단계 배분 방식", "표본 배분 방식"), method, alpha) Then Exit Sub
    Dim minimum As Long, maximum As Long, constraintChoice As String
    If stageMode = "2" Then
        constraintChoice = "1"
    Else
        constraintChoice = Toolkit_ButtonMenu("표본 배분 제약", "각 셀의 최소·최대 표본수를 지정할 수 있습니다.", _
            "1" & vbTab & "추가 제약 없이 배분" & vbLf & "2" & vbTab & "최소·최대 표본수 지정", "1")
    End If
    If Len(constraintChoice) = 0 Then Exit Sub
    If constraintChoice = "2" Then
        answer = Application.InputBox("각 셀의 최소 표본수를 입력하세요. 적용하지 않으려면 0을 입력합니다.", "최소 표본수", 0, Type:=1)
        If VarType(answer) = vbBoolean Then Exit Sub
        If CDbl(answer) < 0 Or CDbl(answer) <> Fix(CDbl(answer)) Or CDbl(answer) > 2147483647# Then Err.Raise 5, , "최소 표본수는 0 이상의 정수여야 합니다."
        minimum = CLng(answer)
        answer = Application.InputBox("각 셀의 최대 표본수를 입력하세요. 별도 제한이 없으면 0을 입력합니다.", "최대 표본수", 0, Type:=1)
        If VarType(answer) = vbBoolean Then Exit Sub
        If CDbl(answer) < 0 Or CDbl(answer) <> Fix(CDbl(answer)) Or CDbl(answer) > 2147483647# Then Err.Raise 5, , "최대 표본수는 0 이상의 정수여야 합니다."
        maximum = CLng(answer)
    ElseIf constraintChoice <> "1" Then
        Err.Raise 5, , "표본 배분 제약을 다시 선택하세요."
    End If
    Dim allocated As Collection, stageReview As Collection
    If stageMode = "2" Then
        Set allocated = SampleCandidate_AllocateStages(records, targetCounts, indexes, stageAxis, _
            firstMethod, firstAlpha, method, alpha, stageReview)
    Else
        Set allocated = SampleCandidate_Allocate(grouped, targetCounts, 0, method, alpha, minimum, maximum)
    End If
    Dim totalTarget As Double, zeroGroups As Long, targetSummary As String
    For Each key In targetCounts.Keys
        totalTarget = totalTarget + CDbl(targetCounts(key))
        If CLng(targetCounts(key)) = 0 Then zeroGroups = zeroGroups + 1
    Next key
    targetSummary = "전체 목표: " & Format$(totalTarget, "#,##0") & "개" & vbCrLf & _
                    "그룹 " & targetCounts.Count & "개, 배분 셀 " & allocated.Count & "개"
    If stageMode = "2" Then targetSummary = targetSummary & vbCrLf & "두 단계 배분: 추가 최소·최대 제약 없이 모집단 한도 적용"
    If zeroGroups > 0 Then targetSummary = targetSummary & vbCrLf & "목표가 0인 그룹: " & zeroGroups & "개"
    If MsgBox("결과 시트를 새로 만들까요?" & vbCrLf & targetSummary & vbCrLf & _
        "원본은 변경하지 않습니다.", vbQuestion + vbYesNo + vbDefaultButton2, "표본설계 최종 확인") <> vbYes Then Exit Sub

    Dim wb As Workbook
    Set wb = populations.Worksheet.Parent
    outputName = SampleCandidate_UniqueSheetName(wb, outputName)
    Application.ScreenUpdating = False
    Set outputWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outputWs.Name = outputName
    SampleCandidate_WriteOutput outputWs, rowAxes, columnAxes, populations, allocated, method, alpha
    If stageMode = "2" Then SampleCandidate_WriteStageSummary outputWs, stageReview, firstMethod, firstAlpha
    Application.ScreenUpdating = oldScreenUpdating
    outputWs.Activate
    MsgBox "표본설계 결과를 만들었습니다: " & outputName & vbCrLf & _
        "최종 표본수는 고정값이며, 이론 배분은 제약 적용 전 수식입니다.", vbInformation, "표본설계"
    Exit Sub
Failed:
    Dim errorText As String: errorText = Err.Description
    On Error Resume Next
    If Not outputWs Is Nothing Then
        Application.DisplayAlerts = False
        outputWs.Delete
        Application.DisplayAlerts = oldDisplayAlerts
    End If
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
    MsgBox "표본설계를 완료하지 못했습니다: " & errorText & vbCrLf & "원본은 변경하지 않았습니다.", vbExclamation, "표본설계"
End Sub

Private Function SampleCandidate_ChooseMethod(ByVal title As String, ByRef method As String, ByRef alpha As Double) As Boolean
    Dim choice As String, answer As Variant
    choice = Toolkit_ButtonMenu(title, "모집단을 기준으로 배분할 방식을 선택하세요.", _
        "1" & vbTab & "모집단에 비례" & vbLf & "2" & vbTab & "제곱근 비례" & vbLf & _
        "3" & vbTab & "멱배분" & vbLf & "4" & vbTab & "균등 배분", "1")
    If Len(choice) = 0 Then Exit Function
    alpha = 0.5
    Select Case choice
        Case "1": method = "proportional"
        Case "2": method = "sqrt"
        Case "3"
            method = "power"
            answer = Application.InputBox("멱배분 지수를 입력하세요. 1은 비례, 0.5는 제곱근 비례입니다.", "멱배분 지수", 0.5, Type:=1)
            If VarType(answer) = vbBoolean Then Exit Function
            alpha = CDbl(answer)
        Case "4": method = "equal"
        Case Else: Err.Raise 5, , "배분 방식을 다시 선택하세요."
    End Select
    SampleCandidate_ChooseMethod = True
End Function

Private Function SampleCandidate_PickRange(ByVal title As String, ByVal guidance As String, ByRef selectedRange As Range) As Boolean
    On Error Resume Next
    Set selectedRange = Application.InputBox(guidance & vbCrLf & vbCrLf & _
        "입력칸에 주소를 입력하거나 시트에서 마우스 또는 키보드로 범위를 선택하세요.", _
        title, Default:="", Type:=8)
    If Err.Number <> 0 Then
        Err.Clear
        Set selectedRange = Nothing
    End If
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Function
    If selectedRange.Areas.Count <> 1 Then
        MsgBox "떨어진 여러 영역은 선택할 수 없습니다. 이어진 범위 하나를 선택하세요.", vbExclamation, title
        Set selectedRange = Nothing
        Exit Function
    End If
    SampleCandidate_PickRange = True
End Function

Private Function SampleCandidate_Preview(ByVal records As Collection, ByVal rowAxes As Range, _
    ByVal columnAxes As Range, ByVal populations As Range, ByVal sumPopulation As Double, ByRef outputName As String) As Boolean
    Dim summary As String, sample As String, rec As Variant, i As Long, j As Long, form As Object
    summary = "시트: " & populations.Worksheet.Name & vbCrLf & _
        "행 구분: " & rowAxes.Address(False, False) & vbCrLf & _
        "가로 머리글: " & columnAxes.Address(False, False) & vbCrLf & _
        "모집단: " & populations.Address(False, False) & vbCrLf & _
        "유효 셀: " & records.Count & "개 / 합계: " & Format$(sumPopulation, "#,##0") & vbCrLf & _
        "합계·소계와 빈 셀은 제외합니다. 구분과 합계를 확인하세요."
    For i = 1 To Application.Min(records.Count, 20)
        rec = records(i)
        For j = 0 To UBound(rec) - 3
            sample = sample & CStr(rec(j)) & vbTab
        Next j
        sample = sample & CStr(rec(UBound(rec) - 2)) & vbCrLf
    Next i
    Set form = VBA.UserForms.Add("frmPopulationListifyPreview")
    form.Configure summary, sample, "가로 구분", "모집단수", outputName
    form.Caption = "표본설계 - 구조 미리보기"
    form.Controls("lblHelp").Caption = "인식한 구분과 모집단을 확인하고, 맞으면 다음을 누르세요."
    form.Controls("lblColumnAxis").Visible = False
    form.Controls("txtColumnAxisName").Visible = False
    form.Controls("lblValueName").Visible = False
    form.Controls("txtValueName").Visible = False
    form.Controls("cmdRun").Caption = "다음"
    form.Show
    If CBool(form.Confirmed) Then
        outputName = CStr(form.OutputSheetName)
        SampleCandidate_Preview = True
    End If
    Unload form
End Function

Private Function SampleCandidate_UniqueSheetName(ByVal wb As Workbook, ByVal requested As String) As String
    Dim name As String, character As Variant, suffix As String, i As Long, ws As Worksheet
    name = Trim$(requested)
    For Each character In Array("/", "\", "[", "]", "*", "?", ":")
        name = Replace(name, CStr(character), "_")
    Next character
    name = Replace(name, "'", "_")
    If Len(name) = 0 Then name = "표본설계_새결과"
    name = Left$(name, 25)
    For i = 1 To 9999
        suffix = "": If i > 1 Then suffix = " (" & CStr(i) & ")"
        On Error Resume Next
        Set ws = wb.Worksheets(Left$(name, 31 - Len(suffix)) & suffix)
        On Error GoTo 0
        If ws Is Nothing Then
            SampleCandidate_UniqueSheetName = Left$(name, 31 - Len(suffix)) & suffix
            Exit Function
        End If
        Set ws = Nothing
    Next i
    Err.Raise 5, , "결과 시트 이름을 만들 수 없습니다."
End Function
