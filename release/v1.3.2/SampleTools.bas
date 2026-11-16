Attribute VB_Name = "SampleTools"
Option Explicit


' ============================================================
' 표본설계/표본할당 도구
' ============================================================

Public Sub SampleDesign_CreateTemplate()
    On Error GoTo Failed
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "표본설계 시트를 만들 통합문서를 찾지 못했습니다.", vbExclamation, "표본설계"
        Exit Sub
    End If

    If wb.IsAddin Then
        MsgBox "작업할 통합문서를 먼저 여세요.", vbExclamation, "표본설계"
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
    Dim exampleWs As Worksheet
    Set guideWs = SampleDesign_FindSheet(wb, "표본설계_가이드")
    If guideWs Is Nothing Then
        Set guideWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_가이드")
        SampleDesign_WriteGuideSheet guideWs, setup, templateMode
    End If
    Set inputWs = SampleDesign_FindInputSheet(wb)
    If inputWs Is Nothing Then
        Set inputWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_입력")
        SampleDesign_WriteInputTemplate inputWs, setup, False
    End If
    If templateMode = "table" Then
        Set tableWs = SampleDesign_FindSheet(wb, "표본설계_집계표입력")
        If tableWs Is Nothing Then
            Set tableWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_집계표입력")
            SampleDesign_WriteTableInputTemplate tableWs, False
        End If
    End If
    Set resultWs = SampleDesign_FindSheet(wb, "표본설계_할당결과")
    If resultWs Is Nothing Then
        Set resultWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_할당결과")
        SampleDesign_WriteEmptyResult resultWs
    End If
    Set reviewWs = SampleDesign_FindSheet(wb, "표본설계_검토")
    If reviewWs Is Nothing Then
        Set reviewWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_검토")
        SampleDesign_WriteEmptyReview reviewWs
    End If
    Set exampleWs = SampleDesign_FindSheet(wb, "표본설계_예시")
    If exampleWs Is Nothing Then
        Set exampleWs = SampleDesign_PrepareOutputSheet(wb, "표본설계_예시")
        SampleDesign_WriteExampleSheet exampleWs
    End If

    exampleWs.Activate
    exampleWs.Range("A1").Select
    If templateMode = "table" Then
        exampleWs.Hyperlinks.Add Anchor:=exampleWs.Range("G9"), Address:="", _
            SubAddress:="'표본설계_집계표입력'!A10", TextToDisplay:="2. 실제 집계표 붙여넣기"
    Else
        exampleWs.Hyperlinks.Add Anchor:=exampleWs.Range("G9"), Address:="", _
            SubAddress:="'표본설계_입력'!A21", TextToDisplay:="2. 실제 자료 입력하기"
    End If
    MsgBox "표본설계_예시에서 자료 형태와 작업 순서를 먼저 확인하세요." & vbCrLf & _
           "연습: [기본 예시로 연습하기]를 누른 뒤 입력 시트의 [할당 계산]을 실행합니다." & vbCrLf & _
           "실제 작업: [실제 자료 입력하기] 또는 [실제 집계표 붙여넣기] 링크로 이동합니다." & vbCrLf & _
           "기존 입력값, 설정, 결과는 유지했습니다. 새 설정은 새로 만든 입력 시트에만 적용됩니다.", _
           vbInformation, "표본설계 시작 안내"
    Exit Sub
Failed:
    MsgBox "표본설계 시트를 준비하지 못했습니다: " & Err.Description & vbCrLf & _
           "기존 입력 시트는 지우지 않습니다. 통합문서 보호 여부를 확인하세요.", vbExclamation, "표본설계"
End Sub

Public Sub SampleDesign_UseExample()
    On Error GoTo Failed
    Dim wb As Workbook, ws As Worksheet
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If wb.IsAddin Then Exit Sub
    If SampleDesign_FindSheet(wb, "표본설계_예시") Is Nothing Then
        MsgBox "먼저 [설계 시트 생성]으로 예시와 실제 입력 시트를 준비하세요.", vbInformation, "표본설계"
        Exit Sub
    End If
    Set ws = SampleDesign_FindInputSheet(wb)
    If Not ws Is Nothing Then
        If MsgBox("입력 시트의 현재 내용과 설정을 기본 연습 자료로 바꿀까요?" & vbCrLf & _
                  "실제 자료가 있다면 [아니요]를 선택하세요. 할당결과와 검토 시트는 바꾸지 않습니다.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "예시로 연습하기") <> vbYes Then Exit Sub
        If ws.ProtectContents Then Err.Raise 1004, , "입력 시트가 보호되어 있습니다."
    Else
        Set ws = SampleDesign_PrepareOutputSheet(wb, "표본설계_입력")
    End If
    SampleDesign_WritePracticeInput ws
    ws.Activate
    ws.Range("A21").Select
    MsgBox "연습 자료를 채웠습니다. 입력 시트 위쪽 [할당 계산]을 누르세요." & vbCrLf & _
           "단순 비례를 선택하면 100개 / 50개 / 30개, 합계 180개가 배정됩니다." & vbCrLf & _
           "이후 할당결과와 검토 시트에서 합계와 경고를 확인하세요.", vbInformation, "표본설계 연습"
    Exit Sub
Failed:
    MsgBox "연습 자료를 준비하지 못했습니다: " & Err.Description, vbExclamation, "표본설계"
End Sub

Private Sub SampleDesign_WritePracticeInput(ByVal ws As Worksheet)
    Dim setup As Object
    Set setup = SampleDesign_DefaultSetup()
    setup("targetN") = 180
    SampleDesign_WriteInputTemplate ws, setup, True
    ws.Range("A2").Value = "연습자료입니다. 실제 분석 전 21행 이하 예시 행을 모두 지우고 실제 모집단수와 목표 표본수를 입력하세요."
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
    Application.ScreenUpdating = True

    MsgBox "구간 열을 생성했습니다." & vbCrLf & _
           "수식 입력: " & rowCount & "행" & vbCrLf & _
           "결과 열에는 IFS 수식이 입력되어 기준을 직접 확인하고 수정할 수 있습니다.", _
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
    form.Configure SampleDesign_DefaultEmployeeBandSpec(), outputHeader
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

    Dim presetInput As Variant
    presetInput = Application.InputBox("기본 예시를 선택하세요." & vbCrLf & _
        "1 = 종사자 구간" & vbCrLf & _
        "2 = 매출액 구간" & vbCrLf & _
        "3 = 직접 입력", "표본설계 구간 만들기", 1, Type:=1)
    If VarType(presetInput) = vbBoolean And presetInput = False Then Exit Function
    If CLng(presetInput) = 1 Then
        specText = SampleDesign_DefaultEmployeeBandSpec()
        SampleDesign_ShowBandBuilderFallback = True
        Exit Function
    ElseIf CLng(presetInput) = 2 Then
        specText = SampleDesign_DefaultRevenueBandSpec()
        SampleDesign_ShowBandBuilderFallback = True
        Exit Function
    End If

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
    formulaText = "=IF(" & sourceAddress & "=" & q & q & "," & SampleDesign_FormulaStringLiteral("자료 없음") & ",IFERROR(IFS("

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

Private Function SampleDesign_DefaultEmployeeBandSpec() As String
    SampleDesign_DefaultEmployeeBandSpec = _
        "1" & vbTab & "" & vbTab & "이상" & vbTab & "5" & vbTab & "미만" & vbTab & "01. 5인 미만" & vbLf & _
        "2" & vbTab & "5" & vbTab & "이상" & vbTab & "10" & vbTab & "미만" & vbTab & "02. 5-9명" & vbLf & _
        "3" & vbTab & "10" & vbTab & "이상" & vbTab & "20" & vbTab & "미만" & vbTab & "03. 10-19명" & vbLf & _
        "4" & vbTab & "20" & vbTab & "이상" & vbTab & "50" & vbTab & "미만" & vbTab & "04. 20-49명" & vbLf & _
        "5" & vbTab & "50" & vbTab & "이상" & vbTab & "100" & vbTab & "미만" & vbTab & "05. 50-99명" & vbLf & _
        "6" & vbTab & "100" & vbTab & "이상" & vbTab & "300" & vbTab & "미만" & vbTab & "06. 100-299명" & vbLf & _
        "7" & vbTab & "300" & vbTab & "이상" & vbTab & "" & vbTab & "미만" & vbTab & "07. 300명 이상"
End Function

Private Function SampleDesign_DefaultRevenueBandSpec() As String
    SampleDesign_DefaultRevenueBandSpec = _
        "1" & vbTab & "" & vbTab & "이상" & vbTab & "100000000" & vbTab & "미만" & vbTab & "01. 1억 미만" & vbLf & _
        "2" & vbTab & "100000000" & vbTab & "이상" & vbTab & "500000000" & vbTab & "미만" & vbTab & "02. 1억-5억 미만" & vbLf & _
        "3" & vbTab & "500000000" & vbTab & "이상" & vbTab & "1000000000" & vbTab & "미만" & vbTab & "03. 5억-10억 미만" & vbLf & _
        "4" & vbTab & "1000000000" & vbTab & "이상" & vbTab & "5000000000" & vbTab & "미만" & vbTab & "04. 10억-50억 미만" & vbLf & _
        "5" & vbTab & "5000000000" & vbTab & "이상" & vbTab & "10000000000" & vbTab & "미만" & vbTab & "05. 50억-100억 미만" & vbLf & _
        "6" & vbTab & "10000000000" & vbTab & "이상" & vbTab & "" & vbTab & "미만" & vbTab & "06. 100억 이상"
End Function

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

Private Sub SampleDesign_WriteInputTemplate(ByVal ws As Worksheet, ByVal setup As Object, Optional ByVal includeExamples As Boolean = True)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 입력"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2:O2").Merge

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
    If includeExamples Then
        For exampleRow = 21 To 23
            SampleDesign_WritePopulationExample ws, exampleRow, headers, enabled, optionStartCol
        Next exampleRow
    End If

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
    ws.Columns("A:O").ColumnWidth = 16
    ws.Columns("A").ColumnWidth = 22
    ws.Range("A2").Value = "예시는 표본설계_예시에서 확인하세요. 이 시트 21행부터 실제 쿼터 조합과 모집단수를 입력합니다."
    ws.Range("A2:O2").WrapText = True
    ws.Rows(2).RowHeight = 30
    Dim helpRow As Long
    ws.Range("D3:I8").UnMerge
    For helpRow = 3 To 8
        ws.Range(ws.Cells(helpRow, 4), ws.Cells(helpRow, 9)).Merge
    Next helpRow
    ws.Range("D3:I8").WrapText = True
    ws.Rows("3:8").RowHeight = 28
    SampleDesign_AddActionButton ws, "sdCalculate", "할당 계산", "SampleDesign_RunAllocation"
    SampleDesign_AddNavigation ws
End Sub

Private Sub SampleDesign_AddActionButton(ByVal ws As Worksheet, ByVal buttonName As String, _
                                       ByVal captionText As String, ByVal macroName As String)
    Dim button As Shape, area As Range
    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo 0
    Set area = ws.Range("D9:F10")
    Set button = ws.Shapes.AddFormControl(xlButtonControl, area.Left, area.Top, area.Width, area.Height)
    button.Name = buttonName
    button.TextFrame.Characters.Text = captionText
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Private Sub SampleDesign_AddNavigation(ByVal ws As Worksheet)
    ws.Range("G9:I10").Merge
    ws.Range("A9:C10").Merge
    ws.Range("A9:I10").WrapText = True
    ws.Rows("9:10").RowHeight = 24
    ws.Hyperlinks.Add Anchor:=ws.Range("G9"), Address:="", _
        SubAddress:="'표본설계_예시'!A1", TextToDisplay:="입력 예시와 작업 순서 보기"
    ws.Hyperlinks.Add Anchor:=ws.Range("A9"), Address:="", _
        SubAddress:="'표본설계_가이드'!A1", TextToDisplay:="상세 사용 가이드 보기"
End Sub

Private Sub SampleDesign_WriteExampleSheet(ByVal ws As Worksheet)
    SampleDesign_WritePracticeInput ws
    ws.Range("A1").Value = "표본설계 시작하기 - 예시로 자료 형태 익히기"
    ws.Range("A2").Value = "연습용 가상 자료입니다. 실제 모집단이 아니며 이 시트는 할당 계산에 직접 사용되지 않습니다."
    ws.Range("A2:O2").Interior.Color = RGB(255, 242, 204)
    ws.Range("D3").Value = "1. 아래 20~23행에서 자료 모양을 확인하세요. 한 행은 산업+대분류+지역이 같은 한 쿼터 조합입니다."
    ws.Range("D4").Value = "모집단수는 그 조합에 해당하는 전체 업체/사람 수입니다. 개인별 명부를 그대로 붙여넣지 마세요."
    ws.Range("D5").Value = "2. 연습하려면 아래 [기본 예시로 연습하기]를 누릅니다. 입력 시트에 연습자료가 채워집니다."
    ws.Range("D6").Value = "3. 입력 시트의 [할당 계산]에서 단순 비례를 선택하면 100 / 50 / 30, 합계 180개가 나옵니다."
    ws.Range("D7").Value = "4. 실제 작업은 아래 이동 링크로 시작하세요. 19행 사용여부와 B4 목표 표본수도 확인합니다."
    ws.Range("D8").Value = "5. 결과 시트에서 최종할당 합계를, 검토 시트에서 목표 차이와 경고를 확인하세요."
    ws.Shapes("sdCalculate").Delete
    SampleDesign_AddActionButton ws, "sdPractice", "1. 기본 예시로 연습하기", "SampleDesign_UseExample"
    ws.Hyperlinks.Add Anchor:=ws.Range("G9"), Address:="", _
        SubAddress:="'표본설계_입력'!A21", TextToDisplay:="2. 실제 자료 입력하기"
    ws.Range("A25:O25").Merge
    ws.Range("A25").Value = "실제 자료로 바꾸는 방법"
    ws.Range("A25").Font.Bold = True
    ws.Range("A26:O27").Merge
    ws.Range("A26").Value = "리스트형: 입력 시트의 20행 헤더는 남기고 21행부터 쿼터ID, 사용할 분류값, 모집단수를 붙여넣습니다. 연습했다면 예시 3행을 모두 삭제한 후 실제 자료로 채우세요. 모집단수는 필수이고 1 이상이어야 합니다."
    ws.Range("A28:O29").Merge
    ws.Range("A28").Value = "목표: 전체 목표는 B4에 입력합니다. 산업별·지역별로 목표가 다르면 13~17행 목표표본 그룹표에 조건과 목표를 적습니다. 전체 목표만 사용한다면 그룹별 목표표본수는 비워두세요. 사용 쿼터 변수명은 20행과 그룹표 12행을 동일하게 유지합니다."
    ws.Range("A30:O31").Merge
    ws.Range("A30").Value = "집계표형: [설계 시트 생성]에서 집계표형을 선택한 뒤 집계표입력 시트 10행부터 원래 집계표를 붙여넣고 [할당 계산]을 실행합니다. 자동 변환은 리스트 입력 시트에 모집단 데이터가 없을 때만 수행합니다."
    ws.Range("A26:O31").WrapText = True
    ws.Rows("26:31").RowHeight = 24
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

Private Sub SampleDesign_WriteTableInputTemplate(ByVal ws As Worksheet, Optional ByVal includeExample As Boolean = True)
    ws.Cells.Clear
    ws.Range("A1").Value = "표본설계 집계표입력"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "사용 방법"
    ws.Range("A4").Value = "1. 이 시트에 모집단 구성표를 원본 형태로 붙여넣습니다. 여러 표를 한 번에 붙여넣어도 됩니다."
    ws.Range("A5").Value = "2. [구성표 리스트화] 버튼으로 탐지 결과를 미리 확인할 수 있습니다."
    ws.Range("A6").Value = "3. [할당 계산] 실행 시 표본설계_입력 시트의 모집단 표로 자동 변환을 시도합니다."
    ws.Range("A7").Value = "4. 변환 후 표본설계_입력 시트에서 쿼터명, 목표표본 그룹표, 모집단수를 검토하세요."
    ws.Range("A8").Value = "10행부터 실제 집계표를 붙여넣으세요. 가상 예시 숫자는 채워져 있지 않습니다."
    If includeExample Then
        ws.Range("A8").Value = "예시"
        SampleDesign_WriteArray ws, 10, 1, Array("구분", "10인 미만", "10~50인 미만", "50인 이상", "합계")
        SampleDesign_WriteArray ws, 11, 1, Array("패키지SW", 42081, 4213, 82, 46376)
        SampleDesign_WriteArray ws, 12, 1, Array("IT서비스", 20775, 1826, 58, 22659)
        SampleDesign_WriteArray ws, 13, 1, Array("인터넷SW", 7509, 474, 26, 8009)
    End If
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

' ============================================================
' Stratified systematic sample extraction core
' ============================================================
Private Function SampleExtract_NormalizeMatchValue(ByVal valueData As Variant) As String
    If IsError(valueData) Or IsNull(valueData) Or Len(Trim$(CStr(valueData))) = 0 Then
        SampleExtract_NormalizeMatchValue = "(공란)"
    ElseIf IsNumeric(valueData) Then
        SampleExtract_NormalizeMatchValue = Replace$(Format$(CDbl(valueData), "0.###############"), Application.DecimalSeparator, ".")
    Else
        SampleExtract_NormalizeMatchValue = Trim$(CStr(valueData))
    End If
End Function

Private Function SampleExtract_BuildStratumKey(ByVal rowValues As Variant, ByVal columnIndexes As Variant) As String
    Dim i As Long, valueText As String, resultText As String
    For i = LBound(columnIndexes) To UBound(columnIndexes)
        valueText = SampleExtract_NormalizeMatchValue(rowValues(1, CLng(columnIndexes(i))))
        resultText = resultText & Len(valueText) & ":" & valueText & "|"
    Next i
    SampleExtract_BuildStratumKey = resultText
End Function

Private Function SampleExtract_SystematicIndexes(ByVal populationCount As Long, _
                                                 ByVal sampleCount As Long, _
                                                 ByVal randomStart As Double) As Variant
    Dim result() As Long, intervalValue As Double, i As Long
    If populationCount <= 0 Or sampleCount <= 0 Then
        SampleExtract_SystematicIndexes = Array()
        Exit Function
    End If
    If sampleCount > populationCount Then sampleCount = populationCount
    intervalValue = populationCount / sampleCount
    If randomStart < 0 Then randomStart = 0
    If randomStart >= intervalValue Then randomStart = intervalValue - 0.0000000001
    ReDim result(0 To sampleCount - 1)
    For i = 0 To sampleCount - 1
        result(i) = Int(randomStart + (i * intervalValue)) + 1
    Next i
    SampleExtract_SystematicIndexes = result
End Function

Private Function SampleExtract_AppendUniqueRound(ByVal availableRows As Collection, _
                                                 ByVal requestedCount As Long, _
                                                 ByVal randomStart As Double, _
                                                 ByVal selectedRows As Object) As Collection
    Dim result As New Collection, indexes As Variant, i As Long, rowNumber As Long
    If availableRows.Count = 0 Or requestedCount <= 0 Then
        Set SampleExtract_AppendUniqueRound = result
        Exit Function
    End If

    indexes = SampleExtract_SystematicIndexes(availableRows.Count, requestedCount, randomStart)
    For i = LBound(indexes) To UBound(indexes)
        rowNumber = CLng(availableRows(CLng(indexes(i))))
        If Not selectedRows.Exists(CStr(rowNumber)) Then
            result.Add rowNumber
            selectedRows.Add CStr(rowNumber), True
        End If
    Next i
    For i = UBound(indexes) To LBound(indexes) Step -1
        availableRows.Remove CLng(indexes(i))
    Next i
    Set SampleExtract_AppendUniqueRound = result
End Function

Public Function SampleExtract_TestSystematicPositions(ByVal populationCount As Long, _
                                                      ByVal sampleCount As Long, _
                                                      ByVal randomStart As Double) As String
    Dim indexes As Variant, i As Long, resultText As String
    indexes = SampleExtract_SystematicIndexes(populationCount, sampleCount, randomStart)
    On Error GoTo EmptyResult
    For i = LBound(indexes) To UBound(indexes)
        If Len(resultText) > 0 Then resultText = resultText & ","
        resultText = resultText & CStr(indexes(i))
    Next i
EmptyResult:
    SampleExtract_TestSystematicPositions = resultText
End Function

Public Function SampleExtract_TestErrorStats(ByVal populationCount As Long, _
                                             ByVal sampleCount As Long, _
                                             ByVal confidenceLevel As Double, _
                                             ByVal expectedProportion As Double) As String
    Dim inclusionProbability As Double, weightValue As Double
    Dim standardError As Double, marginOfError As Double, zValue As Double
    If populationCount <= 0 Or sampleCount <= 0 Then Exit Function
    If sampleCount > populationCount Then sampleCount = populationCount
    inclusionProbability = sampleCount / populationCount
    weightValue = populationCount / sampleCount
    If sampleCount < populationCount And populationCount > 1 Then
        zValue = Application.WorksheetFunction.Norm_S_Inv(0.5 + confidenceLevel / 200)
        standardError = Sqr(((populationCount - sampleCount) / (populationCount - 1)) * _
                            expectedProportion * (1 - expectedProportion) / sampleCount)
        marginOfError = zValue * standardError
    End If
    SampleExtract_TestErrorStats = "PI=" & CStr(inclusionProbability) & _
        ";WEIGHT=" & CStr(weightValue) & ";SE=" & CStr(standardError) & ";MOE=" & CStr(marginOfError)
End Function

Public Sub SampleExtract_Run()
    On Error GoTo ErrHandler
    Dim wizard As Object, actionName As String, populationRange As Range
    Set wizard = VBA.UserForms.Add("frmSampleExtractionWizard")
    Do
        wizard.Show
        actionName = LCase$(Trim$(CStr(wizard.Action)))
        Select Case actionName
            Case "pick_range"
                Set populationRange = SampleExtract_PickPopulationRange(wizard)
            Case "review"
                If populationRange Is Nothing Then
                    MsgBox "먼저 모집단 범위를 선택하세요.", vbExclamation, "표본 추출"
                    wizard.ShowPage 0
                Else
                    SampleExtract_PrepareReview wizard, populationRange
                End If
            Case "run"
                If SampleExtract_Execute(wizard, populationRange) Then Exit Do
            Case Else
                Exit Do
        End Select
    Loop
CleanExit:
    On Error Resume Next
    Unload wizard
    Exit Sub
ErrHandler:
    If Not Toolkit_ShowCancelMessage("표본 추출") Then
        MsgBox "표본 추출 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "표본 추출"
    End If
    Resume CleanExit
End Sub

Private Function SampleExtract_PickPopulationRange(ByVal wizard As Object) As Range
    Dim selectedRange As Range, headers As Variant, minimumRows As Long
    On Error Resume Next
    Set selectedRange = Application.InputBox( _
        "모집단의 헤더와 데이터를 함께 선택하세요." & vbCrLf & _
        "마우스 드래그, 주소 입력, 방향키와 Shift를 모두 사용할 수 있습니다.", _
        "표본 추출 - 모집단 범위", Type:=8)
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Function
    minimumRows = IIf(wizard.HasHeader, 2, 1)
    If selectedRange.Areas.Count > 1 Or selectedRange.Rows.Count < minimumRows Then
        MsgBox "한 시트의 연속된 범위에서 헤더와 데이터 1행 이상을 선택하세요.", vbExclamation, "표본 추출"
        Exit Function
    End If
    headers = SampleExtract_HeaderNames(selectedRange, CBool(wizard.HasHeader))
    wizard.ConfigureColumns selectedRange.Address(External:=True), Join(headers, vbTab), _
        "시트: " & selectedRange.Worksheet.Name & vbCrLf & _
        "데이터 행: " & CStr(selectedRange.Rows.Count - IIf(wizard.HasHeader, 1, 0)) & vbCrLf & _
        "열 수: " & CStr(selectedRange.Columns.Count) & vbCrLf & _
        "다음 단계 전에 ID 열과 층화 기준 열을 확인하세요."
    Set SampleExtract_PickPopulationRange = selectedRange
End Function

Private Function SampleExtract_HeaderNames(ByVal sourceRange As Range, ByVal hasHeader As Boolean) As Variant
    Dim result() As String, c As Long, textValue As String
    ReDim result(0 To sourceRange.Columns.Count - 1)
    For c = 1 To sourceRange.Columns.Count
        If hasHeader Then textValue = Trim$(CStr(sourceRange.Cells(1, c).Value2))
        If Len(textValue) = 0 Then textValue = "열" & CStr(c)
        result(c - 1) = textValue
    Next c
    SampleExtract_HeaderNames = result
End Function

Private Function SampleExtract_ParseIndexes(ByVal csvText As String) As Variant
    Dim parts As Variant, result() As Long, i As Long
    parts = Split(csvText, ",")
    ReDim result(LBound(parts) To UBound(parts))
    For i = LBound(parts) To UBound(parts): result(i) = CLng(parts(i)): Next i
    SampleExtract_ParseIndexes = result
End Function

Private Function SampleExtract_KeyFromData(ByVal values As Variant, ByVal rowIndex As Long, ByVal indexes As Variant) As String
    Dim i As Long, valueText As String, resultText As String
    For i = LBound(indexes) To UBound(indexes)
        valueText = SampleExtract_NormalizeMatchValue(values(rowIndex, CLng(indexes(i))))
        resultText = resultText & Len(valueText) & ":" & valueText & "|"
    Next i
    SampleExtract_KeyFromData = resultText
End Function

Private Function SampleExtract_BuildStrata(ByVal values As Variant, ByVal firstDataRow As Long, ByVal indexes As Variant) As Object
    Dim result As Object, rows As Collection, r As Long, keyText As String
    Set result = CreateObject("Scripting.Dictionary")
    For r = firstDataRow To UBound(values, 1)
        keyText = SampleExtract_KeyFromData(values, r, indexes)
        If Not result.Exists(keyText) Then
            Set rows = New Collection
            result.Add keyText, rows
        End If
        Set rows = result(keyText)
        rows.Add r
    Next r
    Set SampleExtract_BuildStrata = result
End Function

Private Function SampleExtract_NormalizeHeader(ByVal headerText As String) As String
    SampleExtract_NormalizeHeader = LCase$(Replace$(Replace$(Replace$(Trim$(headerText), " ", ""), vbCr, ""), vbLf, ""))
End Function

Private Function SampleExtract_LoadAllocationTargets(ByVal wb As Workbook, ByVal headers As Variant, _
                                                     ByVal stratumIndexes As Variant, ByRef warningText As String) As Object
    Dim result As Object, ws As Worksheet, headerMap As Object, c As Long, r As Long, lastRow As Long
    Dim finalCol As Long, allocationIndexes() As Long, i As Long, keyText As String, valueText As String
    Set result = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = wb.Worksheets("표본설계_할당결과")
    On Error GoTo 0
    If ws Is Nothing Then
        warningText = "표본설계_할당결과 시트를 찾지 못했습니다. 층별 목표를 직접 입력하세요."
        Set SampleExtract_LoadAllocationTargets = result
        Exit Function
    End If
    Set headerMap = CreateObject("Scripting.Dictionary")
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headerMap(SampleExtract_NormalizeHeader(CStr(ws.Cells(1, c).Value2))) = c
    Next c
    If Not headerMap.Exists(SampleExtract_NormalizeHeader("최종할당")) Then
        warningText = "할당결과에서 최종할당 열을 찾지 못했습니다."
        Set SampleExtract_LoadAllocationTargets = result
        Exit Function
    End If
    finalCol = CLng(headerMap(SampleExtract_NormalizeHeader("최종할당")))
    ReDim allocationIndexes(LBound(stratumIndexes) To UBound(stratumIndexes))
    For i = LBound(stratumIndexes) To UBound(stratumIndexes)
        valueText = SampleExtract_NormalizeHeader(CStr(headers(CLng(stratumIndexes(i)) - 1)))
        If Not headerMap.Exists(valueText) Then
            warningText = "할당결과에서 층화 열 '" & CStr(headers(CLng(stratumIndexes(i)) - 1)) & "'을 찾지 못했습니다."
            Set SampleExtract_LoadAllocationTargets = result
            Exit Function
        End If
        allocationIndexes(i) = CLng(headerMap(valueText))
    Next i
    lastRow = ws.Cells(ws.Rows.Count, finalCol).End(xlUp).Row
    For r = 2 To lastRow
        keyText = ""
        For i = LBound(allocationIndexes) To UBound(allocationIndexes)
            valueText = SampleExtract_NormalizeMatchValue(ws.Cells(r, allocationIndexes(i)).Value2)
            keyText = keyText & Len(valueText) & ":" & valueText & "|"
        Next i
        If IsNumeric(ws.Cells(r, finalCol).Value2) Then
            If result.Exists(keyText) Then
                result(keyText) = CLng(result(keyText)) + CLng(ws.Cells(r, finalCol).Value2)
            Else
                result.Add keyText, CLng(ws.Cells(r, finalCol).Value2)
            End If
        End If
    Next r
    Set SampleExtract_LoadAllocationTargets = result
End Function

Private Function SampleExtract_BuildReviewText(ByVal strata As Object, ByVal targets As Object, _
                                               ByVal replacementMultiplier As Long, ByRef targetText As String, _
                                               ByRef warningText As String) As String
    Dim key As Variant, rows As Collection, targetValue As Variant, resultText As String
    resultText = "층 키" & vbTab & "모집단" & vbTab & "목표 원표본" & vbTab & "요청 대체" & vbTab & "상태"
    For Each key In strata.Keys
        Set rows = strata(key)
        If targets.Exists(CStr(key)) Then targetValue = targets(CStr(key)) Else targetValue = ""
        targetText = targetText & IIf(Len(targetText) > 0, vbCrLf, "") & CStr(key) & vbTab & CStr(targetValue)
        resultText = resultText & vbCrLf & CStr(key) & vbTab & rows.Count & vbTab & CStr(targetValue) & vbTab
        If Len(CStr(targetValue)) = 0 Then
            resultText = resultText & "-" & vbTab & "확인 필요"
            warningText = warningText & IIf(Len(warningText) > 0, vbCrLf, "") & "목표표본이 없는 층: " & CStr(key)
        Else
            resultText = resultText & CLng(targetValue) * replacementMultiplier & vbTab & _
                IIf(CLng(targetValue) > rows.Count, "모집단 부족", "실행 가능")
        End If
    Next key
    SampleExtract_BuildReviewText = resultText
End Function

Private Sub SampleExtract_PrepareReview(ByVal wizard As Object, ByVal populationRange As Range)
    Dim values As Variant, headers As Variant, indexes As Variant, strata As Object, targets As Object
    Dim warningText As String, targetText As String, reviewText As String, firstDataRow As Long
    values = populationRange.Value2
    headers = SampleExtract_HeaderNames(populationRange, CBool(wizard.HasHeader))
    indexes = SampleExtract_ParseIndexes(CStr(wizard.StratumColumnIndexes))
    firstDataRow = IIf(wizard.HasHeader, 2, 1)
    Set strata = SampleExtract_BuildStrata(values, firstDataRow, indexes)
    Set targets = CreateObject("Scripting.Dictionary")
    If CStr(wizard.TargetMode) = "auto" Then
        Set targets = SampleExtract_LoadAllocationTargets(populationRange.Worksheet.Parent, headers, indexes, warningText)
    End If
    reviewText = SampleExtract_BuildReviewText(strata, targets, CLng(wizard.ReplacementMultiplier), targetText, warningText)
    wizard.ConfigureReview reviewText, targetText, warningText
End Sub

Private Function SampleExtract_ParseTargetEdits(ByVal targetText As String, ByVal strata As Object, ByRef warningText As String) As Object
    Dim result As Object, lines As Variant, fields As Variant, i As Long, keyText As String, rawTarget As String
    Set result = CreateObject("Scripting.Dictionary")
    lines = Split(Replace$(targetText, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) > 0 Then
            fields = Split(CStr(lines(i)), vbTab)
            If UBound(fields) < 1 Then warningText = "목표 목록은 층 키와 목표값을 Tab으로 구분해야 합니다.": Exit For
            keyText = CStr(fields(0)): rawTarget = Trim$(CStr(fields(1)))
            If Not strata.Exists(keyText) Or Not IsNumeric(rawTarget) Or CDbl(rawTarget) < 0 Or CDbl(rawTarget) <> Fix(CDbl(rawTarget)) Then
                warningText = "층 키 또는 목표값을 확인하세요: " & CStr(lines(i)): Exit For
            End If
            result(keyText) = CLng(rawTarget)
        End If
    Next i
    If Len(warningText) = 0 Then
        Dim key As Variant
        For Each key In strata.Keys
            If Not result.Exists(CStr(key)) Then warningText = "모든 층의 목표 원표본을 입력하세요.": Exit For
        Next key
    End If
    Set SampleExtract_ParseTargetEdits = result
End Function

Private Function SampleExtract_CompareRows(ByVal values As Variant, ByVal rowA As Long, ByVal rowB As Long, ByVal sortColumn As Long) As Long
    If IsNumeric(values(rowA, sortColumn)) And IsNumeric(values(rowB, sortColumn)) Then
        SampleExtract_CompareRows = Sgn(CDbl(values(rowA, sortColumn)) - CDbl(values(rowB, sortColumn)))
    Else
        SampleExtract_CompareRows = StrComp(CStr(values(rowA, sortColumn)), CStr(values(rowB, sortColumn)), vbTextCompare)
    End If
End Function

Private Sub SampleExtract_SortRows(ByVal rows As Collection, ByVal values As Variant, ByVal sortColumn As Long, ByVal descending As Boolean)
    Dim orderedRows() As Long, i As Long
    If rows.Count < 2 Then Exit Sub
    ReDim orderedRows(1 To rows.Count)
    For i = 1 To rows.Count: orderedRows(i) = CLng(rows(i)): Next i
    SampleExtract_QuickSortRows orderedRows, LBound(orderedRows), UBound(orderedRows), values, sortColumn, descending
    Do While rows.Count > 0: rows.Remove 1: Loop
    For i = LBound(orderedRows) To UBound(orderedRows): rows.Add orderedRows(i): Next i
End Sub

Private Function SampleExtract_CopyRows(ByVal sourceRows As Collection) As Collection
    Dim result As New Collection, i As Long
    For i = 1 To sourceRows.Count: result.Add CLng(sourceRows(i)): Next i
    Set SampleExtract_CopyRows = result
End Function

Private Sub SampleExtract_QuickSortRows(ByRef orderedRows() As Long, ByVal low As Long, ByVal high As Long, _
                                        ByVal values As Variant, ByVal sortColumn As Long, ByVal descending As Boolean)
    Dim i As Long, j As Long, pivotRow As Long, temp As Long, direction As Long
    i = low: j = high: pivotRow = orderedRows((low + high) \ 2)
    direction = IIf(descending, -1, 1)
    Do While i <= j
        Do While SampleExtract_CompareRows(values, orderedRows(i), pivotRow, sortColumn) * direction < 0: i = i + 1: Loop
        Do While SampleExtract_CompareRows(values, orderedRows(j), pivotRow, sortColumn) * direction > 0: j = j - 1: Loop
        If i <= j Then
            temp = orderedRows(i): orderedRows(i) = orderedRows(j): orderedRows(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then SampleExtract_QuickSortRows orderedRows, low, j, values, sortColumn, descending
    If i < high Then SampleExtract_QuickSortRows orderedRows, i, high, values, sortColumn, descending
End Sub

Private Function SampleExtract_Execute(ByVal wizard As Object, ByVal populationRange As Range) As Boolean
    On Error GoTo ErrHandler
    Dim wb As Workbook, outputName As String, reviewName As String, warningText As String
    If populationRange Is Nothing Then MsgBox "모집단 범위를 다시 선택하세요.", vbExclamation, "표본 추출": Exit Function
    Set wb = populationRange.Worksheet.Parent
    outputName = CStr(wizard.OutputSheetName): reviewName = "표본추출_검토"
    If SampleExtract_SheetExists(wb, outputName) Or SampleExtract_SheetExists(wb, reviewName) Then
        MsgBox "결과 또는 검토 시트가 이미 있습니다. 기존 시트명을 변경하거나 삭제한 뒤 다시 실행하세요.", vbExclamation, "표본 추출"
        Exit Function
    End If

    Dim values As Variant, headers As Variant, indexes As Variant, strata As Object, availableByKey As Object, targets As Object
    Dim firstDataRow As Long, key As Variant, rows As Collection
    values = populationRange.Value2: headers = SampleExtract_HeaderNames(populationRange, CBool(wizard.HasHeader))
    indexes = SampleExtract_ParseIndexes(CStr(wizard.StratumColumnIndexes))
    firstDataRow = IIf(wizard.HasHeader, 2, 1)
    Set strata = SampleExtract_BuildStrata(values, firstDataRow, indexes)
    Set targets = SampleExtract_ParseTargetEdits(CStr(wizard.TargetEditsText), strata, warningText)
    If Len(warningText) > 0 Then MsgBox warningText, vbExclamation, "표본 추출": Exit Function

    Dim sortMode As String, sortColumn As Long
    sortMode = CStr(wizard.SortMode): sortColumn = CLng(wizard.SortColumnIndex)
    If sortMode <> "keep" Then
        For Each key In strata.Keys
            Set rows = strata(key): SampleExtract_SortRows rows, values, sortColumn, (sortMode = "desc")
        Next key
    End If
    Set availableByKey = CreateObject("Scripting.Dictionary")
    For Each key In strata.Keys
        Set rows = SampleExtract_CopyRows(strata(key))
        availableByKey.Add CStr(key), rows
    Next key

    Dim seedValue As Long
    If Len(CStr(wizard.RandomSeedText)) > 0 And IsNumeric(wizard.RandomSeedText) Then seedValue = CLng(wizard.RandomSeedText) Else seedValue = CLng(Timer * 100)
    Randomize seedValue
    Dim records As New Collection, selectedRows As Object, statsByKey As Object, roundNo As Long, totalSteps As Long, stepNo As Long
    Dim targetN As Long, actualN As Long, intervalValue As Double, startValue As Double, picks As Collection, i As Long
    Dim statsText As String, statParts As Variant, replacementMultiplier As Long
    Set selectedRows = CreateObject("Scripting.Dictionary")
    Set statsByKey = CreateObject("Scripting.Dictionary")
    replacementMultiplier = CLng(wizard.ReplacementMultiplier)
    For Each key In strata.Keys
        Set rows = strata(key): targetN = CLng(targets(CStr(key))): actualN = targetN
        If actualN > rows.Count Then actualN = rows.Count
        If actualN > 0 Then statsByKey(CStr(key)) = SampleExtract_TestErrorStats(rows.Count, actualN, CDbl(wizard.ConfidenceLevel), CDbl(wizard.ExpectedProportion))
    Next key
    totalSteps = strata.Count * (replacementMultiplier + 1)
    Toolkit_ProgressShow "표본 추출", "층별 원표본과 대체표본을 추출합니다.", totalSteps
    For roundNo = 0 To replacementMultiplier
        For Each key In strata.Keys
            Set rows = availableByKey(CStr(key)): targetN = CLng(targets(CStr(key)))
            actualN = targetN: If actualN > rows.Count Then actualN = rows.Count
            If actualN > 0 Then
                intervalValue = rows.Count / actualN: startValue = Rnd * intervalValue
                statsText = CStr(statsByKey(CStr(key)))
                Set picks = SampleExtract_AppendUniqueRound(rows, actualN, startValue, selectedRows)
                For i = 1 To picks.Count
                    records.Add Array(CLng(picks(i)), CStr(key), IIf(roundNo = 0, "원표본", "대체표본"), roundNo, i, i, intervalValue, startValue, statsText)
                Next i
            End If
            stepNo = stepNo + 1
            Toolkit_ProgressUpdate stepNo, totalSteps, "차수 " & roundNo & " / 층 " & CStr(key)
        Next key
    Next roundNo
    SampleExtract_WriteResultSheet wb, outputName, populationRange, values, headers, records, CLng(wizard.IdColumnIndex), seedValue, CDbl(wizard.ConfidenceLevel), CDbl(wizard.ExpectedProportion)
    SampleExtract_WriteReviewSheet wb, reviewName, wizard, strata, targets, records, seedValue
    Toolkit_ProgressClose
    MsgBox "표본추출_결과와 표본추출_검토 시트를 만들었습니다.", vbInformation, "표본 추출"
    SampleExtract_Execute = True
    Exit Function
ErrHandler:
    Toolkit_ProgressClose
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb Is Nothing Then
        If SampleExtract_SheetExists(wb, outputName) Then wb.Worksheets(outputName).Delete
        If SampleExtract_SheetExists(wb, reviewName) Then wb.Worksheets(reviewName).Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
    If Toolkit_IsCancelError(Err.Number) Then
        MsgBox "표본 추출을 취소했습니다. 원본 시트는 변경되지 않았습니다.", vbInformation, "표본 추출"
    Else
        MsgBox "표본 추출 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "표본 추출"
    End If
End Function

Private Function SampleExtract_SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Worksheets(sheetName): On Error GoTo 0
    SampleExtract_SheetExists = Not ws Is Nothing
End Function

Private Sub SampleExtract_WriteResultSheet(ByVal wb As Workbook, ByVal outputName As String, ByVal sourceRange As Range, _
                                           ByVal values As Variant, ByVal headers As Variant, ByVal records As Collection, _
                                           ByVal idColumn As Long, ByVal seedValue As Long, ByVal confidenceLevel As Double, _
                                           ByVal expectedProportion As Double)
    Dim ws As Worksheet, meta As Variant, output() As Variant, r As Long, c As Long, rec As Variant, parts As Variant
    meta = Array("__SAMPLE_SOURCE_SHEET", "__SAMPLE_SOURCE_ROW", "__SAMPLE_ID", "__SAMPLE_STRATUM_ID", "__SAMPLE_STRATUM_KEY", _
        "__SAMPLE_TYPE", "__SAMPLE_REPLACEMENT_ROUND", "__SAMPLE_LINKED_ORIGINAL_ORDER", "__SAMPLE_WITHIN_STRATUM_ORDER", _
        "__SAMPLE_INTERVAL", "__SAMPLE_RANDOM_START", "__SAMPLE_INCLUSION_PROBABILITY", "__SAMPLE_WEIGHT", _
        "__SAMPLE_STANDARD_ERROR", "__SAMPLE_MARGIN_OF_ERROR", "__SAMPLE_CONFIDENCE_LEVEL", "__SAMPLE_EXPECTED_PROPORTION", "__SAMPLE_RANDOM_SEED")
    ReDim output(1 To records.Count + 1, 1 To UBound(values, 2) + UBound(meta) + 1)
    For c = 1 To UBound(values, 2): output(1, c) = headers(c - 1): Next c
    For c = 0 To UBound(meta): output(1, UBound(values, 2) + c + 1) = meta(c): Next c
    For r = 1 To records.Count
        rec = records(r)
        For c = 1 To UBound(values, 2): output(r + 1, c) = values(rec(0), c): Next c
        c = UBound(values, 2)
        output(r + 1, c + 1) = sourceRange.Worksheet.Name
        output(r + 1, c + 2) = sourceRange.Row + CLng(rec(0)) - 1
        If idColumn > 0 And Len(Trim$(CStr(values(rec(0), idColumn)))) > 0 Then output(r + 1, c + 3) = values(rec(0), idColumn) Else output(r + 1, c + 3) = sourceRange.Worksheet.Name & "_R" & output(r + 1, c + 2)
        output(r + 1, c + 4) = rec(1): output(r + 1, c + 5) = rec(1): output(r + 1, c + 6) = rec(2)
        output(r + 1, c + 7) = rec(3): output(r + 1, c + 8) = rec(4): output(r + 1, c + 9) = rec(5)
        output(r + 1, c + 10) = rec(6): output(r + 1, c + 11) = rec(7)
        parts = Split(CStr(rec(8)), ";")
        output(r + 1, c + 12) = CDbl(Split(parts(0), "=")(1)): output(r + 1, c + 13) = CDbl(Split(parts(1), "=")(1))
        output(r + 1, c + 14) = CDbl(Split(parts(2), "=")(1)): output(r + 1, c + 15) = CDbl(Split(parts(3), "=")(1))
        output(r + 1, c + 16) = confidenceLevel: output(r + 1, c + 17) = expectedProportion: output(r + 1, c + 18) = seedValue
    Next r
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = outputName
    ws.Cells(1, 1).Resize(UBound(output, 1), UBound(output, 2)).Value = output
    ws.Rows(1).Font.Bold = True: ws.Rows(1).AutoFilter: ws.Columns.AutoFit
End Sub

Private Sub SampleExtract_WriteReviewSheet(ByVal wb As Workbook, ByVal reviewName As String, ByVal wizard As Object, _
                                           ByVal strata As Object, ByVal targets As Object, ByVal records As Collection, ByVal seedValue As Long)
    Dim ws As Worksheet, rowNo As Long, key As Variant, rows As Collection, rec As Variant, i As Long
    Dim originalCounts As Object, replacementCounts As Object
    Set originalCounts = CreateObject("Scripting.Dictionary")
    Set replacementCounts = CreateObject("Scripting.Dictionary")
    For i = 1 To records.Count
        rec = records(i): key = CStr(rec(1))
        If CLng(rec(3)) = 0 Then
            If originalCounts.Exists(key) Then originalCounts(key) = CLng(originalCounts(key)) + 1 Else originalCounts.Add key, 1
        Else
            If replacementCounts.Exists(key) Then replacementCounts(key) = CLng(replacementCounts(key)) + 1 Else replacementCounts.Add key, 1
        End If
    Next i
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = reviewName
    ws.Range("A1").Value = "표본 추출 검토": ws.Range("A1").Font.Bold = True
    ws.Cells(3, 1).Value = "모집단 범위": ws.Cells(3, 2).Value = wizard.PopulationAddress
    ws.Cells(4, 1).Value = "층화 열": ws.Cells(4, 2).Value = wizard.StratumColumnIndexes
    ws.Cells(5, 1).Value = "정렬": ws.Cells(5, 2).Value = wizard.SortMode
    ws.Cells(6, 1).Value = "대체표본 배수": ws.Cells(6, 2).Value = wizard.ReplacementMultiplier
    ws.Cells(7, 1).Value = "신뢰수준": ws.Cells(7, 2).Value = wizard.ConfidenceLevel
    ws.Cells(8, 1).Value = "예상비율": ws.Cells(8, 2).Value = wizard.ExpectedProportion
    ws.Cells(9, 1).Value = "난수 시드": ws.Cells(9, 2).Value = seedValue
    ws.Range("A11").Value = "표본오차는 단순임의추출 근사치이며, 정렬된 명부의 실제 계통추출 분산과 다를 수 있습니다."
    SampleDesign_WriteArray ws, 13, 1, Array("층 키", "모집단수", "목표 원표본", "실제 원표본", "요청 대체표본", "실제 대체표본")
    rowNo = 14
    For Each key In strata.Keys
        Set rows = strata(key): ws.Cells(rowNo, 1).Value = CStr(key): ws.Cells(rowNo, 2).Value = rows.Count
        ws.Cells(rowNo, 3).Value = targets(CStr(key))
        If originalCounts.Exists(CStr(key)) Then ws.Cells(rowNo, 4).Value = originalCounts(CStr(key)) Else ws.Cells(rowNo, 4).Value = 0
        ws.Cells(rowNo, 5).Value = CLng(targets(CStr(key))) * CLng(wizard.ReplacementMultiplier)
        If replacementCounts.Exists(CStr(key)) Then ws.Cells(rowNo, 6).Value = replacementCounts(CStr(key)) Else ws.Cells(rowNo, 6).Value = 0
        rowNo = rowNo + 1
    Next key
    ws.Columns.AutoFit
End Sub
