Attribute VB_Name = "DatabaseTools"
Option Explicit


Public Sub Database_MergeSheetsByHeaders()
    On Error GoTo ErrHandler
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

    Dim form As Object
    Set form = VBA.UserForms.Add("frmDatabaseMergeSetup")
    form.Configure DBMerge_WizardSheetInfoText(wb), DBMerge_WizardDefaultSheetNames(wb), _
                   keyHeader, outputSheetName, outputPosition, headerRowSetting, headerMode, _
                   idMode, keyMode, useMarkerRow

    Dim sheets As Collection
    Dim manualKeyColumns As Object
    Dim outputHeaders As Collection
    Dim plan As Object

    Do
        form.Show

        Select Case LCase$(CStr(form.Action))
            Case "headers"
                selectedSheetNames = CStr(form.SelectedSheets)
                headerRowSetting = CStr(form.HeaderRowSetting)
                headerMode = CStr(form.HeaderMode)
                form.ConfigureHeaders DBMerge_WizardHeaderInfoText(wb, selectedSheetNames, headerRowSetting, headerMode)

            Case "preview"
                Set plan = Nothing
                selectedSheetNames = CStr(form.SelectedSheets)
                keyHeader = Trim$(CStr(form.KeyHeader))
                outputHeadersCsv = CStr(form.SelectedHeaders)
                outputSheetName = DBMerge_CleanSheetName(CStr(form.OutputSheetName), "통합데이터")
                outputPosition = CStr(form.OutputPosition)
                headerRowSetting = CStr(form.HeaderRowSetting)
                headerMode = CStr(form.HeaderMode)
                idMode = CStr(form.IdMode)
                keyMode = DBMerge_NormalizeKeyMode(CStr(form.KeyMode))
                useMarkerRow = CBool(form.UseMarkerRow)

                If DBMerge_NameIsSelected(selectedSheetNames, outputSheetName) Then
                    MsgBox "결과 시트 이름이 통합 대상 시트와 같습니다." & vbCrLf & _
                           "원본 시트를 덮어쓰지 않도록 다른 결과 시트명을 입력하세요.", vbExclamation, "시트 데이터 통합"
                    form.ConfigureHeaders DBMerge_WizardHeaderInfoText(wb, selectedSheetNames, headerRowSetting, headerMode)
                Else
                    Set sheets = DBMerge_SelectedWorksheets(wb, selectedSheetNames)
                    Set manualKeyColumns = Nothing
                    If keyMode = "manual" Then
                        Set manualKeyColumns = DBMerge_WizardManualKeyColumns(wb, CStr(form.ManualKeyHeaders), headerRowSetting, headerMode)
                        If manualKeyColumns.Count <> sheets.Count Then
                            MsgBox "선택한 모든 시트의 ID 열을 다시 지정하세요.", vbExclamation, "시트 데이터 통합"
                            form.ConfigureHeaders DBMerge_WizardHeaderInfoText(wb, selectedSheetNames, headerRowSetting, headerMode)
                            GoTo ContinueWizard
                        End If
                    End If

                    Set outputHeaders = DBMerge_ParseHeaderList(outputHeadersCsv, keyHeader)
                    Set plan = DBMerge_BuildPlan(sheets, keyHeader, outputHeaders, headerRowSetting, headerMode, idMode, keyMode, useMarkerRow, manualKeyColumns)
                    If plan Is Nothing Then
                        form.ConfigureHeaders DBMerge_WizardHeaderInfoText(wb, selectedSheetNames, headerRowSetting, headerMode)
                    Else
                        form.ConfigureReview DBMerge_PlanSummary(plan), DBMerge_MappingSummary(plan), DBMerge_WizardWarningSummary(plan)
                    End If
                End If

            Case "run"
                If Not plan Is Nothing Then Exit Do

            Case Else
                Unload form
                Exit Sub
        End Select
ContinueWizard:
    Loop

    outputSheetName = DBMerge_CleanSheetName(CStr(form.OutputSheetName), "통합데이터")
    outputPosition = CStr(form.OutputPosition)
    Unload form
    Set form = Nothing

    Dim outWs As Worksheet
    Dim existingOutWs As Worksheet
    Dim sourceIdRollback As Collection
    Dim outputCommitted As Boolean
    Set outWs = DBMerge_PrepareTemporaryOutputSheet(wb, outputSheetName, outputPosition, existingOutWs)
    If outWs Is Nothing Then Exit Sub

    Dim writtenRows As Long
    writtenRows = DBMerge_WriteMergedSheet(plan, outWs)
    If CStr(plan("idMode")) = "source_auto" Then
        Set sourceIdRollback = DBMerge_CreateSourceIdRollbackPlan(plan)
        DBMerge_ApplySourceIds plan, sourceIdRollback
    End If
    DBMerge_CommitTemporaryOutputSheet outWs, existingOutWs, outputSheetName
    outputCommitted = True
    Set outWs = wb.Worksheets(outputSheetName)
    outWs.Activate
    MsgBox sheets.Count & "개 시트를 통합했습니다." & vbCrLf & _
           "생성 행 수: " & writtenRows & vbCrLf & _
           "결과 시트: " & outWs.Name, vbInformation, "시트 데이터 통합"
    Exit Sub

ErrHandler:
    Dim errorText As String
    errorText = Err.Description
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    If Not outputCommitted Then
        DBMerge_RollbackSourceIds sourceIdRollback
        DBMerge_DeleteWorksheetQuietly outWs
    End If
    On Error GoTo 0
    MsgBox "시트 통합 중 오류가 발생했습니다." & vbCrLf & errorText, vbExclamation, "시트 데이터 통합"
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

Public Function DBMerge_WizardSheetInfoText(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim result As String
    For Each ws In wb.Worksheets
        Dim headerRow As Long
        Dim dataCount As Long
        Dim stateText As String
        Dim warningText As String

        dataCount = 0
        stateText = vbNullString
        warningText = vbNullString

        headerRow = DBMerge_DetectHeaderRow(ws)
        If headerRow = 0 Then
            stateText = "확인 필요"
            warningText = "헤더를 찾지 못함"
        Else
            dataCount = DBMerge_LastUsedRow(ws) - headerRow
            If dataCount < 0 Then dataCount = 0
            If DBMerge_WizardIsDefaultSheet(ws) Then
                stateText = "선택"
            Else
                stateText = "제외"
            End If
            If DBMerge_WizardDuplicateHeaderCount(ws, headerRow, "single") > 0 Then warningText = "중복 헤더"
        End If

        If Len(result) > 0 Then result = result & vbLf
        result = result & ws.Name & vbTab & CStr(headerRow) & vbTab & CStr(dataCount) & vbTab & stateText
        If Len(warningText) > 0 Then result = result & vbTab & warningText
    Next ws
    DBMerge_WizardSheetInfoText = result
End Function

Public Function DBMerge_WizardDefaultSheetNames(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If DBMerge_WizardIsDefaultSheet(ws) Then
            If Len(DBMerge_WizardDefaultSheetNames) > 0 Then DBMerge_WizardDefaultSheetNames = DBMerge_WizardDefaultSheetNames & vbLf
            DBMerge_WizardDefaultSheetNames = DBMerge_WizardDefaultSheetNames & ws.Name
        End If
    Next ws
End Function

Public Function DBMerge_WizardHeaderInfoText(ByVal wb As Workbook, _
                                              ByVal selectedSheetNames As String, _
                                              ByVal headerRowSetting As String, _
                                              ByVal headerMode As String) As String
    Dim displayNames As Object
    Dim matchCounts As Object
    Dim headerOrder As Collection
    Set displayNames = CreateObject("Scripting.Dictionary")
    Set matchCounts = CreateObject("Scripting.Dictionary")
    Set headerOrder = New Collection

    Dim selectedCount As Long
    selectedCount = DBMerge_WizardSelectedSheetCount(wb, selectedSheetNames)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If DBMerge_NameIsSelected(selectedSheetNames, ws.Name) Then
            Dim headerRow As Long
            headerRow = DBMerge_HeaderRowForSheet(ws, headerRowSetting, headerMode, False)
            If headerRow > 0 Then
                Dim seenOnSheet As Object
                Set seenOnSheet = CreateObject("Scripting.Dictionary")

                Dim c As Long
                For c = 1 To DBMerge_LastUsedCol(ws)
                    Dim headerText As String
                    Dim norm As String
                    headerText = DBMerge_HeaderTextForColumn(ws, c, headerRow, headerMode)
                    norm = DBMerge_NormalizeHeader(headerText)
                    If Len(norm) > 0 And Not seenOnSheet.Exists(norm) Then
                        seenOnSheet.Add norm, True
                        If Not displayNames.Exists(norm) Then
                            displayNames.Add norm, headerText
                            matchCounts.Add norm, 0
                            headerOrder.Add norm
                        End If
                        matchCounts(norm) = CLng(matchCounts(norm)) + 1
                    End If
                Next c
            End If
        End If
    Next ws

    Dim i As Long
    Dim result As String
    For i = 1 To headerOrder.Count
        Dim key As String
        Dim stateText As String
        key = CStr(headerOrder(i))
        If selectedCount > 0 And CLng(matchCounts(key)) = selectedCount Then
            stateText = "전체 시트"
        Else
            stateText = "일부 시트만 있음"
        End If
        If Len(result) > 0 Then result = result & vbLf
        result = result & CStr(displayNames(key)) & vbTab & CStr(matchCounts(key)) & vbTab & stateText
    Next i
    DBMerge_WizardHeaderInfoText = result
End Function

Private Function DBMerge_WizardDuplicateHeaderCount(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                                     ByVal headerMode As String) As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim c As Long
    For c = 1 To DBMerge_LastUsedCol(ws)
        Dim norm As String
        norm = DBMerge_NormalizeHeader(DBMerge_HeaderTextForColumn(ws, c, headerRow, headerMode))
        If Len(norm) > 0 Then
            If seen.Exists(norm) Then
                DBMerge_WizardDuplicateHeaderCount = DBMerge_WizardDuplicateHeaderCount + 1
            Else
                seen.Add norm, True
            End If
        End If
    Next c
End Function

Private Function DBMerge_WizardIsDefaultSheet(ByVal ws As Worksheet) As Boolean
    If ws.Visible <> xlSheetVisible Then Exit Function
    If Left$(ws.Name, 1) = "_" Then Exit Function

    Dim c As Long
    For c = 1 To DBMerge_LastUsedCol(ws)
        If CStr(ws.Cells(1, c).Value) = "__DBMERGE_SOURCE_SHEET" Then Exit Function
    Next c
    DBMerge_WizardIsDefaultSheet = True
End Function

Private Function DBMerge_WizardSelectedSheetCount(ByVal wb As Workbook, ByVal selectedSheetNames As String) As Long
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If DBMerge_NameIsSelected(selectedSheetNames, ws.Name) Then
            DBMerge_WizardSelectedSheetCount = DBMerge_WizardSelectedSheetCount + 1
        End If
    Next ws
End Function

Public Function DBMerge_WizardManualKeyColumns(ByVal wb As Workbook, ByVal mappingText As String, _
                                                ByVal headerRowSetting As String, ByVal headerMode As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim lines As Variant
    lines = Split(Replace(mappingText, vbCr, vbLf), vbLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim parts As Variant
        parts = Split(CStr(lines(i)), vbTab)
        If UBound(parts) >= 1 Then
            Dim sheetName As String
            Dim headerText As String
            sheetName = Trim$(CStr(parts(0)))
            headerText = Trim$(CStr(parts(1)))
            If Len(sheetName) > 0 And Len(headerText) > 0 Then
                Dim ws As Worksheet
                Set ws = Nothing
                On Error Resume Next
                Set ws = wb.Worksheets(sheetName)
                On Error GoTo 0
                If Not ws Is Nothing Then
                    Dim headerRow As Long
                    headerRow = DBMerge_HeaderRowForSheet(ws, headerRowSetting, headerMode, False)
                    If headerRow > 0 Then
                        Dim headerMap As Object
                        Set headerMap = DBMerge_HeaderMap(ws, headerRow, headerMode)
                        Dim norm As String
                        norm = DBMerge_NormalizeHeader(headerText)
                        If headerMap.Exists(norm) Then result(sheetName) = CLng(headerMap(norm))
                    End If
                End If
            End If
        End If
    Next i
    Set DBMerge_WizardManualKeyColumns = result
End Function

Public Function DBMerge_WizardWarningSummary(ByVal plan As Object) As String
    Dim result As String
    Dim h As Long
    For h = 1 To plan("headers").Count
        Dim headerText As String
        Dim missingSheets As String
        headerText = CStr(plan("headers")(h))

        Dim i As Long
        For i = 1 To plan("sheets").Count
            Dim sheetPlan As Object
            Set sheetPlan = plan("sheets")(i)
            If Not sheetPlan("headerMap").Exists(DBMerge_NormalizeHeader(headerText)) Then
                If Len(missingSheets) > 0 Then missingSheets = missingSheets & ", "
                missingSheets = missingSheets & CStr(sheetPlan("sheetName"))
            End If
        Next i

        If Len(missingSheets) > 0 Then
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & "일부 누락: " & headerText & " (" & missingSheets & ") - 결과에는 빈값으로 생성됩니다."
        End If
    Next h

    If CStr(plan("idMode")) = "source_auto" Then
        If Len(result) > 0 Then result = result & vbCrLf
        result = result & "원본 변경: 자동 ID가 원본 시트의 새 열 또는 빈 ID 셀에 기록됩니다."
    End If
    If Len(result) = 0 Then result = "정상: 확인이 필요한 누락 또는 원본 변경이 없습니다."
    DBMerge_WizardWarningSummary = result
End Function

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

    Dim dataHeaders As Collection
    Set dataHeaders = DBMerge_DataHeaders(plan("headers"), CStr(plan("keyHeader")))

    Dim headerData() As Variant, c As Long
    ReDim headerData(1 To 2, 1 To dataHeaders.Count + 3)
    headerData(1, 1) = "__DBMERGE_SOURCE_SHEET"
    headerData(1, 2) = "__DBMERGE_SOURCE_ROW"
    headerData(1, 3) = "__DBMERGE_KEY"
    headerData(2, 1) = "원본시트"
    headerData(2, 2) = "원본행"
    headerData(2, 3) = CStr(plan("keyHeader"))
    For c = 1 To dataHeaders.Count
        headerData(1, c + 3) = "__FIELD:" & CStr(dataHeaders(c))
        headerData(2, c + 3) = CStr(dataHeaders(c))
    Next c
    outWs.Cells(1, 1).Resize(2, UBound(headerData, 2)).Value = headerData

    Dim totalRows As Long, processedRows As Long, writtenRows As Long
    Dim i As Long, sheetPlan As Object
    For i = 1 To plan("sheets").Count
        Set sheetPlan = plan("sheets")(i)
        totalRows = totalRows + CLng(sheetPlan("lastRow")) - CLng(sheetPlan("dataStartRow")) + 1
    Next i
    Toolkit_ProgressShow "시트 데이터 통합", "원본 데이터를 배치 단위로 읽는 중", totalRows

    Dim batchStart As Long, batchEnd As Long, batchCount As Long, outputBatch As Variant
    For i = 1 To plan("sheets").Count
        Set sheetPlan = plan("sheets")(i)
        batchStart = CLng(sheetPlan("dataStartRow"))
        Do While batchStart <= CLng(sheetPlan("lastRow"))
            batchEnd = batchStart + DBMERGE_OUTPUT_BATCH_ROWS - 1
            If batchEnd > CLng(sheetPlan("lastRow")) Then batchEnd = CLng(sheetPlan("lastRow"))
            outputBatch = DBMerge_BuildOutputBatch(plan, sheetPlan, batchStart, batchEnd, dataHeaders, batchCount)
            If batchCount > 0 Then
                outWs.Cells(writtenRows + 3, 1).Resize(batchCount, UBound(outputBatch, 2)).Value = outputBatch
                writtenRows = writtenRows + batchCount
            End If
            processedRows = processedRows + batchEnd - batchStart + 1
            Toolkit_ProgressUpdate processedRows, totalRows, CStr(sheetPlan("sheetName")) & " 읽는 중"
            batchStart = batchEnd + 1
        Loop
    Next i
    Toolkit_ProgressClose

    Dim lastCol As Long
    lastCol = dataHeaders.Count + 3
    outWs.Rows(1).Font.Color = RGB(120, 120, 120)
    outWs.Rows(1).Interior.Color = RGB(235, 235, 235)
    outWs.Rows(2).Font.Bold = True
    outWs.Rows(2).Interior.Color = RGB(220, 235, 250)
    With outWs.Range(outWs.Cells(1, 1), outWs.Cells(Application.Max(2, writtenRows + 2), lastCol))
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With

    DBMerge_WriteMergedSheet = writtenRows
End Function

Private Function DBMerge_BuildOutputBatch(ByVal plan As Object, ByVal sheetPlan As Object, _
                                          ByVal firstRow As Long, ByVal lastRow As Long, _
                                          ByVal dataHeaders As Collection, ByRef batchCount As Long) As Variant
    Dim ws As Worksheet, headerMap As Object
    Set ws = sheetPlan("worksheet")
    Set headerMap = sheetPlan("headerMap")

    Dim minCol As Long, maxCol As Long, i As Long, colNo As Long, norm As String
    minCol = ws.Columns.Count
    For i = 1 To plan("headers").Count
        norm = DBMerge_NormalizeHeader(CStr(plan("headers")(i)))
        If headerMap.Exists(norm) Then
            colNo = CLng(headerMap(norm))
            If colNo < minCol Then minCol = colNo
            If colNo > maxCol Then maxCol = colNo
        End If
    Next i
    If maxCol = 0 Then
        batchCount = 0
        Exit Function
    End If

    Dim sourceData As Variant
    sourceData = ws.Range(ws.Cells(firstRow, minCol), ws.Cells(lastRow, maxCol)).Value2
    Dim capacity As Long
    capacity = lastRow - firstRow + 1
    Dim pending() As Variant
    ReDim pending(1 To capacity, 1 To dataHeaders.Count + 3)

    Dim relativeRow As Long, c As Long, keyValue As String
    batchCount = 0
    For relativeRow = 1 To capacity
        If DBMerge_ArrayRowHasSelectedData(sourceData, relativeRow, minCol, headerMap, plan("headers"), CStr(plan("keyHeader"))) Then
            batchCount = batchCount + 1
            pending(batchCount, 1) = ws.Name
            pending(batchCount, 2) = firstRow + relativeRow - 1
            keyValue = Trim$(CStr(DBMerge_ArrayCellByHeader(sourceData, relativeRow, minCol, headerMap, CStr(plan("keyHeader")))))
            If Len(keyValue) = 0 Then keyValue = DBMerge_AutoId(ws.Name, firstRow + relativeRow - 1)
            pending(batchCount, 3) = keyValue
            For c = 1 To dataHeaders.Count
                pending(batchCount, c + 3) = DBMerge_ArrayCellByHeader(sourceData, relativeRow, minCol, headerMap, CStr(dataHeaders(c)))
            Next c
        End If
    Next relativeRow

    If batchCount = 0 Then Exit Function
    If batchCount = capacity Then
        DBMerge_BuildOutputBatch = pending
    Else
        Dim trimmed() As Variant, outputCol As Long
        ReDim trimmed(1 To batchCount, 1 To dataHeaders.Count + 3)
        For relativeRow = 1 To batchCount
            For outputCol = 1 To dataHeaders.Count + 3
                trimmed(relativeRow, outputCol) = pending(relativeRow, outputCol)
            Next outputCol
        Next relativeRow
        DBMerge_BuildOutputBatch = trimmed
    End If
End Function

Private Function DBMerge_ArrayRowHasSelectedData(ByVal sourceData As Variant, ByVal relativeRow As Long, _
                                                  ByVal firstSourceCol As Long, ByVal headerMap As Object, _
                                                  ByVal headers As Collection, ByVal keyHeader As String) As Boolean
    Dim keyNorm As String, norm As String, i As Long
    keyNorm = DBMerge_NormalizeHeader(keyHeader)
    For i = 1 To headers.Count
        norm = DBMerge_NormalizeHeader(CStr(headers(i)))
        If norm <> keyNorm And headerMap.Exists(norm) Then
            If Len(Trim$(CStr(DBMerge_ArrayValue(sourceData, relativeRow, CLng(headerMap(norm)) - firstSourceCol + 1)))) > 0 Then
                DBMerge_ArrayRowHasSelectedData = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function DBMerge_ArrayCellByHeader(ByVal sourceData As Variant, ByVal relativeRow As Long, _
                                            ByVal firstSourceCol As Long, ByVal headerMap As Object, _
                                            ByVal headerText As String) As Variant
    Dim norm As String
    norm = DBMerge_NormalizeHeader(headerText)
    If headerMap.Exists(norm) Then
        DBMerge_ArrayCellByHeader = DBMerge_ArrayValue(sourceData, relativeRow, CLng(headerMap(norm)) - firstSourceCol + 1)
    End If
End Function

Private Function DBMerge_ArrayValue(ByVal sourceData As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    If IsArray(sourceData) Then
        DBMerge_ArrayValue = sourceData(rowIndex, colIndex)
    ElseIf rowIndex = 1 And colIndex = 1 Then
        DBMerge_ArrayValue = sourceData
    End If
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

Private Sub DBMerge_ApplySourceIds(ByVal plan As Object, ByVal rollbackPlan As Collection)
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
        Dim rollbackItem As Object, changedCells As Collection
        Set rollbackItem = rollbackPlan(i)
        Set changedCells = rollbackItem("changedCells")

        Dim keyCol As Long
        If headerMap.Exists(keyNorm) Then
            keyCol = CLng(headerMap(keyNorm))
        Else
            keyCol = 1
            ws.Columns(keyCol).Insert Shift:=xlToRight
            rollbackItem("insertedColumn") = True
            DBMerge_WriteSourceIdHeader ws, CLng(sheetPlan("headerRow")), CStr(plan("headerMode")), keyHeader
            Set headerMap = DBMerge_HeaderMap(ws, CLng(sheetPlan("headerRow")), CStr(plan("headerMode")))
            sheetPlan.Remove "headerMap"
            sheetPlan.Add "headerMap", headerMap
        End If

        Dim r As Long
        For r = CLng(sheetPlan("dataStartRow")) To CLng(sheetPlan("lastRow"))
            If DBMerge_RowHasSelectedData(ws, r, headerMap, plan("headers"), keyHeader) Then
                If Len(Trim$(CStr(ws.Cells(r, keyCol).Value))) = 0 Then
                    If Not CBool(rollbackItem("insertedColumn")) Then
                        Dim cellState As Object
                        Set cellState = CreateObject("Scripting.Dictionary")
                        cellState.Add "address", ws.Cells(r, keyCol).Address
                        cellState.Add "formula2", ws.Cells(r, keyCol).Formula2
                        cellState.Add "numberFormat", ws.Cells(r, keyCol).NumberFormat
                        changedCells.Add cellState
                    End If
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

Private Function DBMerge_PrepareTemporaryOutputSheet(ByVal wb As Workbook, ByVal sheetName As String, _
                                                     ByVal outputPosition As String, _
                                                     ByRef existingOutput As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set existingOutput = wb.Worksheets(sheetName)
    On Error GoTo 0

    If Not existingOutput Is Nothing Then
        If MsgBox(sheetName & " 시트가 이미 있습니다. 내용을 지우고 다시 만들까요?", vbQuestion + vbYesNo, "시트 데이터 통합") = vbNo Then Exit Function
    End If

    Select Case outputPosition
        Case "맨 앞"
            Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        Case "현재 시트 뒤"
            Set ws = wb.Worksheets.Add(After:=ActiveSheet)
        Case Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End Select
    ws.Name = DBMerge_UniqueSheetName(wb, "__DBMERGE_TEMP")
    Set DBMerge_PrepareTemporaryOutputSheet = ws
End Function

Private Sub DBMerge_CommitTemporaryOutputSheet(ByVal temporarySheet As Worksheet, _
                                               ByVal existingOutput As Worksheet, _
                                               ByVal outputName As String)
    Dim backupName As String
    Dim previousAlerts As Boolean
    previousAlerts = Application.DisplayAlerts
    On Error GoTo CommitFailed

    If Not existingOutput Is Nothing Then
        backupName = DBMerge_UniqueSheetName(existingOutput.Parent, "__DBMERGE_BACKUP")
        existingOutput.Name = backupName
    End If
    temporarySheet.Name = outputName

    If Not existingOutput Is Nothing Then
        Application.DisplayAlerts = False
        existingOutput.Delete
    End If
    Application.DisplayAlerts = previousAlerts
    Exit Sub

CommitFailed:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not existingOutput Is Nothing And Len(backupName) > 0 Then
        If StrComp(temporarySheet.Name, outputName, vbTextCompare) = 0 Then
            temporarySheet.Name = DBMerge_UniqueSheetName(temporarySheet.Parent, "__DBMERGE_FAILED")
        End If
        existingOutput.Name = outputName
    End If
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0
    Err.Raise errorNumber, "DBMerge_CommitTemporaryOutputSheet", errorDescription
End Sub

Private Function DBMerge_CreateSourceIdRollbackPlan(ByVal plan As Object) As Collection
    Dim rollbackPlan As New Collection
    Dim i As Long

    For i = 1 To plan("sheets").Count
        Dim sheetPlan As Object, item As Object, ws As Worksheet, headerMap As Object
        Dim keyNorm As String, changedCells As Collection
        Set sheetPlan = plan("sheets")(i)
        Set ws = sheetPlan("worksheet")
        Set headerMap = sheetPlan("headerMap")
        keyNorm = DBMerge_NormalizeHeader(CStr(plan("keyHeader")))

        Set item = CreateObject("Scripting.Dictionary")
        Set changedCells = New Collection
        item.Add "worksheet", ws
        item.Add "insertedColumn", False
        item.Add "changedCells", changedCells
        rollbackPlan.Add item
    Next i

    Set DBMerge_CreateSourceIdRollbackPlan = rollbackPlan
End Function

Private Sub DBMerge_RollbackSourceIds(ByVal rollbackPlan As Collection)
    If rollbackPlan Is Nothing Then Exit Sub

    Dim i As Long
    For i = rollbackPlan.Count To 1 Step -1
        Dim item As Object, ws As Worksheet, changedCells As Collection, cellState As Variant
        Set item = rollbackPlan(i)
        Set ws = item("worksheet")
        If CBool(item("insertedColumn")) Then
            ws.Columns(1).Delete
        Else
            Set changedCells = item("changedCells")
            For Each cellState In changedCells
                ws.Range(CStr(cellState("address"))).Formula2 = cellState("formula2")
                ws.Range(CStr(cellState("address"))).NumberFormat = cellState("numberFormat")
            Next cellState
        End If
    Next i
End Sub

Private Sub DBMerge_DeleteWorksheetQuietly(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim previousAlerts As Boolean
    previousAlerts = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0
End Sub

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
