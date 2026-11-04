Attribute VB_Name = "ApiTools"
Option Explicit


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

Public Sub NaverBusinessSearch_SearchSelection()
    Const PURPOSE As String = "네이버 업체 검색"

    If Not NaverLocal_HasCredentials() Then
        MsgBox "먼저 [네이버 인증키 설정] 버튼으로 Client ID와 Client Secret을 저장하세요.", vbExclamation, PURPOSE
        Exit Sub
    End If

    Dim nameRange As Range
    If Not Toolkit_GetRangeFromUserForm(PURPOSE, _
        "검색할 사업체명/기업체명 범위를 선택하세요." & vbCrLf & _
        "선택 범위 오른쪽에 네이버 지역검색 주소와 웹문서 " & NaverBusinessSearch_DefaultKeyword() & " 후보를 추가합니다.", _
        nameRange) Then Exit Sub
    If nameRange Is Nothing Then Exit Sub

    Set nameRange = nameRange.Areas(1)
    If nameRange.Columns.Count > 1 Then Set nameRange = nameRange.Columns(1)

    Dim keywordText As String
    keywordText = InputBox("업체명 뒤에 붙일 검색 키워드를 입력하세요." & vbCrLf & _
                           "예: " & NaverBusinessSearch_DefaultKeyword() & ", " & _
                           NaverBusinessSearch_OfficialHomepageKeyword() & ", " & _
                           NaverBusinessSearch_OfficialMallKeyword(), _
                           PURPOSE, GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUSINESS_KEYWORD_KEY, NaverBusinessSearch_DefaultKeyword()))
    If StrPtr(keywordText) = 0 Then Exit Sub
    keywordText = Trim$(keywordText)
    SaveSetting NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_BUSINESS_KEYWORD_KEY, keywordText

    Dim headerAnswer As VbMsgBoxResult
    headerAnswer = MsgBox("선택 범위의 첫 행을 헤더로 볼까요?", vbQuestion + vbYesNoCancel, PURPOSE)
    If headerAnswer = vbCancel Then Exit Sub
    Dim hasHeader As Boolean
    hasHeader = (headerAnswer = vbYes)

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

    Dim headers As Variant
    headers = NaverBusinessSearch_ResultHeaders()

    Dim outputCol As Long
    outputCol = nameRange.Column + nameRange.Columns.Count
    If Not Toolkit_CanInsertColumns(ws, outputCol, UBound(headers) - LBound(headers) + 1, PURPOSE) Then Exit Sub

    Dim totalRows As Long
    totalRows = lastRow - firstDataRow + 1
    If totalRows > 500 Then
        If MsgBox("검색 대상이 " & totalRows & "행입니다." & vbCrLf & _
                  "행마다 네이버 지역검색과 웹문서 API를 호출하므로 시간이 걸릴 수 있습니다. 계속할까요?", _
                  vbQuestion + vbYesNo, PURPOSE) <> vbYes Then Exit Sub
    End If

    Dim defaultDelay As Double
    defaultDelay = CDbl(Val(GetSetting(NAVER_APP_NAME, NAVER_APP_SECTION, NAVER_MALL_DELAY_KEY, "0.3")))
    If defaultDelay < 0 Then defaultDelay = 0.3

    Dim outputData() As Variant
    ReDim outputData(1 To totalRows, 1 To UBound(headers) - LBound(headers) + 1)

    Dim rowIndex As Long
    Dim outIndex As Long
    Dim companyName As String
    Dim queryText As String
    Dim localItems As Collection
    Dim webItems As Collection
    Dim localItem As Object
    Dim webItem As Object
    Dim requestCount As Long
    Dim errorCount As Long
    Dim localFoundCount As Long
    Dim webFoundCount As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ws.Columns(outputCol).Resize(, UBound(headers) - LBound(headers) + 1).Insert Shift:=xlToRight
    If hasHeader Then NaverBusinessSearch_WriteHeaders ws, nameRange.Row, outputCol

    Toolkit_ProgressShow PURPOSE, "검색 준비 중", totalRows

    For rowIndex = firstDataRow To lastRow
        outIndex = rowIndex - firstDataRow + 1
        companyName = Trim$(CStr(ws.Cells(rowIndex, nameRange.Column).Value))
        Toolkit_ProgressUpdate outIndex, totalRows, "네이버 업체 검색 중: " & CStr(rowIndex) & "행"

        If Len(companyName) > 0 Then
            queryText = NaverBusinessSearch_BuildQuery(companyName, keywordText)
            outputData(outIndex, 1) = queryText
            outputData(outIndex, 11) = Now
            requestCount = requestCount + 1
            Set localItems = Nothing
            Set webItems = Nothing
            Set localItem = Nothing
            Set webItem = Nothing

            On Error Resume Next
            Set localItems = NaverLocal_Search(queryText, 5, "random")
            If Err.Number <> 0 Then
                errorCount = errorCount + 1
                Toolkit_LogApiEvent ws.Parent, PURPOSE, "local", queryText, "ERROR", Err.Description
                Err.Clear
            ElseIf Not localItems Is Nothing Then
                Set localItem = NaverBusinessSearch_FirstCollectionItem(localItems)
                If Not localItem Is Nothing Then
                    localFoundCount = localFoundCount + 1
                    outputData(outIndex, 2) = CStr(localItem("title"))
                    outputData(outIndex, 3) = CStr(localItem("link"))
                    outputData(outIndex, 4) = CStr(localItem("telephone"))
                    outputData(outIndex, 5) = CStr(localItem("address"))
                    outputData(outIndex, 6) = CStr(localItem("roadAddress"))
                End If
            End If

            Set webItems = NaverOfficialMall_SearchWeb(queryText, 10)
            If Err.Number <> 0 Then
                errorCount = errorCount + 1
                Toolkit_LogApiEvent ws.Parent, PURPOSE, "webkr", queryText, "ERROR", Err.Description
                Err.Clear
            Else
                Set webItem = NaverBusinessSearch_FirstWebsiteCandidate(webItems, companyName)
                If Not webItem Is Nothing Then
                    webFoundCount = webFoundCount + 1
                    outputData(outIndex, 7) = CStr(webItem("link"))
                    outputData(outIndex, 8) = CStr(webItem("title"))
                    outputData(outIndex, 9) = CStr(webItem("description"))
                    outputData(outIndex, 10) = NaverOfficialMall_Domain(CStr(webItem("link")))
                End If
            End If
            On Error GoTo ErrHandler

            NaverLocal_Delay defaultDelay
        End If
    Next rowIndex

    ws.Range(ws.Cells(firstDataRow, outputCol), ws.Cells(lastRow, outputCol + UBound(headers))).Value = outputData
    If hasHeader Then ws.Range(ws.Cells(nameRange.Row, outputCol), ws.Cells(nameRange.Row, outputCol + UBound(headers))).Font.Bold = True
    ws.Columns(outputCol).Resize(, UBound(headers) - LBound(headers) + 1).AutoFit

    Toolkit_ProgressClose
    Application.ScreenUpdating = True

    MsgBox "네이버 업체 검색을 완료했습니다." & vbCrLf & _
           "조회: " & requestCount & "건" & vbCrLf & _
           "지역검색 결과: " & localFoundCount & "건" & vbCrLf & _
           "웹문서 후보: " & webFoundCount & "건" & vbCrLf & _
           "조회 오류: " & errorCount & "건", vbInformation, PURPOSE
    Exit Sub

ErrHandler:
    Toolkit_ProgressClose
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage(PURPOSE) Then Exit Sub
    MsgBox "네이버 업체 검색 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbCritical, PURPOSE
End Sub

Public Function NaverBusinessSearch_BuildQueryForTest(ByVal companyName As String, ByVal keywordText As String) As String
    NaverBusinessSearch_BuildQueryForTest = NaverBusinessSearch_BuildQuery(companyName, keywordText)
End Function

Public Function NaverBusinessSearch_ResultHeaderTextForTest() As String
    Dim headers As Variant
    headers = NaverBusinessSearch_ResultHeaders()

    Dim resultText As String
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If Len(resultText) > 0 Then resultText = resultText & "|"
        resultText = resultText & CStr(headers(i))
    Next i
    NaverBusinessSearch_ResultHeaderTextForTest = resultText
End Function

Private Function NaverBusinessSearch_BuildQuery(ByVal companyName As String, ByVal keywordText As String) As String
    companyName = Trim$(CStr(companyName))
    keywordText = Trim$(CStr(keywordText))
    If Len(keywordText) = 0 Then
        NaverBusinessSearch_BuildQuery = companyName
    Else
        NaverBusinessSearch_BuildQuery = Trim$(companyName & " " & keywordText)
    End If
End Function

Private Function NaverBusinessSearch_DefaultKeyword() As String
    NaverBusinessSearch_DefaultKeyword = ChrW(&HACF5) & ChrW(&HC2DD) & " " & ChrW(&HC0AC) & ChrW(&HC774) & ChrW(&HD2B8)
End Function

Private Function NaverBusinessSearch_OfficialHomepageKeyword() As String
    NaverBusinessSearch_OfficialHomepageKeyword = ChrW(&HACF5) & ChrW(&HC2DD) & " " & ChrW(&HD648) & ChrW(&HD398) & ChrW(&HC774) & ChrW(&HC9C0)
End Function

Private Function NaverBusinessSearch_OfficialMallKeyword() As String
    NaverBusinessSearch_OfficialMallKeyword = ChrW(&HACF5) & ChrW(&HC2DD) & ChrW(&HBAB0)
End Function

Private Function NaverBusinessSearch_ResultHeaders() As Variant
    NaverBusinessSearch_ResultHeaders = Array("검색어", "지역검색 업체명", "지역검색 링크", "전화번호", "지번주소", "도로명주소", _
                                              "웹문서 후보 URL", "웹문서 후보 제목", "웹문서 후보 요약", "웹문서 후보 도메인", "검색일시")
End Function

Private Sub NaverBusinessSearch_WriteHeaders(ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long)
    Dim headers As Variant
    headers = NaverBusinessSearch_ResultHeaders()

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowIndex, colIndex + i).Value = headers(i)
    Next i
End Sub

Private Function NaverBusinessSearch_FirstCollectionItem(ByVal items As Collection) As Object
    If items Is Nothing Then Exit Function
    If items.Count = 0 Then Exit Function
    Set NaverBusinessSearch_FirstCollectionItem = items.Item(1)
End Function

Private Function NaverBusinessSearch_FirstWebsiteCandidate(ByVal items As Collection, ByVal companyName As String) As Object
    If items Is Nothing Then Exit Function
    If items.Count = 0 Then Exit Function

    Dim item As Object
    For Each item In items
        Dim classCode As String
        classCode = NaverOfficialMall_ClassifyUrl(CStr(item("link")), CStr(item("title")), CStr(item("description")), companyName)
        If Left$(classCode, 10) = "candidate:" Then
            Set NaverBusinessSearch_FirstWebsiteCandidate = item
            Exit Function
        End If
    Next item
End Function

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
    Dim nameRange As Range
    Dim addressRange As Range

    queryMode = CStr(form.QueryMode)
    displayCount = CLng(form.DisplayCount)
    delaySeconds = CDbl(form.DelaySeconds)
    sortMode = CStr(form.SortMode)
    hasHeader = CBool(form.HasHeader)
    useName = CBool(form.UseName)
    useAddress = CBool(form.UseAddress)
    Unload form

    If useName Then
        If Not Toolkit_GetRangeFromUserForm("네이버 주소 검증 - 업체명 범위", _
            "업체명 또는 기업체명 범위를 선택하세요.", nameRange) Then Exit Function
    End If
    If useAddress Then
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
    Toolkit_ProgressClose
    On Error GoTo 0
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

Public Function NaverLocal_Similarity(ByVal leftText As String, ByVal rightText As String) As Double
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
