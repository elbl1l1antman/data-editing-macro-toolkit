Attribute VB_Name = "KosisTools"
Option Explicit

Private Const KOSIS_SETTINGS As String = "KOSIS"
Private Const KOSIS_ENDPOINT As String = "https://kosis.kr/openapi/statisticsData.do"

#If VBA7 Then
Private Declare PtrSafe Function Kosis_MoveFileEx Lib "kernel32" Alias "MoveFileExW" (ByVal sourcePath As LongPtr, ByVal targetPath As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function Kosis_MoveFileEx Lib "kernel32" Alias "MoveFileExW" (ByVal sourcePath As Long, ByVal targetPath As Long, ByVal flags As Long) As Long
#End If

Public Sub Kosis_SetApiKey()
    Dim answer As Variant, value As String
    answer = Application.InputBox("KOSIS 인증키를 입력하세요. 빈칸은 기존 키를 유지합니다.", "KOSIS 인증키 설정", Type:=2)
    If VarType(answer) = vbBoolean Then Exit Sub
    value = Trim$(CStr(answer))
    If Len(value) = 0 Then Exit Sub
    If Not Kosis_ValidKey(value) Then
        MsgBox "인증키 형식을 확인하세요.", vbExclamation, "KOSIS 인증키"
        Exit Sub
    End If
    SaveSetting BIZ_APP_NAME, KOSIS_SETTINGS, "ApiKey", value
    MsgBox "현재 Windows 사용자 설정에 인증키를 저장했습니다.", vbInformation, "KOSIS 인증키"
End Sub

Public Function Kosis_ValidKey(ByVal value As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9+/=]{8,256}$"
    Kosis_ValidKey = re.Test(value)
End Function

Public Function Kosis_PickItems(ByVal title As String, ByVal guide As String, ByVal metadata As Collection, _
                                ByVal axis As String, ByVal defaults As String, Optional ByRef goBack As Boolean = False) As Collection
    Dim form As Object, chosen As String, output As New Collection, value As Variant
    On Error GoTo Failed
    Set form = VBA.UserForms.Add("frmKosisPicker")
    goBack = False
    form.Configure title, guide, metadata, axis, defaults
    form.Show
    If Not CBool(form.Accepted) Then GoTo Cancelled
    chosen = CStr(form.SelectedKeys)
    For Each value In Split(chosen, vbLf)
        If Len(CStr(value)) > 0 Then output.Add CStr(value)
    Next value
    Unload form
    Set Kosis_PickItems = output
    Exit Function
Cancelled:
    goBack = CBool(form.BackRequested)
    Unload form
    Exit Function
Failed:
    If Not form Is Nothing Then Unload form
    Err.Raise 5, , "KOSIS 선택 창을 열지 못했습니다."
End Function

Public Function Kosis_PeriodChoices(ByVal periods As Collection) As Collection
    Dim result As New Collection, row As Object, entry As Object
    For Each row In periods
        Set entry = CreateObject("Scripting.Dictionary")
        entry.Add "axis", "period"
        entry.Add "code", CStr(row("type")) & CStr(row("period"))
        entry.Add "name", CStr(row("label"))
        entry.Add "parent", ""
        result.Add entry
    Next row
    Set Kosis_PeriodChoices = result
End Function

Public Function Kosis_FindOverlaps(ByVal selected As Collection, ByVal metadata As Collection, ByVal axis As String) As Collection
    Dim parents As Object, names As Object, chosen As Object, entry As Object, code As Variant
    Dim parent As String, current As String, depth As Long, result As New Collection, rootFound As Boolean, rootCode As String
    Set parents = CreateObject("Scripting.Dictionary"): Set names = CreateObject("Scripting.Dictionary")
    Set chosen = CreateObject("Scripting.Dictionary")
    For Each entry In metadata
        If CStr(entry("axis")) = axis Then
            current = CStr(entry("code"))
            If parents.Exists(current) Then Err.Raise 5, , "분류 코드가 중복되었습니다."
            parents.Add current, CStr(entry("parent"))
            names.Add current, CStr(entry("name"))
        End If
    Next entry
    For Each code In selected
        current = CStr(code)
        If Not parents.Exists(current) Then Err.Raise 5, , "선택한 분류의 메타정보가 없습니다."
        If chosen.Exists(current) Then Err.Raise 5, , "분류가 중복 선택되었습니다."
        chosen.Add current, True
    Next code
    Select Case axis
        Case "A": rootCode = "00"
        Case "B", "C", "E": rootCode = "0"
        Case "YRE": rootCode = "000"
    End Select
    For Each code In selected
        current = CStr(code): parent = CStr(parents(current)): depth = 0: rootFound = False
        Do While Len(parent) > 0
            If parent = rootCode Then rootFound = True
            If chosen.Exists(parent) Then result.Add CStr(names(parent)) & " [" & parent & "] / " & CStr(names(CStr(code))) & " [" & CStr(code) & "]"
            If Not parents.Exists(parent) Then Exit Do
            current = parent: parent = CStr(parents(current)): depth = depth + 1
            If depth > parents.Count Then Err.Raise 5, , "분류 계층이 순환합니다."
        Loop
        If Len(rootCode) > 0 And chosen.Exists(rootCode) And CStr(code) <> rootCode And Not rootFound Then _
            result.Add CStr(names(rootCode)) & " [" & rootCode & "] / " & CStr(names(CStr(code))) & " [" & CStr(code) & "]"
    Next code
    Set Kosis_FindOverlaps = result
End Function

Public Function Kosis_ReadMetadata(ByVal tableId As String, Optional ByVal refresh As Boolean = False, Optional ByVal cacheDirectory As String = "") As Collection
    Set Kosis_ReadMetadata = Kosis_ParseMetadata(Kosis_GetMetaXml(tableId, "ITM", refresh, cacheDirectory))
End Function

Public Function Kosis_ReadPeriods(ByVal tableId As String, Optional ByVal refresh As Boolean = False, Optional ByVal cacheDirectory As String = "") As Collection
    Set Kosis_ReadPeriods = Kosis_ParsePeriods(Kosis_GetMetaXml(tableId, "PRD", refresh, cacheDirectory))
End Function

Private Function Kosis_GetMetaXml(ByVal tableId As String, ByVal metaType As String, ByVal refresh As Boolean, ByVal cacheDirectory As String) As Object
    Dim key As String, http As Object, url As String, data As Object, cachePath As String, checked As Collection
    key = GetSetting(BIZ_APP_NAME, KOSIS_SETTINGS, "ApiKey", "")
    If Not Kosis_ValidKey(key) Then Err.Raise 5, , "먼저 KOSIS 인증키를 설정하세요."
    If Not Kosis_ValidTableId(tableId) Then Err.Raise 5, , "통계표 ID 형식이 올바르지 않습니다."
    cachePath = Kosis_CachePath(tableId, metaType, cacheDirectory)
    If Not refresh Then
        Set data = Kosis_LoadCache(cachePath, tableId, metaType)
        If Not data Is Nothing Then Set Kosis_GetMetaXml = data: Exit Function
    End If
    url = KOSIS_ENDPOINT & "?method=getMeta&type=" & metaType & "&detail=Y&format=xml&orgId=101&tblId=" & tableId & _
          "&apiKey=" & Application.WorksheetFunction.EncodeURL(key)
    On Error GoTo Failed
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 5000, 15000, 30000
    http.Open "GET", url, True
    http.Send
    Do While Not http.WaitForResponse(0.2)
        DoEvents
        Toolkit_RaiseIfCancelled
    Loop
    If http.Status <> 200 Then Err.Raise 5, , "KOSIS 서버 응답 오류"
    Set data = CreateObject("MSXML2.DOMDocument.6.0")
    data.async = False
    data.setProperty "ProhibitDTD", True
    data.resolveExternals = False
    If Not data.Load(http.ResponseStream) Then Err.Raise 5, , "KOSIS 응답을 읽을 수 없습니다."
    If Not Kosis_ValidMetaXml(data, tableId, metaType) Then Err.Raise 5, , "KOSIS 메타정보가 올바르지 않습니다."
    If metaType = "ITM" Then Set checked = Kosis_ParseMetadata(data) Else Set checked = Kosis_ParsePeriods(data)
    If checked.Count = 0 Then Err.Raise 5, , "KOSIS 메타정보가 비어 있습니다."
    Kosis_SaveCache data, cachePath
    Set Kosis_GetMetaXml = data
    Exit Function
Failed:
    Dim number As Long
    number = Err.Number
    On Error Resume Next
    If Not http Is Nothing Then http.Abort
    On Error GoTo 0
    If number = TOOLKIT_CANCEL_ERROR_NUMBER Then Err.Raise number
    Err.Raise 5, , "KOSIS 메타정보를 가져오지 못했습니다. 인증키, 네트워크와 통계표를 확인하세요."
End Function

Public Function Kosis_CacheTimestamp(ByVal tableId As String, Optional ByVal cacheDirectory As String = "") As String
    Dim path As String
    path = Kosis_CachePath(tableId, "ITM", cacheDirectory)
    If Len(Dir$(path)) > 0 Then Kosis_CacheTimestamp = Format$(FileDateTime(path), "yyyy-mm-dd hh:nn")
End Function

Private Function Kosis_CachePath(ByVal tableId As String, ByVal metaType As String, ByVal cacheDirectory As String) As String
    Dim root As String
    If Not Kosis_ValidTableId(tableId) Then Err.Raise 5, , "통계표 ID 형식이 올바르지 않습니다."
    If metaType <> "ITM" And metaType <> "PRD" Then Err.Raise 5, , "메타정보 종류가 올바르지 않습니다."
    root = cacheDirectory
    If Len(root) = 0 Then
        root = Environ$("LOCALAPPDATA")
        If Len(root) = 0 Then Err.Raise 5, , "사용자 캐시 경로를 찾을 수 없습니다."
        root = root & "\DataEditingMacroToolkit"
        If Len(Dir$(root, vbDirectory)) = 0 Then MkDir root
        root = root & "\KosisCache"
    End If
    If Len(Dir$(root, vbDirectory)) = 0 Then MkDir root
    Kosis_CachePath = root & "\" & tableId & "_" & metaType & ".xml"
End Function

Private Function Kosis_LoadCache(ByVal path As String, ByVal tableId As String, ByVal metaType As String) As Object
    Dim xml As Object, checked As Collection
    On Error GoTo Invalid
    If Len(Dir$(path)) = 0 Then Exit Function
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False: xml.resolveExternals = False: xml.setProperty "ProhibitDTD", True
    If Not xml.Load(path) Then Exit Function
    If Not Kosis_ValidMetaXml(xml, tableId, metaType) Then Exit Function
    If metaType = "ITM" Then Set checked = Kosis_ParseMetadata(xml) Else Set checked = Kosis_ParsePeriods(xml)
    If checked.Count = 0 Then Exit Function
    Set Kosis_LoadCache = xml
Invalid:
End Function

Private Function Kosis_ValidMetaXml(ByVal xml As Object, ByVal tableId As String, ByVal metaType As String) As Boolean
    Dim header As Object, rows As Object
    Set header = xml.SelectSingleNode("/response/Header/id")
    If header Is Nothing Then Exit Function
    If CStr(header.Text) <> "101_" & tableId Then Exit Function
    If metaType = "ITM" Then
        Set rows = xml.SelectNodes("/response/Structures/MetaRow")
    Else
        Set rows = xml.SelectNodes("/response/Structures/prdDe")
    End If
    Kosis_ValidMetaXml = (rows.Length > 0)
End Function

Private Sub Kosis_SaveCache(ByVal xml As Object, ByVal target As String)
    Dim fso As Object, temp As String, folder As String
    On Error GoTo Cleanup
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(target)
    temp = fso.BuildPath(folder, fso.GetTempName)
    xml.Save temp
    If Kosis_MoveFileEx(StrPtr(temp), StrPtr(target), 9) = 0 Then Err.Raise 5, , "분류정보 캐시를 저장하지 못했습니다."
    Exit Sub
Cleanup:
    Dim number As Long, message As String
    number = Err.Number: message = Err.Description
    On Error Resume Next
    If Len(temp) > 0 Then If fso.FileExists(temp) Then fso.DeleteFile temp, True
    On Error GoTo 0
    Err.Raise number, , message
End Sub

Public Function Kosis_ParsePeriods(ByVal xml As Object) As Collection
    Dim sections As Object, section As Object, node As Object
    Dim output As New Collection, row As Object, seen As Object, frequency As String, value As String
    Set seen = CreateObject("Scripting.Dictionary")
    If xml Is Nothing Then Err.Raise 5, , "수록 시점 정보가 없습니다."
    Set sections = xml.SelectNodes("/response/Structures")
    For Each section In sections
        frequency = ""
        For Each node In section.ChildNodes
            If node.nodeName = "prdSe" Then
                If node.Text = "년" Then
                    frequency = "Y"
                ElseIf node.Text = "월" Then
                    frequency = "M"
                Else
                    Err.Raise 5, , "지원하지 않는 수록 주기입니다."
                End If
            ElseIf node.nodeName = "prdDe" Then
            If Len(frequency) = 0 Then Err.Raise 5, , "수록 주기가 없습니다."
            value = Replace(Trim$(CStr(node.Text)), ".", "")
            If Not Kosis_ValidPeriod(value, frequency) Then Err.Raise 5, , "잘못된 수록 시점입니다."
            If seen.Exists(frequency & value) Then Err.Raise 5, , "수록 시점이 중복되었습니다."
            seen.Add frequency & value, True
            Set row = CreateObject("Scripting.Dictionary")
            row.Add "period", value
            row.Add "type", frequency
            row.Add "label", CStr(node.Text)
            output.Add row
            End If
        Next node
    Next section
    If output.Count = 0 Then Err.Raise 5, , "수록 시점을 확인할 수 없습니다."
    Set Kosis_ParsePeriods = output
End Function

Public Function Kosis_ParseMetadata(ByVal xml As Object) As Collection
    Dim rows As Object, node As Object, result As New Collection, entry As Object
    If xml Is Nothing Then Err.Raise 5, , "메타정보가 없습니다."
    Set rows = xml.SelectNodes("/response/Structures/MetaRow")
    If rows.Length = 0 Then Err.Raise 5, , "KOSIS에서 사용할 메타정보를 반환하지 않았습니다."
    For Each node In rows
        Set entry = CreateObject("Scripting.Dictionary")
        entry.Add "axis", Kosis_XmlText(node, "objId")
        entry.Add "axisName", Kosis_XmlText(node, "objNm")
        entry.Add "axisOrder", Kosis_XmlText(node, "objIdSn")
        entry.Add "code", Kosis_XmlText(node, "itmId")
        entry.Add "name", Kosis_XmlText(node, "itmNm")
        entry.Add "parent", Kosis_XmlText(node, "upItmId")
        entry.Add "unit", Kosis_XmlText(node, "unitNm")
        If Len(entry("axis")) = 0 Or Len(entry("code")) = 0 Or Len(entry("name")) = 0 Then Err.Raise 5, , "메타정보의 필수 항목이 없습니다."
        result.Add entry
    Next node
    Set Kosis_ParseMetadata = result
End Function

Public Function Kosis_ParseData(ByVal xml As Object) As Collection
    Dim series As Object, item As Object, observation As Object, valueNode As Object
    Dim result As New Collection, row As Object, field As Object, id As String, status As String
    If xml Is Nothing Then Err.Raise 5, , "통계자료가 없습니다."
    xml.SetProperty "SelectionLanguage", "XPath"
    xml.SetProperty "SelectionNamespaces", "xmlns:m='http://www.sdmx.org/resources/sdmxml/schemas/v2_1/message' xmlns:g='http://www.sdmx.org/resources/sdmxml/schemas/v2_1/data/generic'"
    Set series = xml.SelectNodes("/m:GenericData/m:DataSet/g:Series")
    If series.Length = 0 Then Err.Raise 5, , "KOSIS에서 사용할 통계자료를 반환하지 않았습니다."
    For Each item In series
        For Each observation In item.SelectNodes("g:Obs")
            Set row = CreateObject("Scripting.Dictionary")
            For Each field In item.SelectNodes("g:SeriesKey/g:Value")
                id = CStr(field.getAttribute("id"))
                If Len(id) = 0 Or row.Exists(id) Then Err.Raise 5, , "통계자료의 분류가 올바르지 않습니다."
                row.Add id, CStr(field.getAttribute("value"))
            Next field
            Set valueNode = observation.SelectSingleNode("g:ObsDimension")
            If valueNode Is Nothing Then Err.Raise 5, , "통계자료의 시점이 없습니다."
            row.Add "period", CStr(valueNode.getAttribute("value"))
            Set valueNode = observation.SelectSingleNode("g:ObsValue")
            If valueNode Is Nothing Then Err.Raise 5, , "통계자료의 값이 없습니다."
            row.Add "value", CStr(valueNode.getAttribute("value"))
            row.Add "observationXml", observation.XML
            status = ""
            For Each field In item.SelectNodes("g:Attributes/g:Value")
                If Len(status) > 0 Then status = status & "; "
                status = status & CStr(field.getAttribute("id")) & "=" & CStr(field.getAttribute("value"))
            Next field
            For Each field In observation.SelectNodes("g:Attributes/g:Value")
                If Len(status) > 0 Then status = status & "; "
                status = status & CStr(field.getAttribute("id")) & "=" & CStr(field.getAttribute("value"))
            Next field
            row.Add "status", status
            Set valueNode = observation.SelectSingleNode("g:LstChnDe")
            If valueNode Is Nothing Then row.Add "updated", "" Else row.Add "updated", CStr(valueNode.getAttribute("value"))
            result.Add row
        Next observation
    Next item
    If result.Count = 0 Then Err.Raise 5, , "KOSIS 통계자료에 관측값이 없습니다."
    Set Kosis_ParseData = result
End Function

' shortcut: 50 codes per axis per request; revise if KOSIS changes its request limits.
Public Function Kosis_FetchBatches(ByVal tableId As String, ByVal period As String, ByVal periodType As String, _
    ByVal itemId As String, ByVal axis1 As Collection, ByVal axis2 As Collection, ByVal axis3 As Collection, _
    ByVal axisId1 As String, ByVal axisId2 As String, ByVal axisId3 As String, ByRef missing As Collection, _
    Optional ByVal progressBase As Long = 0, Optional ByVal progressTotal As Long = 0) As Collection
    Dim first As Variant, second As Long, third As Long, i As Long, j As Long
    Dim endSecond As Long, endThird As Long, sel2 As String, sel3 As String, rows As Collection
    Dim expected As Object, seen As Object, output As New Collection, row As Object, token As Variant, key As String
    Dim secondCode As String, thirdCode As String, expectedTotal As Double
    If axis1 Is Nothing Or axis2 Is Nothing Or axis3 Is Nothing Then Err.Raise 5, , "선택하지 않은 분류가 있습니다."
    If axis1.Count = 0 Or axis2.Count = 0 Or Len(axisId1) = 0 Or Len(axisId2) = 0 Then Err.Raise 5, , "분류를 선택하세요."
    If axis3.Count > 0 And Len(axisId3) = 0 Then Err.Raise 5, , "세 번째 분류 정보가 없습니다."
    expectedTotal = CDbl(axis1.Count) * CDbl(axis2.Count) * IIf(axis3.Count > 0, CDbl(axis3.Count), 1#)
    If expectedTotal > 100000# Then Err.Raise 5, , "선택한 조합이 너무 많습니다. 범위를 나누어 조회하세요."
    Set expected = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set missing = New Collection
    For Each first In axis1
        For second = 1 To axis2.Count
            If axis3.Count = 0 Then
                key = Kosis_CombinationKey(CStr(first), CStr(axis2(second)), "")
                If expected.Exists(key) Then Err.Raise 5, , "같은 분류가 중복 선택되었습니다."
                expected.Add key, Array(CStr(first), CStr(axis2(second)), "")
            Else
                For third = 1 To axis3.Count
                    key = Kosis_CombinationKey(CStr(first), CStr(axis2(second)), CStr(axis3(third)))
                    If expected.Exists(key) Then Err.Raise 5, , "같은 분류가 중복 선택되었습니다."
                    expected.Add key, Array(CStr(first), CStr(axis2(second)), CStr(axis3(third)))
                Next third
            End If
        Next second
    Next first
    For Each first In axis1
        For second = 1 To axis2.Count Step 50
            endSecond = Application.Min(second + 49, axis2.Count)
            sel2 = ""
            For i = second To endSecond
                If Len(sel2) > 0 Then sel2 = sel2 & "+"
                sel2 = sel2 & CStr(axis2(i))
            Next i
            third = 1
            Do
                sel3 = "": endThird = 0
                If axis3.Count > 0 Then
                    endThird = Application.Min(third + 49, axis3.Count)
                    For j = third To endThird
                        If Len(sel3) > 0 Then sel3 = sel3 & "+"
                        sel3 = sel3 & CStr(axis3(j))
                    Next j
                End If
                If progressTotal > 0 Then
                    Toolkit_ProgressUpdate progressBase + seen.Count, progressTotal, "KOSIS 통계자료 조회 중"
                Else
                    Toolkit_ProgressUpdate seen.Count, CLng(expectedTotal), "KOSIS 통계자료 조회 중"
                End If
                Set rows = Kosis_ReadData(tableId, period, periodType, itemId, CStr(first), sel2, sel3)
                For Each row In rows
                    If Not row.Exists("C_" & axisId1) Or Not row.Exists("C_" & axisId2) Or Not row.Exists("ITEM") Then Err.Raise 5, , "응답에 분류 코드가 없습니다."
                    If CStr(row("ITEM")) <> itemId Or CStr(row("period")) <> period Then Err.Raise 5, , "요청하지 않은 항목 또는 시점이 반환되었습니다."
                    secondCode = CStr(row("C_" & axisId2)): thirdCode = ""
                    If axis3.Count > 0 Then
                        If Not row.Exists("C_" & axisId3) Then Err.Raise 5, , "응답에 세 번째 분류가 없습니다."
                        thirdCode = CStr(row("C_" & axisId3))
                    End If
                    key = Kosis_CombinationKey(CStr(row("C_" & axisId1)), secondCode, thirdCode)
                    If Not expected.Exists(key) Then Err.Raise 5, , "선택하지 않은 분류가 반환되었습니다."
                    If seen.Exists(key) Then Err.Raise 5, , "동일한 통계 조합이 중복 반환되었습니다."
                    seen.Add key, True
                    output.Add row
                Next row
                third = endThird + 1
            Loop While axis3.Count > 0 And third <= axis3.Count
        Next second
    Next first
    For Each token In expected.Keys
        If Not seen.Exists(CStr(token)) Then missing.Add expected(token)
    Next token
    Set Kosis_FetchBatches = output
End Function

Private Function Kosis_CombinationKey(ByVal first As String, ByVal second As String, ByVal third As String) As String
    Kosis_CombinationKey = Len(first) & ":" & first & Len(second) & ":" & second & Len(third) & ":" & third
End Function

' topAge denotes the entire open-ended KOSIS category, not a single-year count.
Public Function Kosis_ValidateAgeBands(ByVal bands As Collection, ByVal topAge As Long, ByRef uncovered As Long) As Boolean
    Dim positions() As Boolean, band As Object, first As Double, last As Double, age As Long
    If bands Is Nothing Then Err.Raise 5, , "연령 구간이 없습니다."
    If topAge < 1 Or topAge > 200 Or bands.Count = 0 Then Err.Raise 5, , "연령 구간을 확인하세요."
    ReDim positions(0 To topAge)
    For Each band In bands
        If Not band.Exists("name") Or Not band.Exists("from") Or Not band.Exists("to") Then Err.Raise 5, , "구간 정보가 누락되었습니다."
        If Len(Trim$(CStr(band("name")))) = 0 Then Err.Raise 5, , "구간 이름을 입력하세요."
        If Not IsNumeric(band("from")) Or Not IsNumeric(band("to")) Then Err.Raise 5, , "구간 경계는 정수여야 합니다."
        first = CDbl(band("from")): last = CDbl(band("to"))
        If first <> Fix(first) Or last <> Fix(last) Or first < 0 Or last > topAge Or first > last Then Err.Raise 5, , "연령 구간 경계를 확인하세요."
        For age = CLng(first) To CLng(last)
            If positions(age) Then Err.Raise 5, , "연령 구간이 겹칩니다."
            positions(age) = True
        Next age
    Next band
    uncovered = 0
    For age = 0 To topAge
        If Not positions(age) Then uncovered = uncovered + 1
    Next age
    Kosis_ValidateAgeBands = True
End Function

' The caller must split requests below the KOSIS cell limit and check completeness.
Public Function Kosis_ReadData(ByVal tableId As String, ByVal period As String, ByVal periodType As String, _
                               ByVal itemId As String, ByVal axis1 As String, ByVal axis2 As String, _
                               ByVal axis3 As String) As Collection
    Dim key As String, http As Object, url As String, xml As Object, number As Long
    key = GetSetting(BIZ_APP_NAME, KOSIS_SETTINGS, "ApiKey", "")
    If Not Kosis_ValidKey(key) Then Err.Raise 5, , "먼저 KOSIS 인증키를 설정하세요."
    If Not Kosis_ValidTableId(tableId) Or Not Kosis_ValidSelector(itemId) Or _
       Not Kosis_ValidSelector(axis1) Or Not Kosis_ValidSelector(axis2) Or _
       (Len(axis3) > 0 And Not Kosis_ValidSelector(axis3)) Then Err.Raise 5, , "조회 조건이 올바르지 않습니다."
    If Not Kosis_ValidPeriod(period, periodType) Then Err.Raise 5, , "조회 시점이 올바르지 않습니다."
    url = "https://kosis.kr/openapi/Param/statisticsParameterData.do?method=getList&format=sdmx&type=Generic&version=v2_1" & _
          "&orgId=101&tblId=" & tableId & "&itmId=" & Application.WorksheetFunction.EncodeURL(itemId) & _
          "&objL1=" & Application.WorksheetFunction.EncodeURL(axis1) & "&objL2=" & Application.WorksheetFunction.EncodeURL(axis2)
    If Len(axis3) > 0 Then url = url & "&objL3=" & Application.WorksheetFunction.EncodeURL(axis3)
    url = url & "&prdSe=" & periodType & "&startPrdDe=" & period & "&endPrdDe=" & period & "&smblChk=Y&apiKey=" & Application.WorksheetFunction.EncodeURL(key)
    On Error GoTo Failed
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 5000, 15000, 30000
    http.Open "GET", url, True
    http.Send
    Do While Not http.WaitForResponse(0.2)
        DoEvents
        Toolkit_RaiseIfCancelled
    Loop
    If http.Status <> 200 Then Err.Raise 5
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    xml.setProperty "ProhibitDTD", True
    xml.resolveExternals = False
    If Not xml.Load(http.ResponseStream) Then Err.Raise 5
    Set Kosis_ReadData = Kosis_ParseData(xml)
    Exit Function
Failed:
    number = Err.Number
    On Error Resume Next
    If Not http Is Nothing Then http.Abort
    On Error GoTo 0
    If number = TOOLKIT_CANCEL_ERROR_NUMBER Then Err.Raise number
    Err.Raise 5, , "KOSIS 통계자료 조회에 실패했습니다. 인증키와 조회 조건을 확인하세요."
End Function

Private Function Kosis_ValidSelector(ByVal value As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9_*+,-]{1,4000}$"
    Kosis_ValidSelector = re.Test(value)
End Function

Private Function Kosis_ValidPeriod(ByVal value As String, ByVal periodType As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    If periodType = "Y" Then
        re.Pattern = "^[0-9]{4}$"
    ElseIf periodType = "M" Then
        re.Pattern = "^[0-9]{4}(0[1-9]|1[0-2])$"
    Else
        Exit Function
    End If
    Kosis_ValidPeriod = re.Test(value)
End Function

Private Function Kosis_XmlText(ByVal node As Object, ByVal name As String) As String
    Dim child As Object
    Set child = node.SelectSingleNode(name)
    If Not child Is Nothing Then Kosis_XmlText = child.Text
End Function

Private Function Kosis_ValidTableId(ByVal value As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9_]{1,80}$"
    Kosis_ValidTableId = re.Test(value)
End Function
