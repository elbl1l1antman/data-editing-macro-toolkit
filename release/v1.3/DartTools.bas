Attribute VB_Name = "DartTools"
Option Explicit

Private Const DART_BASE As String = "https://opendart.fss.or.kr/api/"
Private Const DART_SETTINGS As String = "OpenDART"

Public Sub Dart_SetApiKey()
    Dim valueIn As Variant
    valueIn = Application.InputBox("OpenDART 인증키 40자리를 입력하세요. 빈칸은 기존 키를 유지합니다.", "DART 인증키 설정", Type:=2)
    If VarType(valueIn) = vbBoolean Then Exit Sub
    If Len(Trim$(CStr(valueIn))) = 0 Then Exit Sub
    If Not Dart_ValidKey(Trim$(CStr(valueIn))) Then
        MsgBox "영문과 숫자로 된 OpenDART 인증키 40자리를 확인하세요.", vbExclamation, "DART 인증키"
        Exit Sub
    End If
    SaveSetting BIZ_APP_NAME, DART_SETTINGS, "ApiKey", Trim$(CStr(valueIn))
    MsgBox "인증키를 현재 Windows 사용자 설정에 저장했습니다.", vbInformation, "DART 인증키"
End Sub

Public Function Dart_ValidKey(ByVal key As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9]{40}$"
    Dart_ValidKey = re.Test(key)
End Function

Public Sub Dart_SearchCompanies()
    On Error GoTo Failed
    Dim key As String, target As Range, wb As Workbook, resultSheet As Worksheet
    Dim yearValue As Variant, headerChoice As VbMsgBoxResult
    Dim master As Object, codeMap As Object, stockMap As Object, cached As Object
    Dim cells As Variant, output() As Variant, rowIndex As Long, firstData As Long
    Dim lookupText As String, corpCode As String, issue As String, fields As Variant
    Dim rowCount As Long, i As Long, started As Boolean

    key = GetSetting(BIZ_APP_NAME, DART_SETTINGS, "ApiKey", "")
    If Not Dart_ValidKey(key) Then
        MsgBox "먼저 DART 인증키 설정에서 인증키를 입력하세요.", vbExclamation, "DART 기업정보"
        Exit Sub
    End If
    If Not Toolkit_GetRangeFromUserForm("DART 기업정보", "기업명, DART 고유번호 또는 종목코드가 있는 한 열을 선택하세요.", target) Then Exit Sub
    If target.Areas.Count <> 1 Or target.Columns.Count <> 1 Then
        MsgBox "연속된 한 개 열만 선택하세요.", vbExclamation, "DART 기업정보"
        Exit Sub
    End If
    Set wb = target.Worksheet.Parent
    If wb.ReadOnly Or wb.ProtectStructure Then
        MsgBox "결과 시트를 만들 수 있도록 읽기 전용 또는 통합문서 구조 보호를 해제하세요.", vbExclamation, "DART 기업정보"
        Exit Sub
    End If
    headerChoice = MsgBox("선택 범위의 첫 행이 제목인가요?" & vbCrLf & "예: 첫 행 제외 / 아니요: 첫 행부터 조회", vbYesNoCancel + vbQuestion, "DART 기업정보")
    If headerChoice = vbCancel Then Exit Sub
    firstData = 1
    If headerChoice = vbYes Then firstData = 2
    If target.Rows.Count < firstData Then Exit Sub
    yearValue = Application.InputBox("매출액과 종사자 수를 조회할 사업연도를 입력하세요. 사업보고서 기준입니다.", "DART 사업연도", Year(Date) - 1, Type:=1)
    If VarType(yearValue) = vbBoolean Then Exit Sub
    If yearValue <> Fix(yearValue) Or yearValue < 2015 Or yearValue > Year(Date) Then
        MsgBox "2015년부터 현재 연도 사이의 사업연도를 입력하세요.", vbExclamation, "DART 사업연도"
        Exit Sub
    End If

    rowCount = target.Rows.Count - firstData + 1
    If rowCount > 10000 Then
        MsgBox "한 번에 최대 10,000행까지 선택할 수 있습니다. API 호출 제한을 고려해 나누어 조회하세요.", vbExclamation, "DART 기업정보"
        Exit Sub
    End If
    cells = target.Value2
    ReDim output(1 To rowCount, 1 To 21)
    Set master = CreateObject("Scripting.Dictionary")
    Set codeMap = CreateObject("Scripting.Dictionary")
    Set stockMap = CreateObject("Scripting.Dictionary")
    Set cached = CreateObject("Scripting.Dictionary")
    Toolkit_ProgressShow "DART 기업정보", "DART 고유번호 목록을 준비 중입니다.", rowCount
    started = True
    Dart_LoadCompanyMaster key, master, codeMap, stockMap

    For rowIndex = 1 To rowCount
        If target.Rows.Count = 1 Then
            If Not IsError(cells) Then lookupText = Trim$(CStr(cells)) Else lookupText = ""
        Else
            If Not IsError(cells(rowIndex + firstData - 1, 1)) Then lookupText = Trim$(CStr(cells(rowIndex + firstData - 1, 1))) Else lookupText = ""
        End If
        output(rowIndex, 1) = target.Row + rowIndex + firstData - 2
        output(rowIndex, 2) = lookupText
        issue = ""
        corpCode = Dart_FindCompany(lookupText, master, codeMap, stockMap, issue)
        If Len(corpCode) > 0 Then
            If Not cached.Exists(corpCode) Then cached.Add corpCode, Dart_QueryCompany(key, corpCode, CLng(yearValue))
            fields = cached(corpCode)
            For i = 0 To 18
                output(rowIndex, i + 3) = fields(i)
            Next i
        Else
            output(rowIndex, 21) = issue
        End If
        Toolkit_ProgressUpdate rowIndex, rowCount, "기업 조회 중 " & rowIndex & "/" & rowCount
    Next rowIndex
    Toolkit_ProgressClose
    started = False

    Set resultSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    resultSheet.Name = Dart_UniqueSheetName(wb, "DART_기업정보")
    resultSheet.Range("A1:U1").Value = Array("원본 행", "입력값", "DART 고유번호", "기업명", "영문명", "종목코드", "대표자", "법인구분", "사업자등록번호", "법인등록번호", "주소", "홈페이지", "전화번호", "업종코드", "사업연도", "매출액", "매출 통화", "매출 계정", "재무 기준", "종사자 수", "검토 사항")
    resultSheet.Range("A2:U" & rowCount + 1).NumberFormat = "@"
    resultSheet.Range("A2:U" & rowCount + 1).Value2 = output
    resultSheet.Rows(1).Font.Bold = True
    resultSheet.Rows(1).AutoFilter
    resultSheet.Columns("A:U").AutoFit
    MsgBox rowCount & "건을 조회했습니다. 결과 시트의 검토 사항과 공시 기준을 확인하세요.", vbInformation, "DART 기업정보"
    Exit Sub
Failed:
    Dim failureNumber As Long, failureDescription As String
    failureNumber = Err.Number: failureDescription = Replace(Err.Description, key, "[인증키 숨김]")
    If started Then Toolkit_ProgressClose
    If Not resultSheet Is Nothing Then
        Dim oldAlerts As Boolean
        oldAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        resultSheet.Delete
        Application.DisplayAlerts = oldAlerts
    End If
    If failureNumber = TOOLKIT_CANCEL_ERROR_NUMBER Then
        MsgBox "작업을 취소했습니다. 원본 데이터는 변경되지 않았습니다.", vbInformation, "DART 기업정보"
        Exit Sub
    End If
    MsgBox "조회가 중단되었습니다. 원본 데이터는 변경되지 않았습니다." & vbCrLf & failureDescription, vbExclamation, "DART 기업정보"
End Sub

Public Function Dart_UniqueSheetName(ByVal wb As Workbook, ByVal base As String) As String
    Dim name As String, i As Long, ws As Object
    name = base
    Do
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Sheets(name)
        On Error GoTo 0
        If ws Is Nothing Then Exit Do
        i = i + 1
        name = Left$(base, 25) & "_" & CStr(i)
    Loop
    Dart_UniqueSheetName = name
End Function

Public Function Dart_NormalizeName(ByVal valueIn As String) As String
    Dim s As String
    s = UCase$(Trim$(valueIn))
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(&H3000), "")
    s = Replace(s, "㈜", "")
    s = Replace(s, "(주)", "")
    s = Replace(s, "주식회사", "")
    Dart_NormalizeName = s
End Function

Public Function Dart_FindCompany(ByVal valueIn As String, ByVal names As Object, ByVal codes As Object, ByVal stocks As Object, ByRef issue As String) As String
    Dim s As String
    s = Trim$(valueIn)
    If Len(s) = 0 Then issue = "입력값 없음": Exit Function
    If Len(s) = 8 And s Like "########" Then
        If codes.Exists(s) Then Dart_FindCompany = s Else issue = "고유번호 미확인"
        Exit Function
    End If
    If Len(s) = 6 And s Like "######" Then
        If stocks.Exists(s) Then
            If Len(CStr(stocks(s))) > 0 Then Dart_FindCompany = stocks(s) Else issue = "종목코드 중복: 고유번호 확인 필요"
        Else
            issue = "종목코드 미확인"
        End If
        Exit Function
    End If
    s = Dart_NormalizeName(s)
    If Len(s) = 0 Then issue = "기업명 확인 필요": Exit Function
    If Not names.Exists(s) Then issue = "기업명 미일치: 고유번호 확인 필요": Exit Function
    If Len(CStr(names(s))) = 0 Then issue = "동명 기업: 고유번호 확인 필요": Exit Function
    Dart_FindCompany = CStr(names(s))
End Function

Public Sub Dart_LoadCompanyMaster(ByVal key As String, ByVal names As Object, ByVal codes As Object, ByVal stocks As Object)
    Dim fso As Object, folder As String, zipPath As String, xmlPath As String
    Dim http As Object, stream As Object, shell As Object, src As Object, dst As Object, runner As Object
    Dim doc As Object, nodes As Object, node As Object, corp As String, nm As String, stock As String, body As Variant
    Dim started As Date, lastSize As Double, stable As Long, processed As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(fso.GetSpecialFolder(2), "DartToolkit_" & Replace(fso.GetTempName, ".", ""))
    fso.CreateFolder folder
    On Error GoTo Failed
    zipPath = fso.BuildPath(folder, "master.zip")
    xmlPath = fso.BuildPath(folder, "CORPCODE.xml")
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 10000, 30000, 60000
    http.Open "GET", DART_BASE & "corpCode.xml?crtfc_key=" & key, False
    http.send
    If http.Status <> 200 Then Err.Raise vbObjectError + 9600, , "DART 고유번호 목록을 내려받지 못했습니다. HTTP " & http.Status
    body = http.responseBody
    If UBound(body) < 1 Then Err.Raise vbObjectError + 9600, , "DART 고유번호 응답이 비어 있습니다."
    If body(0) <> 80 Or body(1) <> 75 Then
        Set doc = CreateObject("MSXML2.DOMDocument.6.0")
        doc.async = False: doc.setProperty "ProhibitDTD", True
        If doc.LoadXML(http.responseText) Then
            Err.Raise vbObjectError + 9600, , "DART 고유번호 API 오류 " & Dart_NodeText(doc, "/result/status")
        End If
        Err.Raise vbObjectError + 9600, , "DART 고유번호 응답 형식이 올바르지 않습니다."
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write body: stream.SaveToFile zipPath, 2: stream.Close
    Set shell = CreateObject("Shell.Application")
    Set dst = shell.NameSpace(folder)
    Set src = shell.NameSpace(zipPath)
    If src Is Nothing Or dst Is Nothing Then
        Dim tarPath As String, command As String
        tarPath = Environ$("SystemRoot") & "\System32\tar.exe"
        If Not fso.FileExists(tarPath) Then tarPath = Environ$("SystemRoot") & "\Sysnative\tar.exe"
        If Not fso.FileExists(tarPath) Then Err.Raise vbObjectError + 9601, , "Windows ZIP 압축 해제 기능을 사용할 수 없습니다."
        command = """" & tarPath & """ -xf """ & zipPath & """ -C """ & folder & """ CORPCODE.xml"
        Set runner = CreateObject("WScript.Shell")
        runner.Run command, 0, False
    Else
        If src.ParseName("CORPCODE.xml") Is Nothing Then Err.Raise vbObjectError + 9601, , "고유번호 ZIP에 필요한 XML이 없습니다."
        dst.CopyHere src.ParseName("CORPCODE.xml"), 20
    End If
    started = Now
    Do
        DoEvents
        Toolkit_RaiseIfCancelled
        If fso.FileExists(xmlPath) Then
            If fso.GetFile(xmlPath).Size = lastSize And lastSize > 0 Then stable = stable + 1 Else stable = 0
            lastSize = fso.GetFile(xmlPath).Size
            If stable >= 10 Then Exit Do
        End If
        If DateDiff("s", started, Now) > 90 Then Err.Raise vbObjectError + 9602, , "고유번호 목록 압축 해제가 지연되었습니다. 다시 시도하세요."
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False: doc.setProperty "ProhibitDTD", True
    If Not doc.Load(xmlPath) Then Err.Raise vbObjectError + 9603, , "고유번호 XML을 읽지 못했습니다."
    Set nodes = doc.SelectNodes("/result/list")
    If nodes.Length = 0 Then Err.Raise vbObjectError + 9604, , "DART 고유번호 목록이 비어 있습니다."
    For Each node In nodes
        processed = processed + 1
        If processed Mod 1000 = 0 Then
            DoEvents
            Toolkit_RaiseIfCancelled
        End If
        corp = Dart_NodeText(node, "corp_code")
        nm = Dart_NormalizeName(Dart_NodeText(node, "corp_name"))
        stock = Trim$(Dart_NodeText(node, "stock_code"))
        If Len(corp) = 8 Then
            codes(corp) = True
            If Len(nm) > 0 Then
                If names.Exists(nm) Then
                    If CStr(names(nm)) <> corp Then names(nm) = ""
                Else
                    names.Add nm, corp
                End If
            End If
            If Len(stock) = 6 Then
                If stocks.Exists(stock) Then
                    If CStr(stocks(stock)) <> corp Then stocks(stock) = ""
                Else
                    stocks.Add stock, corp
                End If
            End If
        End If
    Next node
    On Error Resume Next
    fso.DeleteFolder folder, True
    Exit Sub
Failed:
    Dim n As Long, description As String
    n = Err.Number: description = Replace(Err.Description, key, "[인증키 숨김]")
    On Error Resume Next
    If fso.FolderExists(folder) Then fso.DeleteFolder folder, True
    On Error GoTo 0
    Err.Raise n, "Dart_LoadCompanyMaster", description
End Sub

Public Function Dart_NodeText(ByVal parent As Object, ByVal path As String) As String
    Dim node As Object
    Set node = parent.SelectSingleNode(path)
    If Not node Is Nothing Then Dart_NodeText = Trim$(CStr(node.Text))
End Function

Public Function Dart_RequestXml(ByVal endpoint As String, ByVal key As String, ByVal params As String) As Object
    Dim http As Object, doc As Object, status As String
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 10000, 30000, 60000
    http.Open "GET", DART_BASE & endpoint & ".xml?crtfc_key=" & key & params, False
    http.send
    If http.Status <> 200 Then Err.Raise vbObjectError + 9605, , "DART API HTTP 오류 " & http.Status
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False: doc.setProperty "ProhibitDTD", True
    If Not doc.LoadXML(http.responseText) Then Err.Raise vbObjectError + 9606, , "DART API 응답을 해석하지 못했습니다."
    status = Dart_NodeText(doc, "/result/status")
    Select Case status
        Case "010", "011", "012", "020", "021", "800", "901"
            Err.Raise vbObjectError + 9607, , "DART API 사용 중단 상태 " & status & ". 인증키, 접근 권한 또는 호출 제한을 확인하세요."
    End Select
    Set Dart_RequestXml = doc
End Function

Public Function Dart_QueryCompany(ByVal key As String, ByVal corp As String, ByVal yearValue As Long) As Variant
    Dim result(0 To 18) As Variant, doc As Object, fs As Object, employee As Object
    Dim status As String, params As String, issue As String, revenue As Variant, employeeCount As Variant
    result(0) = corp
    result(12) = CStr(yearValue)
    Set doc = Dart_RequestXml("company", key, "&corp_code=" & corp)
    status = Dart_NodeText(doc, "/result/status")
    If status <> "000" Then result(18) = "기업개황 오류 " & status: Dart_QueryCompany = result: Exit Function
    result(1) = Dart_NodeText(doc, "/result/corp_name")
    result(2) = Dart_NodeText(doc, "/result/corp_name_eng")
    result(3) = Dart_NodeText(doc, "/result/stock_code")
    result(4) = Dart_NodeText(doc, "/result/ceo_nm")
    result(5) = Dart_NodeText(doc, "/result/corp_cls")
    result(6) = Dart_NodeText(doc, "/result/bizr_no")
    result(7) = Dart_NodeText(doc, "/result/jurir_no")
    result(8) = Dart_NodeText(doc, "/result/adres")
    result(9) = Dart_NodeText(doc, "/result/hm_url")
    result(10) = Dart_NodeText(doc, "/result/phn_no")
    result(11) = Dart_NodeText(doc, "/result/induty_code")
    params = "&corp_code=" & corp & "&bsns_year=" & CStr(yearValue) & "&reprt_code=11011"
    Set fs = Dart_RequestXml("fnlttSinglAcntAll", key, params & "&fs_div=OFS")
    revenue = Dart_ExtractRevenue(fs)
    If Len(CStr(revenue(0))) > 0 Then
        result(13) = revenue(0): result(14) = revenue(1): result(15) = revenue(2): result(16) = "별도/개별 (OFS)"
        If Len(CStr(revenue(1))) > 0 And CStr(revenue(1)) <> "KRW" Then issue = "매출 통화 확인 필요: " & CStr(revenue(1))
    Else
        issue = "별도 매출 미확인 (API 상태 " & Dart_NodeText(fs, "/result/status") & ")"
    End If
    Set employee = Dart_RequestXml("empSttus", key, params)
    employeeCount = Dart_ExtractEmployees(employee)
    result(17) = employeeCount(0)
    If Len(CStr(employeeCount(0))) = 0 Then
        If Len(issue) > 0 Then issue = issue & "; "
        issue = issue & "직원 합계 확인 필요 (API 상태 " & Dart_NodeText(employee, "/result/status") & ")"
    ElseIf Len(CStr(employeeCount(1))) > 0 Then
        If Len(issue) > 0 Then issue = issue & "; "
        issue = issue & CStr(employeeCount(1))
    End If
    result(18) = issue
    Dart_QueryCompany = result
End Function

Public Function Dart_ExtractRevenue(ByVal doc As Object) As Variant
    Dim result(0 To 2) As String, node As Object, accountId As String, accountName As String
    Dim rank As Long, best As Long, statement As String, amount As String
    best = 999
    If Dart_NodeText(doc, "/result/status") <> "000" Then Dart_ExtractRevenue = result: Exit Function
    For Each node In doc.SelectNodes("/result/list")
        statement = Dart_NodeText(node, "sj_div")
        If statement = "IS" Or statement = "CIS" Then
            accountId = Dart_NodeText(node, "account_id")
            accountName = Replace(Dart_NodeText(node, "account_nm"), " ", "")
            rank = 999
            Select Case accountId
                Case "ifrs-full_Revenue": rank = 1
                Case "ifrs-full_RevenueFromContractsWithCustomers": rank = 2
                Case "dart_Revenue": rank = 3
            End Select
            If rank = 999 Then
                Select Case accountName
                    Case "매출액": rank = 4
                    Case "매출": rank = 5
                    Case "수익(매출액)": rank = 6
                    Case "영업수익": rank = 7
                    Case "매출액(영업수익)": rank = 8
                    Case "영업수익(매출액)": rank = 9
                    Case "수익": rank = 10
                    Case "순매출액": rank = 11
                    Case "매출및지분법손익": rank = 12
                    Case "영업수익(수익)": rank = 13
                    Case "고객과의계약에서생기는수익": rank = 16
                    Case "고객과의계약에서발생하는수익": rank = 17
                    Case "영업수익(고객과의계약에서생기는수익)": rank = 18
                    Case "매출액(고객과의계약에서생기는수익)": rank = 19
                End Select
            End If
            If statement = "CIS" Then rank = rank * 2 + 1 Else rank = rank * 2
            amount = Replace(Dart_NodeText(node, "thstrm_amount"), ",", "")
            If rank < best And Len(amount) > 0 And IsNumeric(amount) Then
                best = rank: result(0) = amount: result(1) = Dart_NodeText(node, "currency"): result(2) = Dart_NodeText(node, "account_nm")
            End If
        End If
    Next node
    Dart_ExtractRevenue = result
End Function

Private Function Dart_IsTotal(ByVal valueIn As String) As Boolean
    Dim s As String
    s = Replace(valueIn, " ", "")
    Select Case s
        Case "합계", "계", "총계", "전체", "소계", "총합계": Dart_IsTotal = True
        Case Else: Dart_IsTotal = (Right$(s, 2) = "합계" Or Right$(s, 2) = "총계")
    End Select
End Function

Public Function Dart_ExtractEmployees(ByVal doc As Object) As Variant
    Dim result(0 To 1) As String, node As Object, category As Long, best As Long
    Dim sums(1 To 4) As Double, counts(1 To 4) As Long, invalid(1 To 4) As Boolean
    Dim divisionTotal As Boolean, genderTotal As Boolean, amount As String, basis As String
    If Dart_NodeText(doc, "/result/status") <> "000" Then Dart_ExtractEmployees = result: Exit Function
    For Each node In doc.SelectNodes("/result/list")
        divisionTotal = Dart_IsTotal(Dart_NodeText(node, "fo_bbm"))
        genderTotal = Dart_IsTotal(Dart_NodeText(node, "sexdstn"))
        If divisionTotal And genderTotal Then
            category = 1
        ElseIf divisionTotal Then
            category = 2
        ElseIf genderTotal Then
            category = 3
        Else
            category = 4
        End If
        amount = Replace(Dart_NodeText(node, "sm"), ",", "")
        counts(category) = counts(category) + 1
        If Len(amount) = 0 Or Not IsNumeric(amount) Then
            invalid(category) = True
        Else
            sums(category) = sums(category) + CDbl(amount)
        End If
    Next node
    For category = 1 To 4
        If counts(category) > 0 And Not invalid(category) Then best = category: Exit For
    Next category
    If best > 0 Then
        If best = 1 And counts(1) > 1 Then
            result(1) = "직원 총합계 행 중복: 직접 확인 필요"
        Else
            result(0) = Format$(sums(best), "0")
            Select Case best
                Case 1: basis = "총합계행"
                Case 2: basis = "사업부문 합계행"
                Case 3: basis = "부문별 성별합계행"
                Case 4: basis = "상세행 합산"
            End Select
            result(1) = "직원 기준: " & basis
            If best <> 4 And counts(4) > 0 And Not invalid(4) Then
                If sums(4) <> sums(best) Then result(1) = result(1) & "; 상세행 합계와 불일치"
            End If
        End If
    End If
    Dart_ExtractEmployees = result
End Function
