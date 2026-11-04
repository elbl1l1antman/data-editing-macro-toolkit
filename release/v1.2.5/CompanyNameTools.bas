Attribute VB_Name = "CompanyNameTools"
Option Explicit


Public Sub CompanyName_CleanWithOptions()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("기업/사업체명 정리", _
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
    options("operationOrder") = CStr(form.OperationOrder)

    Unload form
    Set CompanyName_GetOptionsFromUser = options
    Exit Function

Fallback:
    Set CompanyName_GetOptionsFromUser = CompanyName_DefaultOptions()
End Function

Public Function CompanyName_DefaultOptions() As Object
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
    options("operationOrder") = CompanyName_DefaultOperationOrder()
    Set CompanyName_DefaultOptions = options
End Function

Private Function CompanyName_DefaultOperationOrder() As String
    CompanyName_DefaultOperationOrder = _
        "cutBeforePossessive,trimOuter,removeParentheticalText,removeBilingualDuplicates,pronounceEnglishWords,removeSpaces," & _
        "normalizeCorpSymbol,normalizeCorpText,moveCorpToFront,moveCorpToEnd,removeCorpMarkers,removeBranchMarkers," & _
        "removeLegalEntityMarkers,removeSpecialChars,addCorpPrefix,addCorpSuffix,spellAlphabetKorean"
End Function

Private Function CompanyName_UserRequestedOperationOrder() As String
    CompanyName_UserRequestedOperationOrder = _
        "removeSpaces,removeBilingualDuplicates,removeCorpMarkers,removeSpecialChars,removeParentheticalText," & _
        "cutBeforePossessive,trimOuter,pronounceEnglishWords,normalizeCorpSymbol,normalizeCorpText,moveCorpToFront,moveCorpToEnd," & _
        "removeBranchMarkers,removeLegalEntityMarkers,addCorpPrefix,addCorpSuffix,spellAlphabetKorean"
End Function

Private Function CompanyName_PreserveParentheticalFirstOrder() As String
    CompanyName_PreserveParentheticalFirstOrder = _
        "trimOuter,removeParentheticalText,removeLegalEntityMarkers,removeBranchMarkers,removeCorpMarkers,removeBilingualDuplicates," & _
        "pronounceEnglishWords,removeSpecialChars,removeSpaces,normalizeCorpSymbol,normalizeCorpText,moveCorpToFront,moveCorpToEnd," & _
        "addCorpPrefix,addCorpSuffix,spellAlphabetKorean"
End Function

Private Sub CompanyName_CleanRangeToCopiedColumns(ByVal target As Range, ByVal options As Object)
    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(target)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 처리할 기업체명/사업체명을 찾지 못했습니다.", vbInformation, "기업/사업체명 정리"
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
    Toolkit_RowProgressStart "기업/사업체명 정리", totalRows, "정리 준비 중"
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
        Toolkit_RowProgressUpdate relRow, totalRows, "기업/사업체명 정리 중: " & CStr(relRow) & "/" & CStr(totalRows) & "행"
    Next relRow

    outputRange.Columns.AutoFit
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    MsgBox changedCount & "개 기업체명/사업체명을 정리했습니다." & vbCrLf & _
           blankPreventedCount & "개 비공란 원본은 공란화를 방지해 원본값을 유지했습니다." & vbCrLf & _
           skippedFormulaCount & "개 수식 셀은 건너뛰었습니다." & vbCrLf & _
           "원본 범위는 변경하지 않았습니다.", vbInformation, "기업/사업체명 정리"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose totalRows
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("기업/사업체명 정리") Then Exit Sub
    MsgBox "기업/사업체명 정리 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "기업/사업체명 정리"
End Sub

' ============================================================
' 프로시저 : DataCheck_CreateCharacterCheckFormula
' 설  명  : 선택한 데이터 범위 오른쪽에 문자 구성 검진 수식을 생성한다.
'           Excel 365의 SEQUENCE를 우선 사용하고, 구버전에서는 Formula로 재시도한다.
' ============================================================
Public Sub DataCheck_CreateCharacterCheckFormula()
    Dim targetRange As Range
    If Not Toolkit_GetRangeFromUserForm("문자 검진 함수 생성", _
        "검진할 데이터 셀 범위를 선택하세요. 헤더가 있으면 제외하고 실제 데이터만 선택합니다." & vbCrLf & _
        "선택 범위 오른쪽에 알파벳/특수기호/정상 판정 수식 열이 생성됩니다.", _
        targetRange) Then Exit Sub

    Dim workRange As Range
    Set workRange = Toolkit_TrimRangeToActualData(targetRange)
    If workRange Is Nothing Then
        MsgBox "선택한 범위에서 검진할 데이터를 찾지 못했습니다.", vbInformation, "문자 검진 함수 생성"
        Exit Sub
    End If

    Dim outputRange As Range
    Set outputRange = CopyRangeToAdjacentColumns(workRange, "문자검진")
    If outputRange Is Nothing Then Exit Sub

    Dim alphaChars As String
    alphaChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Dim specialChars As String
    specialChars = "!@#$%^&*()_+-=[]{}|\" & Chr$(34) & "<>,.?/~" & Chr$(96)

    Dim sourceCell As Range
    Dim outputCell As Range
    Dim formulaText As String
    Dim q As String
    q = Chr$(34)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Toolkit_RowProgressStart "문자 검진 함수 생성", workRange.Rows.Count, "수식 생성 준비 중"

    Dim relRow As Long
    Dim relCol As Long
    For relRow = 1 To workRange.Rows.Count
        For relCol = 1 To workRange.Columns.Count
            Set sourceCell = workRange.Cells(relRow, relCol)
            Set outputCell = outputRange.Cells(relRow, relCol)
            formulaText = "=IF(" & sourceCell.Address(False, False) & "=" & q & q & "," & q & q & _
                ",IF(SUMPRODUCT(--ISNUMBER(FIND(MID(" & sourceCell.Address(False, False) & ",SEQUENCE(LEN(" & sourceCell.Address(False, False) & ")),1)," & q & alphaChars & q & ")))>0," & _
                q & "알파벳 포함" & q & ",IF(SUMPRODUCT(--ISNUMBER(FIND(MID(" & sourceCell.Address(False, False) & ",SEQUENCE(LEN(" & sourceCell.Address(False, False) & ")),1)," & q & specialChars & q & ")))>0," & _
                q & "특수기호 포함" & q & "," & q & "정상" & q & ")))"
            On Error Resume Next
            outputCell.Formula2 = formulaText
            If Err.Number <> 0 Then
                Err.Clear
                outputCell.Formula = formulaText
            End If
            On Error GoTo ErrHandler
        Next relCol
        Toolkit_RowProgressUpdate relRow, workRange.Rows.Count, "문자 검진 수식 생성 중: " & CStr(relRow) & "/" & CStr(workRange.Rows.Count) & "행"
    Next relRow

    outputRange.Columns.AutoFit
    Toolkit_RowProgressClose workRange.Rows.Count
    Application.ScreenUpdating = True
    MsgBox "문자 검진 수식을 " & CStr(outputRange.Cells.Count) & "개 생성했습니다." & vbCrLf & _
           "결과 열에서 알파벳 포함, 특수기호 포함, 정상 여부를 확인하세요.", vbInformation, "문자 검진 함수 생성"
    Exit Sub

ErrHandler:
    Toolkit_RowProgressClose workRange.Rows.Count
    Application.ScreenUpdating = True
    If Toolkit_ShowCancelMessage("문자 검진 함수 생성") Then Exit Sub
    MsgBox "문자 검진 수식 생성 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation, "문자 검진 함수 생성"
End Sub

Public Function CompanyName_CleanOne(ByVal valueIn As String, ByVal options As Object) As String
    Dim textValue As String
    textValue = CStr(valueIn)

    Dim orderText As String
    If options.Exists("operationOrder") Then
        orderText = CStr(options("operationOrder"))
    End If
    If Len(Trim$(orderText)) = 0 Then orderText = CompanyName_DefaultOperationOrder()

    Dim steps As Variant
    steps = Split(orderText, ",")

    Dim i As Long
    For i = LBound(steps) To UBound(steps)
        textValue = CompanyName_ApplyOperation(textValue, Trim$(CStr(steps(i))), options)
    Next i

    CompanyName_CleanOne = textValue
End Function

Private Function CompanyName_ApplyOperation(ByVal textValue As String, ByVal operationName As String, ByVal options As Object) As String
    Select Case operationName
        Case "cutBeforePossessive"
            If CBool(options("cutBeforePossessive")) Then
                Dim cutPos As Long
                cutPos = InStr(1, textValue, "의 ", vbTextCompare)
                If cutPos > 0 Then textValue = Left$(textValue, cutPos - 1)
            End If
        Case "trimOuter"
            If CBool(options("trimOuter")) Then textValue = Trim$(textValue)
        Case "removeParentheticalText"
            If CBool(options("removeParentheticalText")) Then textValue = CompanyName_RemoveParentheticalText(textValue)
        Case "removeBilingualDuplicates"
            If CBool(options("removeBilingualDuplicates")) Then textValue = CompanyName_RemoveBilingualDuplicates(textValue)
        Case "pronounceEnglishWords"
            If CBool(options("pronounceEnglishWords")) Then textValue = CompanyName_PronounceEnglishWords(textValue)
        Case "removeSpaces"
            If CBool(options("removeSpaces")) Then textValue = CompanyName_RemoveSpaces(textValue)
        Case "normalizeCorpSymbol"
            If CBool(options("normalizeCorpSymbol")) Then textValue = CompanyName_NormalizeCorpSymbol(textValue)
        Case "normalizeCorpText"
            If CBool(options("normalizeCorpText")) Then textValue = Replace(textValue, "주식회사", "(주)")
        Case "moveCorpToFront"
            If CBool(options("moveCorpToFront")) Then textValue = CompanyName_MoveCorpMarker(textValue, True)
        Case "moveCorpToEnd"
            If CBool(options("moveCorpToEnd")) Then textValue = CompanyName_MoveCorpMarker(textValue, False)
        Case "removeCorpMarkers"
            If CBool(options("removeCorpMarkers")) Then textValue = CompanyName_RemoveCorpMarkers(textValue)
        Case "removeBranchMarkers"
            If CBool(options("removeBranchMarkers")) Then textValue = CompanyName_RemoveBranchMarkers(textValue)
        Case "removeLegalEntityMarkers"
            If CBool(options("removeLegalEntityMarkers")) Then textValue = CompanyName_RemoveLegalEntityMarkers(textValue)
        Case "removeSpecialChars"
            If CBool(options("removeSpecialChars")) Then
                If Not CBool(options("removeLegalEntityMarkers")) Then textValue = CompanyName_ReplaceBusinessSymbols(textValue)
                textValue = CompanyName_RemoveSpecialChars(textValue)
            End If
        Case "addCorpPrefix"
            If CBool(options("addCorpPrefix")) Then
                If Not CompanyName_HasCorpMarker(textValue) Then textValue = "(주)" & textValue
            End If
        Case "addCorpSuffix"
            If CBool(options("addCorpSuffix")) Then
                If Not CompanyName_HasCorpMarker(textValue) Then textValue = textValue & "(주)"
            End If
        Case "spellAlphabetKorean"
            If CBool(options("spellAlphabetKorean")) Then textValue = CompanyName_SpellAlphabetKorean(textValue)
    End Select

    CompanyName_ApplyOperation = textValue
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
    result = Replace(result, "+", "플러스")
    result = Replace(result, "＋", "플러스")

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
