Attribute VB_Name = "KosisAge"
Option Explicit

Public Function Kosis_FindTopAge(ByVal metadata As Collection) As Long
    Dim entry As Object, re As Object, matches As Object, found As Boolean, n As Long
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^([0-9]+)세\s*이상$"
    For Each entry In metadata
        If CStr(entry("axis")) = "YRE" Then
            If re.Test(Trim$(CStr(entry("name")))) Then
                Set matches = re.Execute(Trim$(CStr(entry("name"))))
                n = CLng(matches(0).SubMatches(0))
                If found Then Err.Raise 5, , "최상위 연령 묶음이 여러 개입니다."
                Kosis_FindTopAge = n: found = True
            End If
        End If
    Next entry
    If Not found Then Err.Raise 5, , "최상위 연령 묶음을 찾지 못했습니다."
End Function

Public Function Kosis_AgeSourceCodes(ByVal metadata As Collection, ByVal bands As Collection, ByVal topAge As Long) As Collection
    Dim output As New Collection, entry As Object, band As Object, re As Object, matches As Object
    Dim age As Long, name As String, seen As Object, i As Long, required As Boolean
    Set seen = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^([0-9]+)세\s*(이상)?$"
    For Each entry In metadata
        If CStr(entry("axis")) = "YRE" Then
            name = Trim$(CStr(entry("name")))
            If re.Test(name) Then
                Set matches = re.Execute(name): age = CLng(matches(0).SubMatches(0))
                If age >= 0 And age <= topAge Then
                    If age = topAge And Len(CStr(matches(0).SubMatches(1))) = 0 Then Err.Raise 5, , "최상위 묶음 연령이 아닙니다."
                    If age < topAge And Len(CStr(matches(0).SubMatches(1))) > 0 Then Err.Raise 5, , "개별 연령이 아닌 묶음 자료입니다."
                    If seen.Exists(CStr(age)) Then Err.Raise 5, , "동일 연령 코드가 중복되었습니다."
                    seen.Add CStr(age), True
                    required = False
                    For Each band In bands
                        If age >= CLng(band("from")) And age <= CLng(band("to")) Then required = True: Exit For
                    Next band
                    If required Then output.Add CStr(entry("code"))
                End If
            End If
        End If
    Next entry
    For i = 0 To topAge
        If Not seen.Exists(CStr(i)) Then Err.Raise 5, , "원자료의 개별 연령이 누락되었습니다."
    Next i
    If output.Count = 0 Then Err.Raise 5, , "선택한 연령이 없습니다."
    Set Kosis_AgeSourceCodes = output
End Function

' The final ordinal is the whole open-ended KOSIS category, never a single age.
Public Function Kosis_ParseAgeBands(ByVal spec As String, ByVal topAge As Long, ByRef uncovered As Long) As Collection
    Dim lines As Variant, line As Variant, parts As Variant, lower As String, upper As String
    Dim startOp As String, endOp As String, name As String, age As Long, first As Long, last As Long
    Dim band As Object, bands As New Collection, tailAtStart As Boolean, tailAtEnd As Boolean
    If topAge < 1 Or topAge > 200 Then Err.Raise 5, , "최상위 연령 구간을 확인하세요."
    lines = Split(Replace(spec, vbCrLf, vbLf), vbLf)
    For Each line In lines
        If Len(Trim$(CStr(line))) > 0 Then
            parts = Split(CStr(line), vbTab)
            If UBound(parts) < 5 Then Err.Raise 5, , "연령 구간 정보가 누락되었습니다."
            lower = Trim$(CStr(parts(1))): startOp = CStr(parts(2))
            upper = Trim$(CStr(parts(3))): endOp = CStr(parts(4)): name = Trim$(CStr(parts(5)))
            If Len(name) = 0 Then Err.Raise 5, , "연령 구간 이름을 입력하세요."
            If startOp <> "이상" And startOp <> "초과" Then Err.Raise 5, , "시작 조건을 확인하세요."
            If endOp <> "미만" And endOp <> "이하" Then Err.Raise 5, , "종료 조건을 확인하세요."
            If Len(lower) = 0 And Len(upper) = 0 Then Err.Raise 5, , "연령 구간 경계를 입력하세요."
            Kosis_ValidateAgeBound lower
            Kosis_ValidateAgeBound upper
            tailAtStart = Kosis_AgeWithin(topAge, lower, startOp, upper, endOp)
            tailAtEnd = Kosis_AgeWithin(1000001, lower, startOp, upper, endOp)
            If tailAtStart <> tailAtEnd Then Err.Raise 5, , "원자료의 최상위 연령 묶음은 나누어 합산할 수 없습니다."
            first = -1: last = -1
            For age = 0 To topAge
                If Kosis_AgeWithin(age, lower, startOp, upper, endOp) Then
                    If first < 0 Then first = age
                    last = age
                End If
            Next age
            If first < 0 Then Err.Raise 5, , "선택한 연령에 해당하지 않는 구간이 있습니다."
            Set band = CreateObject("Scripting.Dictionary")
            band.Add "name", name: band.Add "from", first: band.Add "to", last
            bands.Add band
        End If
    Next line
    If Not Kosis_ValidateAgeBands(bands, topAge, uncovered) Then Err.Raise 5, , "연령 구간을 확인하세요."
    Set Kosis_ParseAgeBands = bands
End Function

Private Sub Kosis_ValidateAgeBound(ByVal value As String)
    Dim n As Double
    If Len(value) = 0 Then Exit Sub
    If Not IsNumeric(value) Then Err.Raise 5, , "연령 경계는 정수여야 합니다."
    n = CDbl(value)
    If n <> Fix(n) Or n < 0 Or n > 1000000 Then Err.Raise 5, , "연령 경계는 0 이상의 정수여야 합니다."
End Sub

Private Function Kosis_AgeWithin(ByVal age As Long, ByVal lower As String, ByVal startOp As String, _
                                 ByVal upper As String, ByVal endOp As String) As Boolean
    If Len(lower) > 0 Then
        If startOp = "이상" And age < CLng(lower) Then Exit Function
        If startOp = "초과" And age <= CLng(lower) Then Exit Function
    End If
    If Len(upper) > 0 Then
        If endOp = "이하" And age > CLng(upper) Then Exit Function
        If endOp = "미만" And age >= CLng(upper) Then Exit Function
    End If
    Kosis_AgeWithin = True
End Function

Public Function Kosis_AggregateAges(ByVal records As Collection, ByVal metadata As Collection, ByVal bands As Collection, _
    ByVal regions As Collection, ByVal period As String, ByVal item As String, ByVal topAge As Long, _
    ByRef review As Collection) As Collection
    Dim codes As Object, ages As Object, buckets As Object, seen As Object, selected As Object, labels As Object
    Dim re As Object, matches As Object, entry As Object, row As Object, bucket As Object, value As Variant
    Dim age As Long, bandIndex As Long, i As Long, region As Variant, code As String, key As String, groupCode As String
    Dim result As New Collection, raw As String, uncovered As Long, number As Double, flagged As Boolean
    If Not Kosis_ValidateAgeBands(bands, topAge, uncovered) Then Err.Raise 5, , "연령 구간을 확인하세요."
    Set codes = CreateObject("Scripting.Dictionary"): Set ages = CreateObject("Scripting.Dictionary")
    Set buckets = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    Set selected = CreateObject("Scripting.Dictionary"): Set labels = CreateObject("Scripting.Dictionary")
    Set review = New Collection
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^([0-9]+)세\s*(이상)?$"
    For Each entry In metadata
        If CStr(entry("axis")) = "YRE" Then
            code = CStr(entry("code")): raw = Trim$(CStr(entry("name")))
            If labels.Exists(code) Then Err.Raise 5, , "연령 메타정보가 중복되었습니다."
            labels.Add code, raw
            If Left$(code, 10) <> "KOSIS_AGE_" And re.Test(raw) Then
                Set matches = re.Execute(raw): age = CLng(matches(0).SubMatches(0))
                If age < 0 Or age > topAge Then Err.Raise 5, , "원자료의 연령 범위를 확인하세요."
                If age = topAge And Len(CStr(matches(0).SubMatches(1))) = 0 Then Err.Raise 5, , "최상위 묶음 연령이 아닙니다."
                If age < topAge And Len(CStr(matches(0).SubMatches(1))) > 0 Then Err.Raise 5, , "개별 연령이 아닌 묶음 자료입니다."
                If ages.Exists(CStr(age)) Then Err.Raise 5, , "동일 연령 코드가 중복되었습니다."
                ages.Add CStr(age), True: codes.Add code, age
            End If
        End If
    Next entry
    For age = 0 To topAge
        If Not ages.Exists(CStr(age)) Then Err.Raise 5, , "원자료의 개별 연령이 누락되었습니다."
    Next age
    For Each region In regions
        If selected.Exists(CStr(region)) Then Err.Raise 5, , "지역이 중복 선택되었습니다."
        selected.Add CStr(region), True
        For i = 1 To bands.Count
            key = Kosis_AgeBucketKey(CStr(region), i)
            Set bucket = CreateObject("Scripting.Dictionary")
            bucket.Add "region", CStr(region): bucket.Add "index", i
            bucket.Add "count", 0&: bucket.Add "sum", 0#: bucket.Add "problem", False: bucket.Add "status", False
            buckets.Add key, bucket
        Next i
    Next region
    If selected.Count = 0 Then Err.Raise 5, , "지역을 선택하세요."
    For Each row In records
        If Not row.Exists("C_A") Or Not row.Exists("C_YRE") Or Not row.Exists("ITEM") Or Not row.Exists("period") Or Not row.Exists("value") Then Err.Raise 5, , "연령 자료에 필수 정보가 없습니다."
        raw = CStr(row("C_A")): code = CStr(row("C_YRE"))
        If Not selected.Exists(raw) Or Not codes.Exists(code) Then Err.Raise 5, , "선택하지 않은 지역 또는 연령이 반환되었습니다."
        If CStr(row("period")) <> period Or CStr(row("ITEM")) <> item Then Err.Raise 5, , "서로 다른 시점 또는 항목이 섞였습니다."
        key = Len(raw) & ":" & raw & Len(code) & ":" & code
        If seen.Exists(key) Then Err.Raise 5, , "동일한 연령 자료가 중복되었습니다."
        seen.Add key, True: age = CLng(codes(code)): bandIndex = 0
        For i = 1 To bands.Count
            Set entry = bands(i)
            If age >= CLng(entry("from")) And age <= CLng(entry("to")) Then bandIndex = i: Exit For
        Next i
        If bandIndex > 0 Then
            Set bucket = buckets(Kosis_AgeBucketKey(raw, bandIndex))
            bucket("count") = CLng(bucket("count")) + 1
            raw = CStr(row("value"))
            If IsNumeric(raw) Then
                number = CDbl(raw)
                If number < 0 Or number > 9007199254740991# Then
                    bucket("problem") = True
                Else
                    bucket("sum") = CDbl(bucket("sum")) + number
                End If
            Else
                bucket("problem") = True
            End If
            If row.Exists("observationXml") Then
                raw = CStr(row("observationXml"))
                If InStr(1, raw, "OBS_STATUS", vbTextCompare) > 0 Then bucket("status") = True
            End If
            If row.Exists("status") Then
                If Len(CStr(row("status"))) > 0 Then bucket("status") = True
            End If
        End If
    Next row
    For i = 1 To bands.Count
        groupCode = "KOSIS_AGE_" & Format$(i, "000")
        Set entry = bands(i)
        If labels.Exists(groupCode) Then
            If CStr(labels(groupCode)) <> CStr(entry("name")) Then Err.Raise 5, , "연령 구간 코드가 충돌합니다."
        Else
            Set row = CreateObject("Scripting.Dictionary")
            row.Add "axis", "YRE": row.Add "axisName", "연령구간": row.Add "axisOrder", "2"
            row.Add "code", groupCode: row.Add "name", CStr(entry("name")): row.Add "parent", "": row.Add "unit", "명"
            metadata.Add row
        End If
    Next i
    For Each region In regions
        For i = 1 To bands.Count
            Set entry = bands(i): Set bucket = buckets(Kosis_AgeBucketKey(CStr(region), i))
            flagged = CBool(bucket("problem")) Or CLng(bucket("count")) <> CLng(entry("to")) - CLng(entry("from")) + 1
            If CDbl(bucket("sum")) > 9007199254740991# Then flagged = True
            Set row = CreateObject("Scripting.Dictionary")
            row.Add "C_A", CStr(region): row.Add "C_YRE", "KOSIS_AGE_" & Format$(i, "000")
            row.Add "ITEM", item: row.Add "period", period
            If flagged Then
                row.Add "value", "확인 필요"
                review.Add "지역 " & CStr(region) & " / " & CStr(entry("name")) & ": 연령 자료 누락 또는 통계부호"
            Else
                row.Add "value", CStr(bucket("sum"))
            End If
            If CBool(bucket("status")) Then review.Add "지역 " & CStr(region) & " / " & CStr(entry("name")) & ": 원자료 상태부호 확인"
            result.Add row
        Next i
    Next region
    Set Kosis_AggregateAges = result
End Function

Private Function Kosis_AgeBucketKey(ByVal region As String, ByVal bandIndex As Long) As String
    Kosis_AgeBucketKey = Len(region) & ":" & region & ":" & CStr(bandIndex)
End Function
