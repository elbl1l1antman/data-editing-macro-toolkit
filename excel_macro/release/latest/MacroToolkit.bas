Attribute VB_Name = "MacroToolkit"
Option Explicit

Public Const TOOLKIT_VERSION As String = "0.1.0"

Private Const VERSION_SHEET As String = "_MacroToolkit"

Public Sub Toolkit_About()
    MsgBox "Data Editing Macro Toolkit" & vbCrLf & _
           "Version: " & TOOLKIT_VERSION & vbCrLf & vbCrLf & _
           "Current features:" & vbCrLf & _
           "- Convert text-formatted numbers to real numbers" & vbCrLf & _
           "- Normalize mixed Korean phone number formats", _
           vbInformation, "Macro Toolkit"
End Sub

Public Sub Toolkit_SetupWorkbook()
    Dim ws As Worksheet
    Set ws = EnsureVersionSheet()
    ws.Range("A1").Value = "Data Editing Macro Toolkit"
    ws.Range("A2").Value = "Version"
    ws.Range("B2").Value = TOOLKIT_VERSION
    ws.Range("A3").Value = "Updated"
    ws.Range("B3").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("A5").Value = "Available macros"
    ws.Range("A6").Value = "Toolkit_About"
    ws.Range("A7").Value = "ConvertTextNumbersInSelection"
    ws.Range("A8").Value = "ConvertTextNumbersInUsedRange"
    ws.Range("A9").Value = "ConvertTextNumbersInColumns"
    ws.Range("A10").Value = "NormalizePhonesInSelection"
    ws.Range("A11").Value = "NormalizePhonesInColumns"
    ws.Columns("A:B").AutoFit
End Sub

Public Sub ConvertTextNumbersInSelection()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cells to convert first.", vbExclamation, "Macro Toolkit"
        Exit Sub
    End If
    ConvertTextNumbersInRange Selection
End Sub

Public Sub ConvertTextNumbersInUsedRange()
    ConvertTextNumbersInRange ActiveSheet.UsedRange
End Sub

Public Sub ConvertTextNumbersInColumns()
    Dim columnText As String
    columnText = InputBox("Enter target column letters separated by commas." & vbCrLf & _
                          "Example: A,C,F:H", "Convert Text Numbers")
    If Len(Trim$(columnText)) = 0 Then Exit Sub
    ConvertTextNumbersInRange ColumnsFromInput(ActiveSheet, columnText)
End Sub

Public Sub NormalizePhonesInSelection()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the phone number cells first.", vbExclamation, "Macro Toolkit"
        Exit Sub
    End If
    NormalizePhonesInRange Selection
End Sub

Public Sub NormalizePhonesInColumns()
    Dim columnText As String
    columnText = InputBox("Enter phone number column letters separated by commas." & vbCrLf & _
                          "Example: B,D,F:G", "Normalize Phone Numbers")
    If Len(Trim$(columnText)) = 0 Then Exit Sub
    NormalizePhonesInRange ColumnsFromInput(ActiveSheet, columnText)
End Sub

Private Sub ConvertTextNumbersInRange(ByVal target As Range)
    Dim cell As Range
    Dim convertedCount As Long
    Dim skippedFormulaCount As Long
    Dim newValue As Variant
    
    Application.ScreenUpdating = False
    For Each cell In target.Cells
        If cell.HasFormula Then
            skippedFormulaCount = skippedFormulaCount + 1
        ElseIf TryNumberFromText(cell.Value, newValue) Then
            cell.Value = newValue
            convertedCount = convertedCount + 1
        End If
    Next cell
    Application.ScreenUpdating = True
    
    MsgBox convertedCount & " cell(s) converted." & vbCrLf & _
           skippedFormulaCount & " formula cell(s) skipped.", _
           vbInformation, "Convert Text Numbers"
End Sub

Private Sub NormalizePhonesInRange(ByVal target As Range)
    Dim cell As Range
    Dim changedCount As Long
    Dim skippedFormulaCount As Long
    Dim normalized As String
    
    Application.ScreenUpdating = False
    For Each cell In target.Cells
        If cell.HasFormula Then
            skippedFormulaCount = skippedFormulaCount + 1
        ElseIf NormalizeKoreanPhone(CStr(cell.Value), normalized) Then
            If CStr(cell.Value) <> normalized Then
                cell.NumberFormat = "@"
                cell.Value = normalized
                changedCount = changedCount + 1
            End If
        End If
    Next cell
    Application.ScreenUpdating = True
    
    MsgBox changedCount & " phone number(s) normalized." & vbCrLf & _
           skippedFormulaCount & " formula cell(s) skipped.", _
           vbInformation, "Normalize Phone Numbers"
End Sub

Private Function TryNumberFromText(ByVal valueIn As Variant, ByRef valueOut As Variant) As Boolean
    Dim textValue As String
    Dim negativeValue As Boolean
    Dim percentValue As Boolean
    
    If IsError(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If IsNumeric(valueIn) And VarType(valueIn) <> vbString Then Exit Function
    
    textValue = Trim$(CStr(valueIn))
    If Len(textValue) = 0 Then Exit Function
    
    If Left$(textValue, 1) = "'" Then textValue = Mid$(textValue, 2)
    textValue = Replace(textValue, ChrW(160), "")
    textValue = Replace(textValue, " ", "")
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, "\", "")
    textValue = Replace(textValue, "₩", "")
    
    If Left$(textValue, 1) = "(" And Right$(textValue, 1) = ")" Then
        negativeValue = True
        textValue = Mid$(textValue, 2, Len(textValue) - 2)
    End If
    
    If Right$(textValue, 1) = "%" Then
        percentValue = True
        textValue = Left$(textValue, Len(textValue) - 1)
    End If
    
    If Len(textValue) = 0 Or Not IsNumeric(textValue) Then Exit Function
    
    valueOut = CDbl(textValue)
    If negativeValue Then valueOut = -CDbl(valueOut)
    If percentValue Then valueOut = CDbl(valueOut) / 100
    TryNumberFromText = True
End Function

Private Function NormalizeKoreanPhone(ByVal valueIn As String, ByRef valueOut As String) As Boolean
    Dim source As String
    Dim digits As String
    
    source = Trim$(valueIn)
    If Len(source) = 0 Then Exit Function
    
    digits = DigitsOnly(source)
    If Len(digits) = 0 Then Exit Function
    
    If Left$(digits, 4) = "0082" Then digits = "0" & Mid$(digits, 5)
    If Left$(digits, 2) = "82" Then digits = "0" & Mid$(digits, 3)
    
    If Left$(digits, 2) = "02" Then
        Select Case Len(digits)
            Case 9
                valueOut = "02-" & Mid$(digits, 3, 3) & "-" & Right$(digits, 4)
            Case 10
                valueOut = "02-" & Mid$(digits, 3, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function
        End Select
    ElseIf Left$(digits, 3) Like "01#" Then
        Select Case Len(digits)
            Case 10
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 3) & "-" & Right$(digits, 4)
            Case 11
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function
        End Select
    ElseIf Left$(digits, 3) Like "0##" Then
        Select Case Len(digits)
            Case 10
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 3) & "-" & Right$(digits, 4)
            Case 11
                valueOut = Left$(digits, 3) & "-" & Mid$(digits, 4, 4) & "-" & Right$(digits, 4)
            Case Else
                Exit Function
        End Select
    Else
        Exit Function
    End If
    
    NormalizeKoreanPhone = True
End Function

Private Function DigitsOnly(ByVal source As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(source)
        ch = Mid$(source, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

Private Function ColumnsFromInput(ByVal ws As Worksheet, ByVal columnText As String) As Range
    Dim parts() As String
    Dim part As Variant
    Dim rangeText As String
    Dim outputRange As Range
    
    parts = Split(columnText, ",")
    For Each part In parts
        rangeText = Trim$(CStr(part))
        If Len(rangeText) > 0 Then
            If outputRange Is Nothing Then
                Set outputRange = ws.Range(rangeText & ":" & rangeText)
            Else
                Set outputRange = Union(outputRange, ws.Range(rangeText & ":" & rangeText))
            End If
        End If
    Next part
    
    Set ColumnsFromInput = outputRange
End Function

Private Function EnsureVersionSheet() As Worksheet
    On Error Resume Next
    Set EnsureVersionSheet = ThisWorkbook.Worksheets(VERSION_SHEET)
    On Error GoTo 0
    
    If EnsureVersionSheet Is Nothing Then
        Set EnsureVersionSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        EnsureVersionSheet.Name = VERSION_SHEET
    End If
End Function
