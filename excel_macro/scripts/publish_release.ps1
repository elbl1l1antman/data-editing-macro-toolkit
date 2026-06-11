param(
    [string]$Version,
    [string]$ProjectRoot
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Set-ZipTextEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [string]$Text
    )
    foreach ($name in @($EntryName, $EntryName.Replace('/', '\'))) {
        $old = $Archive.GetEntry($name)
        if ($null -ne $old) { $old.Delete() }
    }
    $entry = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Get-ZipTextEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName
    )
    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) { $entry = $Archive.GetEntry($EntryName.Replace('/', '\')) }
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Add-ToolkitButtonMenuUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmToolkitButtonMenu"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "메뉴 선택"
    $form.Properties.Item("Width").Value = 460
    $form.Properties.Item("Height").Value = 420
    $form.Properties.Item("ScrollBars").Value = 2

    $lblGuide = $designer.Controls.Add("Forms.Label.1", "lblGuide", $true)
    $lblGuide.Left = 12
    $lblGuide.Top = 10
    $lblGuide.Width = 420
    $lblGuide.Height = 38
    $lblGuide.WordWrap = $true

    for ($i = 1; $i -le 20; $i++) {
        $btn = $designer.Controls.Add("Forms.CommandButton.1", "cmdOption$i", $true)
        $btn.Left = 16
        $btn.Top = 54 + (($i - 1) * 30)
        $btn.Width = 400
        $btn.Height = 24
        $btn.Visible = $false
    }

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 336
    $cmdCancel.Top = 360
    $cmdCancel.Width = 80
    $cmdCancel.Height = 26

    $handlers = ""
    for ($i = 1; $i -le 20; $i++) {
        $handlers += @"
Private Sub cmdOption$i`_Click()
    PickOption $i
End Sub

"@
    }

    $formCode = @"
Option Explicit
Public SelectedKey As String

Public Sub Configure(ByVal titleText As String, ByVal guideText As String, ByVal optionsText As String, ByVal defaultKey As String)
    Caption = titleText
    lblGuide.Caption = guideText
    SelectedKey = ""

    Dim i As Long
    For i = 1 To 20
        Controls("cmdOption" & CStr(i)).Visible = False
        Controls("cmdOption" & CStr(i)).Tag = ""
    Next i

    Dim lines As Variant
    lines = Split(optionsText, vbLf)
    Dim visibleCount As Long
    For i = LBound(lines) To UBound(lines)
        If Len(Trim`$(CStr(lines(i)))) > 0 And visibleCount < 20 Then
            Dim parts As Variant
            parts = Split(CStr(lines(i)), vbTab)
            If UBound(parts) >= 1 Then
                visibleCount = visibleCount + 1
                With Controls("cmdOption" & CStr(visibleCount))
                    .Caption = CStr(parts(1))
                    .Tag = CStr(parts(0))
                    .Visible = True
                    If CStr(parts(0)) = defaultKey Then .Default = True
                End With
            End If
        End If
    Next i

    cmdCancel.Top = 60 + visibleCount * 30
    Height = Application.WorksheetFunction.Min(560, cmdCancel.Top + 62)
    ScrollHeight = cmdCancel.Top + 56
End Sub

Private Sub PickOption(ByVal idx As Long)
    SelectedKey = CStr(Controls("cmdOption" & CStr(idx)).Tag)
    Hide
End Sub

$handlers
Private Sub cmdCancel_Click()
    SelectedKey = ""
    Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        SelectedKey = ""
    End If
End Sub
"@
    $form.CodeModule.AddFromString($formCode)
}

function Add-ToolkitCheckListMenuUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmToolkitCheckListMenu"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "항목 선택"
    $form.Properties.Item("Width").Value = 560
    $form.Properties.Item("Height").Value = 520
    $form.Properties.Item("ScrollBars").Value = 2

    $lblGuide = $designer.Controls.Add("Forms.Label.1", "lblGuide", $true)
    $lblGuide.Left = 12
    $lblGuide.Top = 10
    $lblGuide.Width = 450
    $lblGuide.Height = 38
    $lblGuide.WordWrap = $true

    for ($i = 1; $i -le 30; $i++) {
        $chk = $designer.Controls.Add("Forms.CheckBox.1", "chkOption$i", $true)
        $chk.Left = 16
        $chk.Top = 54 + (($i - 1) * 24)
        $chk.Width = 430
        $chk.Height = 20
        $chk.Visible = $false
    }

    $cmdAll = $designer.Controls.Add("Forms.CommandButton.1", "cmdAll", $true)
    $cmdAll.Caption = "모두 선택"
    $cmdAll.Left = 16
    $cmdAll.Top = 420
    $cmdAll.Width = 80
    $cmdAll.Height = 26

    $cmdNone = $designer.Controls.Add("Forms.CommandButton.1", "cmdNone", $true)
    $cmdNone.Caption = "모두 해제"
    $cmdNone.Left = 104
    $cmdNone.Top = 420
    $cmdNone.Width = 80
    $cmdNone.Height = 26

    $cmdOK = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOK.Caption = "확인"
    $cmdOK.Left = 292
    $cmdOK.Top = 420
    $cmdOK.Width = 70
    $cmdOK.Height = 26

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 372
    $cmdCancel.Top = 420
    $cmdCancel.Width = 70
    $cmdCancel.Height = 26

    $formCode = @'
Option Explicit
Public SelectedKeys As String
Private optionCount As Long

Public Sub Configure(ByVal titleText As String, ByVal guideText As String, ByVal optionsText As String, ByVal defaultCsv As String)
    Caption = titleText
    lblGuide.Caption = guideText
    SelectedKeys = ""
    optionCount = 0

    Dim i As Long
    For i = 1 To 30
        Controls("chkOption" & CStr(i)).Visible = False
        Controls("chkOption" & CStr(i)).Value = False
        Controls("chkOption" & CStr(i)).Tag = ""
    Next i

    Dim defaults As Object
    Set defaults = CreateObject("Scripting.Dictionary")
    Dim defaultParts As Variant
    defaultParts = Split(defaultCsv, ",")
    For i = LBound(defaultParts) To UBound(defaultParts)
        If Len(Trim$(CStr(defaultParts(i)))) > 0 Then defaults(Trim$(CStr(defaultParts(i)))) = True
    Next i

    Dim lines As Variant
    lines = Split(optionsText, vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) > 0 And optionCount < 30 Then
            Dim parts As Variant
            parts = Split(CStr(lines(i)), vbTab)
            If UBound(parts) >= 1 Then
                optionCount = optionCount + 1
                With Controls("chkOption" & CStr(optionCount))
                    .Caption = CStr(parts(1))
                    .Tag = CStr(parts(0))
                    .Value = defaults.Exists(CStr(parts(0)))
                    .Visible = True
                End With
            End If
        End If
    Next i

    Dim buttonTop As Long
    buttonTop = 62 + optionCount * 24
    cmdAll.Top = buttonTop
    cmdNone.Top = buttonTop
    cmdOK.Top = buttonTop
    cmdCancel.Top = buttonTop
    Height = Application.WorksheetFunction.Min(620, buttonTop + 64)
    ScrollHeight = buttonTop + 58
End Sub

Private Sub cmdAll_Click()
    Dim i As Long
    For i = 1 To optionCount
        Controls("chkOption" & CStr(i)).Value = True
    Next i
End Sub

Private Sub cmdNone_Click()
    Dim i As Long
    For i = 1 To optionCount
        Controls("chkOption" & CStr(i)).Value = False
    Next i
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    SelectedKeys = ""
    For i = 1 To optionCount
        If Controls("chkOption" & CStr(i)).Value Then
            If Len(SelectedKeys) > 0 Then SelectedKeys = SelectedKeys & ","
            SelectedKeys = SelectedKeys & CStr(Controls("chkOption" & CStr(i)).Tag)
        End If
    Next i
    Hide
End Sub

Private Sub cmdCancel_Click()
    SelectedKeys = ""
    Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        SelectedKeys = ""
    End If
End Sub
'@
    $form.CodeModule.AddFromString($formCode)
}

function Add-PopulationListifyPreviewUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmPopulationListifyPreview"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "모집단 구성표 리스트화 미리보기"
    $form.Properties.Item("Width").Value = 760
    $form.Properties.Item("Height").Value = 560

    $lblHelp = $designer.Controls.Add("Forms.Label.1", "lblHelp", $true)
    $lblHelp.Caption = "탐지된 집계표 구조와 변환 예시를 확인한 뒤 실행하세요. 합계 행/열은 결과에서 제외됩니다."
    $lblHelp.Left = 12
    $lblHelp.Top = 10
    $lblHelp.Width = 720
    $lblHelp.Height = 24

    $lblSummary = $designer.Controls.Add("Forms.Label.1", "lblSummary", $true)
    $lblSummary.Caption = "탐지 결과"
    $lblSummary.Left = 12
    $lblSummary.Top = 42
    $lblSummary.Width = 180
    $lblSummary.Height = 16

    $txtSummary = $designer.Controls.Add("Forms.TextBox.1", "txtSummary", $true)
    $txtSummary.Left = 12
    $txtSummary.Top = 62
    $txtSummary.Width = 720
    $txtSummary.Height = 142
    $txtSummary.MultiLine = $true
    $txtSummary.ScrollBars = 2
    $txtSummary.Locked = $true

    $lblColumnAxis = $designer.Controls.Add("Forms.Label.1", "lblColumnAxis", $true)
    $lblColumnAxis.Caption = "가로축 열 이름"
    $lblColumnAxis.Left = 12
    $lblColumnAxis.Top = 220
    $lblColumnAxis.Width = 100
    $lblColumnAxis.Height = 18

    $txtColumnAxisName = $designer.Controls.Add("Forms.TextBox.1", "txtColumnAxisName", $true)
    $txtColumnAxisName.Left = 116
    $txtColumnAxisName.Top = 216
    $txtColumnAxisName.Width = 150
    $txtColumnAxisName.Height = 22

    $lblValueName = $designer.Controls.Add("Forms.Label.1", "lblValueName", $true)
    $lblValueName.Caption = "값 열 이름"
    $lblValueName.Left = 286
    $lblValueName.Top = 220
    $lblValueName.Width = 80
    $lblValueName.Height = 18

    $txtValueName = $designer.Controls.Add("Forms.TextBox.1", "txtValueName", $true)
    $txtValueName.Left = 366
    $txtValueName.Top = 216
    $txtValueName.Width = 120
    $txtValueName.Height = 22

    $lblSheetName = $designer.Controls.Add("Forms.Label.1", "lblSheetName", $true)
    $lblSheetName.Caption = "결과 시트"
    $lblSheetName.Left = 506
    $lblSheetName.Top = 220
    $lblSheetName.Width = 70
    $lblSheetName.Height = 18

    $txtOutputSheetName = $designer.Controls.Add("Forms.TextBox.1", "txtOutputSheetName", $true)
    $txtOutputSheetName.Left = 576
    $txtOutputSheetName.Top = 216
    $txtOutputSheetName.Width = 156
    $txtOutputSheetName.Height = 22

    $lblPreview = $designer.Controls.Add("Forms.Label.1", "lblPreview", $true)
    $lblPreview.Caption = "변환 예시"
    $lblPreview.Left = 12
    $lblPreview.Top = 252
    $lblPreview.Width = 180
    $lblPreview.Height = 16

    $txtPreview = $designer.Controls.Add("Forms.TextBox.1", "txtPreview", $true)
    $txtPreview.Left = 12
    $txtPreview.Top = 272
    $txtPreview.Width = 720
    $txtPreview.Height = 196
    $txtPreview.MultiLine = $true
    $txtPreview.ScrollBars = 3
    $txtPreview.Locked = $true

    $cmdRun = $designer.Controls.Add("Forms.CommandButton.1", "cmdRun", $true)
    $cmdRun.Caption = "실행"
    $cmdRun.Left = 548
    $cmdRun.Top = 486
    $cmdRun.Width = 86
    $cmdRun.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 646
    $cmdCancel.Top = 486
    $cmdCancel.Width = 86
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Confirmed As Boolean

Public Sub Configure(ByVal summaryText As String, ByVal previewText As String, _
                     ByVal defaultColumnAxisName As String, ByVal defaultValueName As String, _
                     ByVal defaultOutputSheetName As String)
    Confirmed = False
    txtSummary.Text = summaryText
    txtPreview.Text = previewText
    txtColumnAxisName.Text = defaultColumnAxisName
    txtValueName.Text = defaultValueName
    txtOutputSheetName.Text = defaultOutputSheetName
End Sub

Public Property Get ColumnAxisName() As String
    ColumnAxisName = Trim$(txtColumnAxisName.Text)
End Property

Public Property Get ValueName() As String
    ValueName = Trim$(txtValueName.Text)
End Property

Public Property Get OutputSheetName() As String
    OutputSheetName = Trim$(txtOutputSheetName.Text)
End Property

Private Sub cmdRun_Click()
    If Len(Trim$(txtColumnAxisName.Text)) = 0 Then
        MsgBox "가로축 열 이름을 입력하세요.", vbExclamation, "모집단 구성표 리스트화"
        txtColumnAxisName.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtValueName.Text)) = 0 Then
        MsgBox "값 열 이름을 입력하세요.", vbExclamation, "모집단 구성표 리스트화"
        txtValueName.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtOutputSheetName.Text)) = 0 Then
        MsgBox "결과 시트 이름을 입력하세요.", vbExclamation, "모집단 구성표 리스트화"
        txtOutputSheetName.SetFocus
        Exit Sub
    End If

    Confirmed = True
    Hide
End Sub

Private Sub cmdCancel_Click()
    Confirmed = False
    Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Confirmed = False
End Sub
'@
    $form.CodeModule.AddFromString($formCode)
}

function Add-SpssRebuildUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmSpssRebuildOptions"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "SPSS 표/배너 재구성"
    $form.Properties.Item("Width").Value = 760
    $form.Properties.Item("Height").Value = 500

    $lblHelp = $designer.Controls.Add("Forms.Label.1", "lblHelp", $true)
    $lblHelp.Caption = "표는 선택 후 버튼이나 드래그로 순서를 바꾸고, [삭제/복원]으로 삭제 여부를 표시하세요. Shift+클릭/Ctrl+클릭 다중 선택을 지원합니다."
    $lblHelp.Left = 12
    $lblHelp.Top = 10
    $lblHelp.Width = 720
    $lblHelp.Height = 28

    $lblTables = $designer.Controls.Add("Forms.Label.1", "lblTables", $true)
    $lblTables.Caption = "집계표 순서/삭제"
    $lblTables.Left = 12
    $lblTables.Top = 44
    $lblTables.Width = 390
    $lblTables.Height = 16

    $lstTables = $designer.Controls.Add("Forms.ListBox.1", "lstTables", $true)
    $lstTables.Left = 12
    $lstTables.Top = 64
    $lstTables.Width = 390
    $lstTables.Height = 286
    $lstTables.ColumnCount = 3
    $lstTables.ColumnWidths = "305 pt;0 pt;60 pt"
    $lstTables.MultiSelect = 2

    $cmdTableTop = $designer.Controls.Add("Forms.CommandButton.1", "cmdTableTop", $true)
    $cmdTableTop.Caption = "맨 위"
    $cmdTableTop.Left = 12
    $cmdTableTop.Top = 356
    $cmdTableTop.Width = 58
    $cmdTableTop.Height = 24

    $cmdTableUp = $designer.Controls.Add("Forms.CommandButton.1", "cmdTableUp", $true)
    $cmdTableUp.Caption = "위로"
    $cmdTableUp.Left = 76
    $cmdTableUp.Top = 356
    $cmdTableUp.Width = 58
    $cmdTableUp.Height = 24

    $cmdTableDown = $designer.Controls.Add("Forms.CommandButton.1", "cmdTableDown", $true)
    $cmdTableDown.Caption = "아래로"
    $cmdTableDown.Left = 140
    $cmdTableDown.Top = 356
    $cmdTableDown.Width = 58
    $cmdTableDown.Height = 24

    $cmdTableBottom = $designer.Controls.Add("Forms.CommandButton.1", "cmdTableBottom", $true)
    $cmdTableBottom.Caption = "맨 아래"
    $cmdTableBottom.Left = 204
    $cmdTableBottom.Top = 356
    $cmdTableBottom.Width = 66
    $cmdTableBottom.Height = 24

    $cmdToggleTableDelete = $designer.Controls.Add("Forms.CommandButton.1", "cmdToggleTableDelete", $true)
    $cmdToggleTableDelete.Caption = "삭제/복원"
    $cmdToggleTableDelete.Left = 278
    $cmdToggleTableDelete.Top = 356
    $cmdToggleTableDelete.Width = 82
    $cmdToggleTableDelete.Height = 24

    $lblBanners = $designer.Controls.Add("Forms.Label.1", "lblBanners", $true)
    $lblBanners.Caption = "삭제할 배너"
    $lblBanners.Left = 422
    $lblBanners.Top = 44
    $lblBanners.Width = 310
    $lblBanners.Height = 16

    $lstBanners = $designer.Controls.Add("Forms.ListBox.1", "lstBanners", $true)
    $lstBanners.Left = 422
    $lstBanners.Top = 64
    $lstBanners.Width = 310
    $lstBanners.Height = 286
    $lstBanners.ColumnCount = 2
    $lstBanners.ColumnWidths = "280 pt;0 pt"
    $lstBanners.MultiSelect = 2

    $cmdUp = $designer.Controls.Add("Forms.CommandButton.1", "cmdBannerUp", $true)
    $cmdUp.Caption = "위로"
    $cmdUp.Left = 422
    $cmdUp.Top = 356
    $cmdUp.Width = 58
    $cmdUp.Height = 24

    $cmdDown = $designer.Controls.Add("Forms.CommandButton.1", "cmdBannerDown", $true)
    $cmdDown.Caption = "아래로"
    $cmdDown.Left = 486
    $cmdDown.Top = 356
    $cmdDown.Width = 58
    $cmdDown.Height = 24

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "삭제 상태가 표시된 표와 선택된 배너만 제거됩니다. 결과 시트는 원본을 유지한 채 새로 만들고, 표 사이에는 공백 행 1개를 둡니다."
    $lblTip.Left = 12
    $lblTip.Top = 392
    $lblTip.Width = 720
    $lblTip.Height = 34

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 546
    $cmdOk.Top = 452
    $cmdOk.Width = 86
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 646
    $cmdCancel.Top = 452
    $cmdCancel.Width = 86
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean
Public DeleteTablesText As String
Public DeleteBannersText As String
Public BannerOrderText As String
Public TableOrderText As String

Private DraggingTables As Boolean
Private DragTableStartIndex As Long

Private Sub UserForm_Initialize()
    Cancelled = True
    DraggingTables = False
    DragTableStartIndex = -1
End Sub

Public Sub InitializeOptions(ByVal tables As Object, ByVal banners As Object)
    Dim i As Long
    Dim item As Variant

    lstTables.Clear
    For i = 1 To tables.Count
        item = tables(i)
        lstTables.AddItem CStr(i) & ". " & CStr(item(0))
        lstTables.List(lstTables.ListCount - 1, 1) = CStr(i)
        lstTables.List(lstTables.ListCount - 1, 2) = ""
    Next i

    lstBanners.Clear
    For i = 1 To banners.Count
        lstBanners.AddItem CStr(i) & ". " & CStr(banners(i))
        lstBanners.List(lstBanners.ListCount - 1, 1) = CStr(i)
    Next i
End Sub

Private Sub cmdTableTop_Click()
    MoveSelectedTableRowsToEdge -1
End Sub

Private Sub cmdTableUp_Click()
    MoveSelectedTableRows -1
End Sub

Private Sub cmdTableDown_Click()
    MoveSelectedTableRows 1
End Sub

Private Sub cmdTableBottom_Click()
    MoveSelectedTableRowsToEdge 1
End Sub

Private Sub cmdToggleTableDelete_Click()
    ToggleSelectedTableDelete
End Sub

Private Sub lstTables_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        DraggingTables = True
        DragTableStartIndex = ListIndexFromY(lstTables, Y)
    End If
End Sub

Private Sub lstTables_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not DraggingTables Then Exit Sub
    DraggingTables = False
    If Button <> 1 Then Exit Sub

    Dim targetIndex As Long
    targetIndex = ListIndexFromY(lstTables, Y)
    If targetIndex >= 0 And targetIndex <> DragTableStartIndex Then MoveSelectedTablesToIndex targetIndex
End Sub

Private Function ListIndexFromY(ByVal listBox As MSForms.ListBox, ByVal yPos As Single) As Long
    Dim idx As Long
    idx = listBox.TopIndex + CLng(Fix(yPos / 12))
    If idx < 0 Then idx = 0
    If idx >= listBox.ListCount Then idx = listBox.ListCount - 1
    ListIndexFromY = idx
End Function

Private Sub MoveSelectedTableRowsToEdge(ByVal direction As Long)
    Dim guard As Long
    If direction < 0 Then
        Do While FirstSelectedIndex(lstTables) > 0 And guard < lstTables.ListCount
            MoveSelectedTableRows -1
            guard = guard + 1
        Loop
    Else
        Do While LastSelectedIndex(lstTables) < lstTables.ListCount - 1 And guard < lstTables.ListCount
            MoveSelectedTableRows 1
            guard = guard + 1
        Loop
    End If
End Sub

Private Sub MoveSelectedTablesToIndex(ByVal targetIndex As Long)
    Dim guard As Long
    If targetIndex < FirstSelectedIndex(lstTables) Then
        Do While FirstSelectedIndex(lstTables) > targetIndex And guard < lstTables.ListCount
            MoveSelectedTableRows -1
            guard = guard + 1
        Loop
    ElseIf targetIndex > LastSelectedIndex(lstTables) Then
        Do While LastSelectedIndex(lstTables) < targetIndex And guard < lstTables.ListCount
            MoveSelectedTableRows 1
            guard = guard + 1
        Loop
    End If
End Sub

Private Sub MoveSelectedTableRows(ByVal direction As Long)
    Dim i As Long
    If lstTables.ListCount = 0 Then Exit Sub

    If direction < 0 Then
        For i = 1 To lstTables.ListCount - 1
            If lstTables.Selected(i) And Not lstTables.Selected(i - 1) Then SwapTableRows i, i - 1
        Next i
    Else
        For i = lstTables.ListCount - 2 To 0 Step -1
            If lstTables.Selected(i) And Not lstTables.Selected(i + 1) Then SwapTableRows i, i + 1
        Next i
    End If
End Sub

Private Sub SwapTableRows(ByVal a As Long, ByVal b As Long)
    Dim text0 As String, id0 As String, status0 As String
    Dim selectedA As Boolean, selectedB As Boolean

    text0 = lstTables.List(a, 0)
    id0 = lstTables.List(a, 1)
    status0 = lstTables.List(a, 2)
    selectedA = lstTables.Selected(a)
    selectedB = lstTables.Selected(b)

    lstTables.List(a, 0) = lstTables.List(b, 0)
    lstTables.List(a, 1) = lstTables.List(b, 1)
    lstTables.List(a, 2) = lstTables.List(b, 2)
    lstTables.List(b, 0) = text0
    lstTables.List(b, 1) = id0
    lstTables.List(b, 2) = status0

    lstTables.Selected(a) = selectedB
    lstTables.Selected(b) = selectedA
End Sub

Private Sub ToggleSelectedTableDelete()
    Dim i As Long
    For i = 0 To lstTables.ListCount - 1
        If lstTables.Selected(i) Then
            If Len(CStr(lstTables.List(i, 2))) > 0 Then
                lstTables.List(i, 2) = ""
            Else
                lstTables.List(i, 2) = "삭제"
            End If
        End If
    Next i
End Sub

Private Sub cmdBannerUp_Click()
    MoveSelectedBanner -1
End Sub

Private Sub cmdBannerDown_Click()
    MoveSelectedBanner 1
End Sub

Private Sub MoveSelectedBanner(ByVal direction As Long)
    Dim i As Long
    If direction < 0 Then
        For i = 1 To lstBanners.ListCount - 1
            If lstBanners.Selected(i) Then SwapBannerRows i, i - 1
        Next i
    Else
        For i = lstBanners.ListCount - 2 To 0 Step -1
            If lstBanners.Selected(i) Then SwapBannerRows i, i + 1
        Next i
    End If
End Sub

Private Sub SwapBannerRows(ByVal a As Long, ByVal b As Long)
    Dim text0 As String, id0 As String
    Dim selectedA As Boolean, selectedB As Boolean

    text0 = lstBanners.List(a, 0)
    id0 = lstBanners.List(a, 1)
    selectedA = lstBanners.Selected(a)
    selectedB = lstBanners.Selected(b)

    lstBanners.List(a, 0) = lstBanners.List(b, 0)
    lstBanners.List(a, 1) = lstBanners.List(b, 1)
    lstBanners.List(b, 0) = text0
    lstBanners.List(b, 1) = id0

    lstBanners.Selected(a) = selectedB
    lstBanners.Selected(b) = selectedA
End Sub

Private Sub cmdOK_Click()
    DeleteTablesText = DeletedTableIds()
    DeleteBannersText = SelectedIds(lstBanners)
    TableOrderText = AllIds(lstTables)
    BannerOrderText = AllIds(lstBanners)
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Function SelectedIds(ByVal listBox As MSForms.ListBox) As String
    Dim result As String
    Dim i As Long
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            If Len(result) > 0 Then result = result & ","
            result = result & CStr(listBox.List(i, 1))
        End If
    Next i
    SelectedIds = result
End Function

Private Function AllIds(ByVal listBox As MSForms.ListBox) As String
    Dim result As String
    Dim i As Long
    For i = 0 To listBox.ListCount - 1
        If Len(result) > 0 Then result = result & ","
        result = result & CStr(listBox.List(i, 1))
    Next i
    AllIds = result
End Function

Private Function DeletedTableIds() As String
    Dim result As String
    Dim i As Long
    For i = 0 To lstTables.ListCount - 1
        If CStr(lstTables.List(i, 2)) = "삭제" Then
            If Len(result) > 0 Then result = result & ","
            result = result & CStr(lstTables.List(i, 1))
        End If
    Next i
    DeletedTableIds = result
End Function

Private Function FirstSelectedIndex(ByVal listBox As MSForms.ListBox) As Long
    Dim i As Long
    FirstSelectedIndex = -1
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            FirstSelectedIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function LastSelectedIndex(ByVal listBox As MSForms.ListBox) As Long
    Dim i As Long
    LastSelectedIndex = -1
    For i = listBox.ListCount - 1 To 0 Step -1
        If listBox.Selected(i) Then
            LastSelectedIndex = i
            Exit Function
        End If
    Next i
End Function
'@

    $form.CodeModule.AddFromString($formCode)
}

function Add-CompanyNameOptionsUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmCompanyNameOptions"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "기업체명/사업체명 정리 옵션"
    $form.Properties.Item("Width").Value = 500
    $form.Properties.Item("Height").Value = 430

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "적용할 정리 규칙을 선택하세요."
    $lblTitle.Left = 12
    $lblTitle.Top = 10
    $lblTitle.Width = 460
    $lblTitle.Height = 18

    $labels = @(
        @("chkTrimOuter", "앞뒤 공백 제거", 12, 42, $false),
        @("chkRemoveSpaces", "모든 공백 제거", 12, 68, $true),
        @("chkNormalizeCorpSymbol", "㈜ 를 (주) 로 통일", 12, 94, $false),
        @("chkNormalizeCorpText", "주식회사 를 (주) 로 통일", 12, 120, $false),
        @("chkMoveCorpToFront", "(주)를 회사명 앞으로 이동", 12, 146, $false),
        @("chkMoveCorpToEnd", "(주)를 회사명 뒤로 이동", 12, 172, $false),
        @("chkRemoveCorpMarkers", "(주), ㈜, 주식회사 삭제", 250, 42, $true),
        @("chkAddCorpPrefix", "회사형태가 없으면 앞에 (주) 추가", 250, 68, $false),
        @("chkAddCorpSuffix", "회사형태가 없으면 뒤에 (주) 추가", 250, 94, $false),
        @("chkCutBeforePossessive", "'의 ' 앞부분만 남기기", 250, 120, $true),
        @("chkPronounceEnglishWords", "영어 단어를 한글 발음으로 변환", 250, 146, $false),
        @("chkSpellAlphabetKorean", "알파벳을 한글 읽기로 변환", 250, 172, $true)
    )

    foreach ($item in $labels) {
        $chk = $designer.Controls.Add("Forms.CheckBox.1", $item[0], $true)
        $chk.Caption = $item[1]
        $chk.Left = [int]$item[2]
        $chk.Top = [int]$item[3]
        $chk.Width = 280
        $chk.Height = 20
        $chk.Value = [bool]$item[4]
    }

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "기본값은 제공된 기업이름정리.xlsm의 체크 상태를 따릅니다. 결과는 원본 오른쪽 복사본 열에 생성됩니다."
    $lblTip.Left = 12
    $lblTip.Top = 214
    $lblTip.Width = 520
    $lblTip.Height = 42

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 346
    $cmdOk.Top = 330
    $cmdOk.Width = 86
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 446
    $cmdCancel.Top = 330
    $cmdCancel.Width = 86
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = True
End Sub

Public Property Get TrimOuter() As Boolean
    TrimOuter = CBool(chkTrimOuter.Value)
End Property

Public Property Get RemoveSpaces() As Boolean
    RemoveSpaces = CBool(chkRemoveSpaces.Value)
End Property

Public Property Get NormalizeCorpSymbol() As Boolean
    NormalizeCorpSymbol = CBool(chkNormalizeCorpSymbol.Value)
End Property

Public Property Get NormalizeCorpText() As Boolean
    NormalizeCorpText = CBool(chkNormalizeCorpText.Value)
End Property

Public Property Get MoveCorpToFront() As Boolean
    MoveCorpToFront = CBool(chkMoveCorpToFront.Value)
End Property

Public Property Get MoveCorpToEnd() As Boolean
    MoveCorpToEnd = CBool(chkMoveCorpToEnd.Value)
End Property

Public Property Get RemoveCorpMarkers() As Boolean
    RemoveCorpMarkers = CBool(chkRemoveCorpMarkers.Value)
End Property

Public Property Get AddCorpPrefix() As Boolean
    AddCorpPrefix = CBool(chkAddCorpPrefix.Value)
End Property

Public Property Get AddCorpSuffix() As Boolean
    AddCorpSuffix = CBool(chkAddCorpSuffix.Value)
End Property

Public Property Get CutBeforePossessive() As Boolean
    CutBeforePossessive = CBool(chkCutBeforePossessive.Value)
End Property

Public Property Get PronounceEnglishWords() As Boolean
    PronounceEnglishWords = CBool(chkPronounceEnglishWords.Value)
End Property

Public Property Get SpellAlphabetKorean() As Boolean
    SpellAlphabetKorean = CBool(chkSpellAlphabetKorean.Value)
End Property

Private Sub cmdOK_Click()
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancelled = True
    Cancel = True
    Me.Hide
End Sub
'@

    $form.CodeModule.AddFromString($formCode)
}

function Add-BizStatusFieldsOptionsUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmBizStatusFieldsOptions"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "사업자등록 상태조회 반환값 선택"
    $form.Properties.Item("Width").Value = 560
    $form.Properties.Item("Height").Value = 430

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "조회 결과로 추가할 열을 선택하세요."
    $lblTitle.Left = 12
    $lblTitle.Top = 12
    $lblTitle.Width = 500
    $lblTitle.Height = 22

    $chkAll = $designer.Controls.Add("Forms.CheckBox.1", "chkAll", $true)
    $chkAll.Caption = "모두 포함"
    $chkAll.Left = 12
    $chkAll.Top = 42
    $chkAll.Width = 120
    $chkAll.Height = 20
    $chkAll.Value = $true

    $items = @(
        @("chkStatusCode", "응답상태코드", 12, 78),
        @("chkRequestedNo", "요청 사업자등록번호", 12, 104),
        @("chkBNo", "사업자등록번호", 12, 130),
        @("chkBStt", "납세자상태", 12, 156),
        @("chkBSttCd", "납세자상태코드", 12, 182),
        @("chkTaxType", "과세유형", 12, 208),
        @("chkTaxTypeCd", "과세유형코드", 12, 234),
        @("chkEndDt", "폐업일자", 280, 78),
        @("chkUtccYn", "단위과세전환폐업여부", 280, 104),
        @("chkTaxTypeChangeDt", "과세유형전환일자", 280, 130),
        @("chkInvoiceApplyDt", "세금계산서적용일자", 280, 156),
        @("chkRbfTaxType", "직전과세유형", 280, 182),
        @("chkRbfTaxTypeCd", "직전과세유형코드", 280, 208)
    )

    foreach ($item in $items) {
        $chk = $designer.Controls.Add("Forms.CheckBox.1", $item[0], $true)
        $chk.Caption = $item[1]
        $chk.Left = [int]$item[2]
        $chk.Top = [int]$item[3]
        $chk.Width = 240
        $chk.Height = 20
        $chk.Value = $true
    }

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "선택한 항목 수만큼 원본 오른쪽에 열이 추가됩니다. 오류가 난 행은 선택된 항목 중 응답상태코드, 요청번호, 납세자상태/과세유형 칸에 확인 메시지를 표시합니다."
    $lblTip.Left = 12
    $lblTip.Top = 282
    $lblTip.Width = 510
    $lblTip.Height = 48

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 350
    $cmdOk.Top = 348
    $cmdOk.Width = 80
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 442
    $cmdCancel.Top = 348
    $cmdCancel.Width = 80
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = True
End Sub

Private Sub chkAll_Click()
    SetAllFields CBool(chkAll.Value)
End Sub

Private Sub SetAllFields(ByVal valueIn As Boolean)
    chkStatusCode.Value = valueIn
    chkRequestedNo.Value = valueIn
    chkBNo.Value = valueIn
    chkBStt.Value = valueIn
    chkBSttCd.Value = valueIn
    chkTaxType.Value = valueIn
    chkTaxTypeCd.Value = valueIn
    chkEndDt.Value = valueIn
    chkUtccYn.Value = valueIn
    chkTaxTypeChangeDt.Value = valueIn
    chkInvoiceApplyDt.Value = valueIn
    chkRbfTaxType.Value = valueIn
    chkRbfTaxTypeCd.Value = valueIn
End Sub

Public Property Get SelectedKeys() As String
    Dim result As String
    AddKey result, chkStatusCode.Value, "status_code"
    AddKey result, chkRequestedNo.Value, "requested_b_no"
    AddKey result, chkBNo.Value, "b_no"
    AddKey result, chkBStt.Value, "b_stt"
    AddKey result, chkBSttCd.Value, "b_stt_cd"
    AddKey result, chkTaxType.Value, "tax_type"
    AddKey result, chkTaxTypeCd.Value, "tax_type_cd"
    AddKey result, chkEndDt.Value, "end_dt"
    AddKey result, chkUtccYn.Value, "utcc_yn"
    AddKey result, chkTaxTypeChangeDt.Value, "tax_type_change_dt"
    AddKey result, chkInvoiceApplyDt.Value, "invoice_apply_dt"
    AddKey result, chkRbfTaxType.Value, "rbf_tax_type"
    AddKey result, chkRbfTaxTypeCd.Value, "rbf_tax_type_cd"
    SelectedKeys = result
End Property

Private Sub AddKey(ByRef result As String, ByVal includeField As Boolean, ByVal key As String)
    If includeField Then
        If Len(result) > 0 Then result = result & ","
        result = result & key
    End If
End Sub

Private Sub cmdOK_Click()
    If Len(SelectedKeys) = 0 Then
        MsgBox "반환받을 값을 하나 이상 선택하세요.", vbExclamation, Me.Caption
        Exit Sub
    End If
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancelled = True
    Cancel = True
    Me.Hide
End Sub
'@

    $form.CodeModule.AddFromString($formCode)
}

function Add-NaverLocalBuildOptionsUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmNaverLocalBuildOptions"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "네이버 지역검색 신규 리스트"
    $form.Properties.Item("Width").Value = 460
    $form.Properties.Item("Height").Value = 300

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "선택한 행정동마다 업종을 검색해 신규 업체 리스트를 만듭니다."
    $lblTitle.Left = 12
    $lblTitle.Top = 12
    $lblTitle.Width = 420
    $lblTitle.Height = 20

    $lblIndustry = $designer.Controls.Add("Forms.Label.1", "lblIndustry", $true)
    $lblIndustry.Caption = "업종/검색어"
    $lblIndustry.Left = 12
    $lblIndustry.Top = 48
    $lblIndustry.Width = 110
    $lblIndustry.Height = 18

    $txtIndustry = $designer.Controls.Add("Forms.TextBox.1", "txtIndustry", $true)
    $txtIndustry.Left = 128
    $txtIndustry.Top = 44
    $txtIndustry.Width = 290
    $txtIndustry.Height = 22
    $txtIndustry.Text = "서점"

    $lblDisplay = $designer.Controls.Add("Forms.Label.1", "lblDisplay", $true)
    $lblDisplay.Caption = "행정동별 결과 수"
    $lblDisplay.Left = 12
    $lblDisplay.Top = 84
    $lblDisplay.Width = 110
    $lblDisplay.Height = 18

    $txtDisplay = $designer.Controls.Add("Forms.TextBox.1", "txtDisplay", $true)
    $txtDisplay.Left = 128
    $txtDisplay.Top = 80
    $txtDisplay.Width = 60
    $txtDisplay.Height = 22
    $txtDisplay.Text = "5"

    $lblDelay = $designer.Controls.Add("Forms.Label.1", "lblDelay", $true)
    $lblDelay.Caption = "요청 간격(초)"
    $lblDelay.Left = 220
    $lblDelay.Top = 84
    $lblDelay.Width = 90
    $lblDelay.Height = 18

    $txtDelay = $designer.Controls.Add("Forms.TextBox.1", "txtDelay", $true)
    $txtDelay.Left = 318
    $txtDelay.Top = 80
    $txtDelay.Width = 60
    $txtDelay.Height = 22
    $txtDelay.Text = "0.4"

    $lblSort = $designer.Controls.Add("Forms.Label.1", "lblSort", $true)
    $lblSort.Caption = "정렬"
    $lblSort.Left = 12
    $lblSort.Top = 120
    $lblSort.Width = 110
    $lblSort.Height = 18

    $cmbSort = $designer.Controls.Add("Forms.ComboBox.1", "cmbSort", $true)
    $cmbSort.Left = 128
    $cmbSort.Top = 116
    $cmbSort.Width = 140
    $cmbSort.Height = 22

    $lblSido = $designer.Controls.Add("Forms.Label.1", "lblSido", $true)
    $lblSido.Caption = "광역시도 필터"
    $lblSido.Left = 12
    $lblSido.Top = 154
    $lblSido.Width = 110
    $lblSido.Height = 18

    $txtSido = $designer.Controls.Add("Forms.TextBox.1", "txtSido", $true)
    $txtSido.Left = 128
    $txtSido.Top = 150
    $txtSido.Width = 120
    $txtSido.Height = 22

    $lblSigungu = $designer.Controls.Add("Forms.Label.1", "lblSigungu", $true)
    $lblSigungu.Caption = "시군구 필터"
    $lblSigungu.Left = 260
    $lblSigungu.Top = 154
    $lblSigungu.Width = 82
    $lblSigungu.Height = 18

    $txtSigungu = $designer.Controls.Add("Forms.TextBox.1", "txtSigungu", $true)
    $txtSigungu.Left = 344
    $txtSigungu.Top = 150
    $txtSigungu.Width = 74
    $txtSigungu.Height = 22

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "내장 행정동 마스터 또는 현재 통합문서의 행정동_마스터 시트를 사용합니다. 필터를 비우면 전체 행정동을 대상으로 합니다."
    $lblTip.Left = 12
    $lblTip.Top = 188
    $lblTip.Width = 420
    $lblTip.Height = 44

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 250
    $cmdOk.Top = 232
    $cmdOk.Width = 80
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 342
    $cmdCancel.Top = 232
    $cmdCancel.Width = 80
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = True
    cmbSort.AddItem "comment"
    cmbSort.AddItem "random"
    cmbSort.Value = "comment"
End Sub

Public Property Get IndustryText() As String
    IndustryText = Trim$(txtIndustry.Value)
End Property

Public Property Get DisplayCount() As Long
    DisplayCount = CLng(Val(txtDisplay.Value))
    If DisplayCount < 1 Then DisplayCount = 1
    If DisplayCount > 5 Then DisplayCount = 5
End Property

Public Property Get DelaySeconds() As Double
    DelaySeconds = CDbl(Val(txtDelay.Value))
    If DelaySeconds < 0 Then DelaySeconds = 0
End Property

Public Property Get SortMode() As String
    SortMode = CStr(cmbSort.Value)
    If SortMode <> "random" Then SortMode = "comment"
End Property

Public Property Get SidoFilter() As String
    SidoFilter = Trim$(txtSido.Value)
End Property

Public Property Get SigunguFilter() As String
    SigunguFilter = Trim$(txtSigungu.Value)
End Property

Private Sub cmdOK_Click()
    If Len(IndustryText) = 0 Then
        MsgBox "업종/검색어를 입력하세요.", vbExclamation, Me.Caption
        Exit Sub
    End If
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancelled = True
    Cancel = True
    Me.Hide
End Sub
'@

    $form.CodeModule.AddFromString($formCode)
}

function Add-NaverLocalVerifyOptionsUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmNaverLocalVerifyOptions"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "네이버 지역검색 주소 검증"
    $form.Properties.Item("Width").Value = 560
    $form.Properties.Item("Height").Value = 430

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "분석에 사용할 데이터와 범위를 선택하세요."
    $lblTitle.Left = 12
    $lblTitle.Top = 12
    $lblTitle.Width = 520
    $lblTitle.Height = 24

    $chkUseName = $designer.Controls.Add("Forms.CheckBox.1", "chkUseName", $true)
    $chkUseName.Caption = "업체명 사용"
    $chkUseName.Left = 12
    $chkUseName.Top = 46
    $chkUseName.Width = 100
    $chkUseName.Height = 20
    $chkUseName.Value = $true

    $txtNameRange = $designer.Controls.Add("Forms.TextBox.1", "txtNameRange", $true)
    $txtNameRange.Left = 120
    $txtNameRange.Top = 44
    $txtNameRange.Width = 300
    $txtNameRange.Height = 22

    $cmdPickName = $designer.Controls.Add("Forms.CommandButton.1", "cmdPickName", $true)
    $cmdPickName.Caption = "범위 선택"
    $cmdPickName.Left = 432
    $cmdPickName.Top = 42
    $cmdPickName.Width = 90
    $cmdPickName.Height = 26

    $chkUseAddress = $designer.Controls.Add("Forms.CheckBox.1", "chkUseAddress", $true)
    $chkUseAddress.Caption = "주소 사용"
    $chkUseAddress.Left = 12
    $chkUseAddress.Top = 82
    $chkUseAddress.Width = 100
    $chkUseAddress.Height = 20
    $chkUseAddress.Value = $true

    $txtAddressRange = $designer.Controls.Add("Forms.TextBox.1", "txtAddressRange", $true)
    $txtAddressRange.Left = 120
    $txtAddressRange.Top = 80
    $txtAddressRange.Width = 300
    $txtAddressRange.Height = 22

    $cmdPickAddress = $designer.Controls.Add("Forms.CommandButton.1", "cmdPickAddress", $true)
    $cmdPickAddress.Caption = "범위 선택"
    $cmdPickAddress.Left = 432
    $cmdPickAddress.Top = 78
    $cmdPickAddress.Width = 90
    $cmdPickAddress.Height = 26

    $lblMode = $designer.Controls.Add("Forms.Label.1", "lblMode", $true)
    $lblMode.Caption = "검색 방식"
    $lblMode.Left = 12
    $lblMode.Top = 126
    $lblMode.Width = 110
    $lblMode.Height = 18

    $cmbMode = $designer.Controls.Add("Forms.ComboBox.1", "cmbMode", $true)
    $cmbMode.Left = 128
    $cmbMode.Top = 122
    $cmbMode.Width = 220
    $cmbMode.Height = 22

    $chkHeader = $designer.Controls.Add("Forms.CheckBox.1", "chkHeader", $true)
    $chkHeader.Caption = "첫 행은 헤더"
    $chkHeader.Left = 370
    $chkHeader.Top = 124
    $chkHeader.Width = 100
    $chkHeader.Height = 20
    $chkHeader.Value = $true

    $lblDisplay = $designer.Controls.Add("Forms.Label.1", "lblDisplay", $true)
    $lblDisplay.Caption = "행별 후보 수"
    $lblDisplay.Left = 12
    $lblDisplay.Top = 162
    $lblDisplay.Width = 110
    $lblDisplay.Height = 18

    $txtDisplay = $designer.Controls.Add("Forms.TextBox.1", "txtDisplay", $true)
    $txtDisplay.Left = 128
    $txtDisplay.Top = 158
    $txtDisplay.Width = 60
    $txtDisplay.Height = 22
    $txtDisplay.Text = "5"

    $lblDelay = $designer.Controls.Add("Forms.Label.1", "lblDelay", $true)
    $lblDelay.Caption = "요청 간격(초)"
    $lblDelay.Left = 220
    $lblDelay.Top = 162
    $lblDelay.Width = 90
    $lblDelay.Height = 18

    $txtDelay = $designer.Controls.Add("Forms.TextBox.1", "txtDelay", $true)
    $txtDelay.Left = 318
    $txtDelay.Top = 158
    $txtDelay.Width = 60
    $txtDelay.Height = 22
    $txtDelay.Text = "0.4"

    $lblSort = $designer.Controls.Add("Forms.Label.1", "lblSort", $true)
    $lblSort.Caption = "정렬"
    $lblSort.Left = 12
    $lblSort.Top = 198
    $lblSort.Width = 110
    $lblSort.Height = 18

    $cmbSort = $designer.Controls.Add("Forms.ComboBox.1", "cmbSort", $true)
    $cmbSort.Left = 128
    $cmbSort.Top = 194
    $cmbSort.Width = 140
    $cmbSort.Height = 22

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "업체명만 또는 주소만 선택할 수 있습니다. 둘 다 선택하면 두 범위는 같은 시트와 같은 행 수여야 합니다."
    $lblTip.Left = 12
    $lblTip.Top = 236
    $lblTip.Width = 520
    $lblTip.Height = 54

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 350
    $cmdOk.Top = 332
    $cmdOk.Width = 80
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 446
    $cmdCancel.Top = 332
    $cmdCancel.Width = 80
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean
Private mNameRange As Object
Private mAddressRange As Object

Private Sub UserForm_Initialize()
    Cancelled = True
    cmbMode.AddItem "업체명+주소"
    cmbMode.AddItem "업체명만"
    cmbMode.AddItem "주소만"
    cmbMode.Value = "업체명+주소"
    cmbSort.AddItem "comment"
    cmbSort.AddItem "random"
    cmbSort.Value = "comment"
End Sub

Private Sub cmdPickName_Click()
    On Error Resume Next
    Set mNameRange = Application.InputBox("업체명 범위를 선택하세요.", Me.Caption, Type:=8)
    If Err.Number = 0 And Not mNameRange Is Nothing Then
        txtNameRange.Value = mNameRange.Address(External:=True)
        chkUseName.Value = True
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub cmdPickAddress_Click()
    On Error Resume Next
    Set mAddressRange = Application.InputBox("주소 범위를 선택하세요.", Me.Caption, Type:=8)
    If Err.Number = 0 And Not mAddressRange Is Nothing Then
        txtAddressRange.Value = mAddressRange.Address(External:=True)
        chkUseAddress.Value = True
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Property Get QueryMode() As String
    If UseName And Not UseAddress Then
        QueryMode = "name_only"
    ElseIf UseAddress And Not UseName Then
        QueryMode = "address_only"
    Else
        Select Case cmbMode.Value
            Case "업체명만"
                QueryMode = "name_only"
            Case "주소만"
                QueryMode = "address_only"
            Case Else
                QueryMode = "name_address"
        End Select
    End If
End Property

Public Property Get UseName() As Boolean
    UseName = CBool(chkUseName.Value)
End Property

Public Property Get UseAddress() As Boolean
    UseAddress = CBool(chkUseAddress.Value)
End Property

Public Property Get NameRange() As Object
    Set NameRange = mNameRange
End Property

Public Property Get AddressRange() As Object
    Set AddressRange = mAddressRange
End Property

Public Property Get HasHeader() As Boolean
    HasHeader = CBool(chkHeader.Value)
End Property

Public Property Get DisplayCount() As Long
    DisplayCount = CLng(Val(txtDisplay.Value))
    If DisplayCount < 1 Then DisplayCount = 1
    If DisplayCount > 5 Then DisplayCount = 5
End Property

Public Property Get DelaySeconds() As Double
    DelaySeconds = CDbl(Val(txtDelay.Value))
    If DelaySeconds < 0 Then DelaySeconds = 0
End Property

Public Property Get SortMode() As String
    SortMode = CStr(cmbSort.Value)
    If SortMode <> "random" Then SortMode = "comment"
End Property

Private Sub cmdOK_Click()
    If Not UseName And Not UseAddress Then
        MsgBox "업체명 또는 주소 중 하나 이상을 선택하세요.", vbExclamation, Me.Caption
        Exit Sub
    End If

    If UseName And mNameRange Is Nothing Then
        If Len(Trim$(txtNameRange.Value)) > 0 Then
            On Error Resume Next
            Set mNameRange = Application.Range(Trim$(txtNameRange.Value))
            On Error GoTo 0
        End If
    End If
    If UseAddress And mAddressRange Is Nothing Then
        If Len(Trim$(txtAddressRange.Value)) > 0 Then
            On Error Resume Next
            Set mAddressRange = Application.Range(Trim$(txtAddressRange.Value))
            On Error GoTo 0
        End If
    End If

    If UseName And mNameRange Is Nothing Then
        MsgBox "업체명 사용을 선택했습니다. 업체명 범위를 지정하세요.", vbExclamation, Me.Caption
        Exit Sub
    End If
    If UseAddress And mAddressRange Is Nothing Then
        MsgBox "주소 사용을 선택했습니다. 주소 범위를 지정하세요.", vbExclamation, Me.Caption
        Exit Sub
    End If

    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancelled = True
    Cancel = True
    Me.Hide
End Sub
'@

    $form.CodeModule.AddFromString($formCode)
}

function Add-AdminDongMasterWorksheet {
    param(
        $Workbook,
        [string]$CsvPath
    )

    if (-not (Test-Path -LiteralPath $CsvPath)) {
        throw "Admin dong master CSV not found: $CsvPath"
    }

    foreach ($sheet in @($Workbook.Worksheets)) {
        if ($sheet.Name -eq "_NaverLocalDongMaster") {
            $sheet.Delete()
            break
        }
    }

    $ws = $Workbook.Worksheets.Add()
    $ws.Name = "_NaverLocalDongMaster"

    $rows = Import-Csv -LiteralPath $CsvPath -Encoding UTF8
    $headers = @("광역시도", "시군구", "읍면동", "검색주소", "사용여부")
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }

    $r = 2
    foreach ($row in $rows) {
        $ws.Cells.Item($r, 1).Value2 = $row."광역시도"
        $ws.Cells.Item($r, 2).Value2 = $row."시군구"
        $ws.Cells.Item($r, 3).Value2 = $row."읍면동"
        $ws.Cells.Item($r, 4).Value2 = $row."검색주소"
        $ws.Cells.Item($r, 5).Value2 = $row."사용여부"
        $r++
    }

    $ws.Range("A1:E1").Font.Bold = $true
    $ws.Columns("A:E").AutoFit() | Out-Null
    # Keep the embedded master inert inside the add-in. It is read only when
    # Naver local-list features run, and copied to user workbooks only by the
    # explicit "행정동 양식" command.
    $ws.Visible = 2
}

function Add-ProgressStatusUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmProgressStatus"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "작업 진행 현황"
    $form.Properties.Item("Width").Value = 420
    $form.Properties.Item("Height").Value = 180

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "작업을 준비 중입니다."
    $lblTitle.Left = 14
    $lblTitle.Top = 12
    $lblTitle.Width = 380
    $lblTitle.Height = 22

    $lblMessage = $designer.Controls.Add("Forms.Label.1", "lblMessage", $true)
    $lblMessage.Caption = ""
    $lblMessage.Left = 14
    $lblMessage.Top = 42
    $lblMessage.Width = 380
    $lblMessage.Height = 32

    $barBack = $designer.Controls.Add("Forms.Label.1", "barBack", $true)
    $barBack.Caption = ""
    $barBack.Left = 14
    $barBack.Top = 84
    $barBack.Width = 360
    $barBack.Height = 18
    $barBack.BackStyle = 1
    $barBack.BackColor = 12632256

    $barFill = $designer.Controls.Add("Forms.Label.1", "barFill", $true)
    $barFill.Caption = ""
    $barFill.Left = 14
    $barFill.Top = 84
    $barFill.Width = 1
    $barFill.Height = 18
    $barFill.BackStyle = 1
    $barFill.BackColor = 5287936

    $lblPercent = $designer.Controls.Add("Forms.Label.1", "lblPercent", $true)
    $lblPercent.Caption = "0% (0/0)"
    $lblPercent.Left = 14
    $lblPercent.Top = 112
    $lblPercent.Width = 200
    $lblPercent.Height = 18

    $lblWait = $designer.Controls.Add("Forms.Label.1", "lblWait", $true)
    $lblWait.Caption = "작업 중에는 Excel이 잠시 느려질 수 있습니다."
    $lblWait.Left = 180
    $lblWait.Top = 112
    $lblWait.Width = 210
    $lblWait.Height = 18

    $formCode = @'
Option Explicit

Private mTotal As Long
Private Const BAR_MAX_WIDTH As Single = 360

Public Sub Configure(ByVal title As String, ByVal totalCount As Long, ByVal message As String)
    Me.Caption = title
    lblTitle.Caption = title
    mTotal = totalCount
    UpdateProgress 0, totalCount, message
End Sub

Public Sub UpdateProgress(ByVal currentCount As Long, ByVal totalCount As Long, ByVal message As String)
    Dim pct As Double
    If totalCount > 0 Then
        pct = currentCount / totalCount
    Else
        pct = 0
    End If
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1

    lblMessage.Caption = message
    lblPercent.Caption = Format(pct, "0%") & " (" & CStr(currentCount) & "/" & CStr(totalCount) & ")"
    barFill.Width = BAR_MAX_WIDTH * pct
    Repaint
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
'@

    $form.CodeModule.AddFromString($formCode)
}

function Add-RibbonCustomUi {
    param([string]$PackagePath)

    $customUiXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabDataEditingToolkit" label="데이터 툴킷" insertAfterMso="TabView">
        <group id="grpToolkitLaunchers" label="통합 실행">
          <button id="btnLaunchDataCleanup" label="값 정리" size="large" imageMso="NumberFormatGeneral" onAction="Ribbon_LaunchDataCleanup" screentip="값 정리 통합 실행" supertip="숫자 변환, 전화번호 정리, 사업자등록번호 정리, 기업명 정리를 하나의 선택 창에서 실행합니다." />
          <button id="btnLaunchAddressSearch" label="주소/검색" size="large" imageMso="AddressBook" onAction="Ribbon_LaunchAddressSearch" screentip="주소/검색 통합 실행" supertip="주소 유형 판별, 주소 체계 분리, 행정구역 분해, 네이버 주소 검증과 신규 리스트 생성을 하나의 선택 창에서 실행합니다." />
          <button id="btnLaunchRosterQuality" label="명부 품질" size="large" imageMso="ReviewAcceptChange" onAction="Ribbon_LaunchRosterQuality" screentip="명부 품질 통합 실행" supertip="품질 진단, 중복 탐지, 업체명 매칭, 전화번호 검사, 사업자등록 상태조회를 하나의 선택 창에서 실행합니다." />
          <button id="btnLaunchSpssTools" label="SPSS 표 도구" size="large" imageMso="CreateReport" onAction="Ribbon_LaunchSpssTools" screentip="SPSS 표 도구 통합 실행" supertip="목차 생성, 표/배너 재구성, 세로 배너 삭제, 표 메타 추출을 하나의 선택 창에서 실행합니다." />
          <button id="btnLaunchSampleDesign" label="표본설계" size="large" imageMso="PivotTableInsert" onAction="Ribbon_LaunchSampleDesign" screentip="표본설계 통합 실행" supertip="표본설계 시트 생성, 단순 비례/제곱근 비례/멱배분 할당 계산, 진행률 관리표 생성을 하나의 선택 창에서 실행합니다." />
        </group>
        <group id="grpNumberTools" label="값 정리">
          <button id="btnConvertRange" label="숫자 변환" size="large" imageMso="NumberFormatGeneral" onAction="Ribbon_ConvertTextNumbersInColumns" screentip="선택 범위 숫자 변환" supertip="선택한 범위에서 텍스트 형식 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnConvertSheet" label="시트 숫자 변환" imageMso="TableSelect" onAction="Ribbon_ConvertTextNumbersInUsedRange" screentip="시트 전체 숫자 변환" supertip="현재 시트에서 실제 데이터가 있는 마지막 행과 열까지만 잡아 텍스트 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnPhoneRange" label="전화번호 정리" size="large" imageMso="ContactPictureMenu" onAction="Ribbon_NormalizePhonesInColumns" screentip="전화번호 하이픈 정리" supertip="선택 범위 오른쪽에 복사본 열을 만들고 한국 전화번호를 하이픈 포함 표준 형식으로 정리합니다." />
          <button id="btnPhoneDashless" label="전화번호 숫자만" imageMso="ClearFormats" onAction="Ribbon_RemovePhoneDashes" screentip="전화번호 숫자만 남기기" supertip="선택한 전화번호 범위 오른쪽에 복사본 열을 만들고 대시와 구분기호를 제거해 숫자만 남깁니다." />
          <button id="btnCompanyNameClean" label="기업명 정리" size="large" imageMso="TextBoxInsert" onAction="Ribbon_CompanyNameClean" screentip="기업체명/사업체명 정리" supertip="선택 범위 오른쪽에 복사본 열을 만들고 공백, 회사형태 표기, 영어 단어 한글 발음, 알파벳 한글 읽기 등 선택한 규칙으로 기업체명/사업체명을 정리합니다." />
          <button id="btnAddressConvert" label="주소 유형 판별" size="large" imageMso="AddressBook" onAction="Ribbon_AddressConvertOldToRoad" screentip="도로명/동 기준 주소 판별" supertip="선택한 주소 범위 오른쪽에 도로명주소인지 법정/행정동 기준 주소인지 판별 결과를 생성합니다." />
          <button id="btnAddressSplit" label="주소 체계 분리" imageMso="AddressBook" onAction="Ribbon_AddressSplitDongAndRoad" screentip="동 기준/도로명 주소 분리" supertip="주소 범위를 선택하면 원본 오른쪽에 동 기준 주소와 도로명주소 2개 열을 추가합니다. 변환표가 있으면 지번/도로명 양방향 매칭을 사용합니다." />
        </group>
        <group id="grpBusinessNoTools" label="사업자등록">
          <button id="btnBizApiKey" label="인증키 설정" imageMso="ServerConnection" onAction="Ribbon_BizNoSetApiKey" screentip="사업자등록 API 인증키 설정" supertip="공공데이터포털 국세청 사업자등록정보 API 인증키를 저장하거나 변경합니다." />
          <button id="btnBizDashless" label="번호 숫자만" imageMso="ClearFormats" onAction="Ribbon_RemoveBizNoDashes" screentip="사업자등록번호 숫자만 남기기" supertip="선택한 사업자등록번호 범위 오른쪽에 복사본 열을 만들고 대시와 구분기호를 제거해 숫자 10자리 형태로 정리합니다." />
          <button id="btnBizMerge" label="3열 병합" imageMso="CellsMerge" onAction="Ribbon_BizNoMergeThreeColumns" screentip="사업자등록번호 3열 병합" supertip="3개 열로 나뉜 사업자등록번호를 000-00-00000 형식의 새 열로 병합합니다." />
          <button id="btnBizStatus" label="상태조회" size="large" imageMso="DataRefreshAll" onAction="Ribbon_BizNoQueryStatus" screentip="사업자등록 상태조회" supertip="선택한 사업자등록번호 범위를 국세청 API로 조회하고 오른쪽에 결과 열을 추가합니다." />
        </group>
        <group id="grpNaverLocalTools" label="네이버 지역검색">
          <button id="btnNaverApiKey" label="인증키 설정" imageMso="ServerConnection" onAction="Ribbon_NaverLocalSetApiCredentials" screentip="네이버 지역검색 API 인증키 설정" supertip="네이버 Developers의 Client ID와 Client Secret을 저장하거나 변경합니다." />
          <button id="btnNaverBuildList" label="신규 리스트" size="large" imageMso="AddressBook" onAction="Ribbon_NaverLocalBuildListByDong" screentip="행정동 기준 업체 리스트 생성" supertip="행정동 목록 범위와 업종을 입력해 네이버 지역검색 결과를 신규 시트에 생성합니다." />
          <button id="btnNaverVerifyAddress" label="주소 검증" size="large" imageMso="ReviewCompareTwoVersions" onAction="Ribbon_NaverLocalVerifyExistingData" screentip="업체명/주소 검증" supertip="업체명 범위와 주소 범위를 선택하면 네이버 지역검색 결과, 유사도, 검증등급 열을 추가합니다." />
          <button id="btnNaverDongMaster" label="행정동 양식" imageMso="TableInsert" onAction="Ribbon_NaverLocalCreateDongMasterTemplate" screentip="행정동 마스터 양식 생성" supertip="현재 통합문서에 내장 행정동 마스터와 같은 양식의 행정동_마스터 시트를 생성합니다. 이 시트를 수정하면 신규 리스트 생성에 우선 사용됩니다." />
        </group>
        <group id="grpSpssTableTools" label="SPSS 표 정리">
          <button id="btnSpssCreateIndex" label="목차 생성" imageMso="TableOfContentsGallery" onAction="Ribbon_SPSSCreateIndexForActiveSheet" screentip="SPSS 표 목차 생성" supertip="목차가 없는 집계표 시트에서 [로 시작하는 표 제목 행을 찾아 새 목차 시트와 링크를 생성합니다." />
          <button id="btnSpssRebuildFromIndex" label="표/배너 재구성" size="large" imageMso="CreateReport" onAction="Ribbon_SPSSRebuildTablesFromIndex" screentip="SPSS 표/배너 재구성" supertip="표 목록을 드래그하거나 버튼으로 이동하고, 삭제할 표/배너를 선택해 새 집계표와 목차 링크를 생성합니다." />
          <button id="btnSpssDeleteVertical" label="세로 배너 삭제" imageMso="DeleteRows" onAction="Ribbon_SPSSDeleteVerticalBanners" screentip="세로 배너 삭제" supertip="원본 집계표 시트를 복사한 뒤 ▣ 세로 배너와 BOT/TOP/MID 계열 세로형 배너 행을 삭제하고 새 목차를 생성합니다." />
          <button id="btnSpssMetadata" label="표 메타 추출" imageMso="TablePropertiesDialog" onAction="Ribbon_SPSSExtractTableMetadata" screentip="SPSS 표 메타 추출" supertip="집계표의 표 제목, 행 위치, 배너 목록, 이동 링크를 새 시트로 추출합니다." />
        </group>
        <group id="grpRosterQualityTools" label="명부 품질">
          <button id="btnRosterQuality" label="품질 진단" size="large" imageMso="ReviewAcceptChange" onAction="Ribbon_RosterQualityReport" screentip="명부 품질 진단 리포트" supertip="선택한 명부에서 빈값, 전화번호, 주소 유형, 사업자번호 상태를 요약한 리포트 시트를 만듭니다." />
          <button id="btnRosterDuplicate" label="중복 탐지" size="large" imageMso="ReviewCompareTwoVersions" onAction="Ribbon_RosterDetectDuplicates" screentip="중복 업체 탐지" supertip="업체명, 주소, 전화번호, 사업자번호를 조합해 중복 의심 쌍과 점수를 새 시트에 생성합니다." />
          <button id="btnRosterCompanyMatch" label="업체명 매칭" imageMso="TextBoxInsert" onAction="Ribbon_RosterMatchCompanyNames" screentip="업체명 유사그룹 생성" supertip="업체명을 표준화하고 같은 표준화명을 유사그룹으로 묶는 열을 생성합니다." />
          <button id="btnAddressParts" label="행정구역 분해" imageMso="AddressBook" onAction="Ribbon_AddressSplitAdministrativeParts" screentip="주소 행정구역 분해" supertip="주소를 광역시도, 시군구, 읍면동, 상세주소 열로 분해합니다." />
          <button id="btnPhoneValidate" label="전화번호 검사" imageMso="ContactPictureMenu" onAction="Ribbon_PhoneValidateSelection" screentip="전화번호 유효성 검사" supertip="전화번호 범위 오른쪽에 번호 유형과 정상/확인 필요 결과를 생성합니다." />
        </group>
        <group id="grpOutlierTools" label="이상치 탐지">
          <button id="btnOutlierRun" label="분석 열 선택" size="large" imageMso="ConditionalFormattingHighlightCellsRules" onAction="Ribbon_OutlierDetectRun" screentip="이상치 확인 범위 선택" supertip="범위를 직접 지정하면 범위 검증 후 기초통계량표를 먼저 만들고, 오른쪽에 Q1/Q3 기준값 열과 원본 검토복사 열을 추가해 조건부 서식으로 표시합니다." />
        </group>
        <group id="grpSampleDesignTools" label="표본설계">
          <button id="btnSampleDesignListify" label="구성표 리스트화" size="large" imageMso="TableInsert" onAction="Ribbon_SampleDesignListifyPopulationTable" screentip="모집단 구성표 리스트화" supertip="선택 범위 안의 여러 집계표 블록을 자동 탐지하고 미리보기 창에서 구조를 확인한 뒤 리스트형 데이터로 변환합니다." />
          <button id="btnSampleDesignTemplate" label="설계 시트 생성" size="large" imageMso="TableInsert" onAction="Ribbon_SampleDesignCreateTemplate" screentip="표본설계 입력 시트 생성" supertip="모집단 정보, 쿼터, 목표 표본 수, 배분 방식을 입력할 표본설계 시트를 생성합니다." />
          <button id="btnSampleDesignAllocate" label="할당 계산" size="large" imageMso="PivotTableInsert" onAction="Ribbon_SampleDesignRunAllocation" screentip="표본할당 계산" supertip="표본설계_입력 시트를 읽어 단순 비례, 제곱근 비례, 멱배분 등 선택한 방식으로 최종 표본할당표와 검토 시트를 생성합니다." />
          <button id="btnSampleDesignProgress" label="진행률 관리표" imageMso="ChartInsert" onAction="Ribbon_SampleDesignCreateProgressTracker" screentip="표본 진행률 관리표 생성" supertip="할당결과를 기반으로 회수, 부족, 초과, 달성률을 관리하는 표본설계_진행관리 시트를 생성합니다." />
        </group>
        <group id="grpToolkitInfo" label="정보">
          <button id="btnSelectA1AllSheets" label="모든 시트 A1 선택" imageMso="GoToTop" onAction="Ribbon_SelectA1OnAllSheets" screentip="모든 시트 A1 선택" supertip="통합문서의 모든 워크시트를 순회하며 선택 셀을 A1로 변경하고 화면 위치도 최상단으로 맞춥니다." />
          <button id="btnToolkitAbout" label="버전 정보" imageMso="Info" onAction="Ribbon_ToolkitAbout" screentip="툴킷 정보" supertip="현재 데이터 편집 매크로 툴킷 버전과 기능 요약을 표시합니다." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"@

    $archive = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        Set-ZipTextEntry -Archive $archive -EntryName "customUI/customUI.xml" -Text $customUiXml

        $rels = Get-ZipTextEntry -Archive $archive -EntryName "_rels/.rels"
        if ($null -eq $rels) { throw "Package relationship file _rels/.rels not found." }
        $rels = [regex]::Replace($rels, '\s*<Relationship[^>]+Type="http://schemas\.microsoft\.com/office/2006/relationships/ui/extensibility"[^>]*/>', '')
        $nextId = 1
        [regex]::Matches($rels, 'Id="rId(\d+)"') | ForEach-Object {
            $n = [int]$_.Groups[1].Value
            if ($n -ge $nextId) { $nextId = $n + 1 }
        }
        $customRel = "  <Relationship Id=`"rId$nextId`" Type=`"http://schemas.microsoft.com/office/2006/relationships/ui/extensibility`" Target=`"customUI/customUI.xml`"/>"
        $rels = $rels -replace '</Relationships>', "$customRel`r`n</Relationships>"
        Set-ZipTextEntry -Archive $archive -EntryName "_rels/.rels" -Text $rels

        $contentTypes = Get-ZipTextEntry -Archive $archive -EntryName "[Content_Types].xml"
        if ($null -eq $contentTypes) { throw "Package [Content_Types].xml not found." }
        if ($contentTypes -notmatch '/customUI/customUI\.xml') {
            $override = '  <Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>'
            $contentTypes = $contentTypes -replace '</Types>', "$override`r`n</Types>"
            Set-ZipTextEntry -Archive $archive -EntryName "[Content_Types].xml" -Text $contentTypes
        }
    }
    finally {
        $archive.Dispose()
    }
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ProjectRoot) { $ProjectRoot = Split-Path -Parent $ScriptDir }

$ModulePath = Join-Path $ProjectRoot "src\MacroToolkit.bas"
$AdminDongMasterCsv = Join-Path $ProjectRoot "resources\admin_dong_master.csv"
$DevWorkbook = Join-Path $ProjectRoot "dev\DataEditingMacroToolkit_dev.xlsm"
$ReleaseRoot = Join-Path $ProjectRoot "release"

$commit = "unknown"
$wasDirtyBeforePublish = $false
try {
    $commit = (git -C $ProjectRoot rev-parse --short HEAD).Trim()
    $repoRoot = (git -C $ProjectRoot rev-parse --show-toplevel).Trim()
    $projectRelPath = (Resolve-Path -LiteralPath $ProjectRoot).Path.Substring((Resolve-Path -LiteralPath $repoRoot).Path.Length).TrimStart('\', '/').Replace('\', '/')
    if ($projectRelPath) {
        $prePublishStatus = (git -C $repoRoot status --short -- $projectRelPath)
    } else {
        $prePublishStatus = (git -C $repoRoot status --short)
    }
    if ($prePublishStatus) { $wasDirtyBeforePublish = $true }
} catch {}

if (-not (Test-Path -LiteralPath $ModulePath)) { throw "VBA module not found: $ModulePath" }
if (-not (Test-Path -LiteralPath $DevWorkbook)) { throw "Development workbook not found: $DevWorkbook" }

if (-not $Version) {
    $moduleText = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
    if ($moduleText -match 'TOOLKIT_VERSION\s+As\s+String\s*=\s*"([^"]+)"') { $Version = $Matches[1] }
    else { throw "Could not read TOOLKIT_VERSION from $ModulePath" }
}

$VersionTag = "v$Version"
$VersionDir = Join-Path $ReleaseRoot $VersionTag
$LatestDir = Join-Path $ReleaseRoot "latest"
$ReleaseWorkbookName = "DataEditingMacroToolkit_$VersionTag.xlsm"
$ReleaseAddinName = "DataEditingMacroToolkit_$VersionTag.xlam"
$LatestWorkbookName = "DataEditingMacroToolkit_latest.xlsm"
$LatestAddinName = "DataEditingMacroToolkit_latest.xlam"

New-Item -ItemType Directory -Force -Path $VersionDir,$LatestDir | Out-Null

$ReleaseWorkbook = Join-Path $VersionDir $ReleaseWorkbookName
$ReleaseAddin = Join-Path $VersionDir $ReleaseAddinName
$LatestWorkbook = Join-Path $LatestDir $LatestWorkbookName
$LatestAddin = Join-Path $LatestDir $LatestAddinName
$ImportModulePath = Join-Path $VersionDir "_tmp_MacroToolkit_cp949.bas"

Copy-Item -Path $DevWorkbook -Destination $ReleaseWorkbook -Force
Copy-Item -Path $DevWorkbook -Destination $LatestWorkbook -Force
Copy-Item -Path $ModulePath -Destination (Join-Path $VersionDir "MacroToolkit.bas") -Force
Copy-Item -Path $ModulePath -Destination (Join-Path $LatestDir "MacroToolkit.bas") -Force

$moduleSource = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
$cp949 = [System.Text.Encoding]::GetEncoding(949)
[System.IO.File]::WriteAllText($ImportModulePath, $moduleSource, $cp949)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($DevWorkbook)
    try {
        foreach ($component in @($workbook.VBProject.VBComponents)) {
            if (($component.Type -eq 1 -and ($component.Name -eq "MacroToolkit" -or $component.Name -like "Module*")) -or
                ($component.Type -eq 3 -and ($component.Name -eq "frmSpssRebuildOptions" -or $component.Name -eq "frmRangeSelectOptions" -or $component.Name -eq "frmCompanyNameOptions" -or $component.Name -eq "frmBizStatusFieldsOptions" -or $component.Name -eq "frmNaverLocalBuildOptions" -or $component.Name -eq "frmNaverLocalVerifyOptions" -or $component.Name -eq "frmProgressStatus" -or $component.Name -eq "frmToolkitButtonMenu" -or $component.Name -eq "frmToolkitCheckListMenu" -or $component.Name -eq "frmPopulationListifyPreview"))) {
                $workbook.VBProject.VBComponents.Remove($component)
            }
        }
        $imported = $workbook.VBProject.VBComponents.Import($ImportModulePath)
        $imported.Name = "MacroToolkit"
        Add-ToolkitButtonMenuUserForm -Workbook $workbook
        Add-ToolkitCheckListMenuUserForm -Workbook $workbook
        Add-PopulationListifyPreviewUserForm -Workbook $workbook
        Add-SpssRebuildUserForm -Workbook $workbook
        Add-CompanyNameOptionsUserForm -Workbook $workbook
        Add-BizStatusFieldsOptionsUserForm -Workbook $workbook
        Add-NaverLocalBuildOptionsUserForm -Workbook $workbook
        Add-NaverLocalVerifyOptionsUserForm -Workbook $workbook
        Add-ProgressStatusUserForm -Workbook $workbook
        Add-AdminDongMasterWorksheet -Workbook $workbook -CsvPath $AdminDongMasterCsv
    }
    catch {
        throw "Could not import VBA module into .xlam. Enable Excel Trust Center > Macro Settings > Trust access to the VBA project object model. Details: $($_.Exception.Message)"
    }

    $workbook.IsAddin = $true
    $tempAddin = Join-Path $VersionDir ("_tmp_" + $ReleaseAddinName)
    if (Test-Path -LiteralPath $tempAddin) { Remove-Item -LiteralPath $tempAddin -Force }

    $workbook.SaveAs($tempAddin, 55)
    $workbook.Close($false)
    $workbook = $null

    Add-RibbonCustomUi -PackagePath $tempAddin

    Copy-Item -Path $tempAddin -Destination $ReleaseAddin -Force
    Copy-Item -Path $tempAddin -Destination $LatestAddin -Force
    Remove-Item -LiteralPath $tempAddin -Force
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) | Out-Null } catch {} }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if (Test-Path -LiteralPath $ImportModulePath) { Remove-Item -LiteralPath $ImportModulePath -Force }
}

if ($wasDirtyBeforePublish) { $commit = "$commit-dirty" }

$manifest = [ordered]@{
    name = "Data Editing Macro Toolkit"
    version = $Version
    versionTag = $VersionTag
    publishedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss K")
    gitCommit = $commit
    addin = $ReleaseAddinName
    workbook = $ReleaseWorkbookName
    sourceModule = "MacroToolkit.bas"
    installFile = $ReleaseAddinName
    containsVbaModule = $true
    containsRibbonTab = $true
    ribbonTab = "데이터 툴킷"
    ribbonInsertAfter = "보기"
    vbaImportEncoding = "CP949"
    notes = "Install the .xlam file as an Excel add-in. It adds a 데이터 툴킷 ribbon tab next to View."
}

$manifest | ConvertTo-Json | Set-Content -Encoding UTF8 -Path (Join-Path $VersionDir "manifest.json")
$manifest | ConvertTo-Json | Set-Content -Encoding UTF8 -Path (Join-Path $LatestDir "manifest.json")

Write-Output "Published add-in: $ReleaseAddin"
Write-Output "Published workbook: $ReleaseWorkbook"
Write-Output "Updated latest add-in: $LatestAddin"

