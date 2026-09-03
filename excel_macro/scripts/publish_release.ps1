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

function Test-ApiSecretStorageName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $normalized = $Name.ToLowerInvariant() -replace '[\s_\-\.]', ''
    if ($normalized -match '(apikey|servicekey|clientsecret|clientid|credential|secret|accesstoken|refreshtoken|authtoken|authkey)') {
        return $true
    }
    return ($Name -match '(인증키|서비스키|비밀키|토큰)')
}

function Test-ApiSecretWorksheetName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $normalized = $Name.ToLowerInvariant() -replace '[\s_\-\.]', ''
    return ($normalized -match '^(macrotk|macrotoolkit|toolkit|api)?(credentials|secrets|apikeys|tokens)$' -or
        $Name -match '^_(MacroToolkit_)?(Credentials|Secrets|ApiKeys|Tokens)$')
}

function Clear-ComNamedValues {
    param($Collection)

    if ($null -eq $Collection) { return }
    try {
        for ($i = $Collection.Count; $i -ge 1; $i--) {
            $item = $Collection.Item($i)
            $itemName = ""
            try { $itemName = [string]$item.Name } catch { $itemName = "" }
            if (Test-ApiSecretStorageName $itemName) {
                try { $item.Delete() | Out-Null } catch {
                    try { $item.Value = "" } catch {}
                }
            }
        }
    } catch {}
}

function Clear-WorkbookStoredApiSecrets {
    param($Workbook)

    if ($null -eq $Workbook) { return }

    Clear-ComNamedValues -Collection $Workbook.Names
    try { Clear-ComNamedValues -Collection $Workbook.CustomDocumentProperties } catch {}
    try { Clear-ComNamedValues -Collection $Workbook.Variables } catch {}

    try {
        foreach ($worksheet in @($Workbook.Worksheets)) {
            if (Test-ApiSecretWorksheetName $worksheet.Name) {
                $worksheet.Delete() | Out-Null
            }
        }
    } catch {}
}

function Clear-WorkbookFileStoredApiSecrets {
    param(
        $Excel,
        [string]$Path
    )

    if ($null -eq $Excel -or -not (Test-Path -LiteralPath $Path)) { return }
    $sanitizingWorkbook = $null
    try {
        $sanitizingWorkbook = $Excel.Workbooks.Open($Path)
        Clear-WorkbookStoredApiSecrets -Workbook $sanitizingWorkbook
        $sanitizingWorkbook.Save()
    }
    finally {
        if ($null -ne $sanitizingWorkbook) {
            try { $sanitizingWorkbook.Close($false) | Out-Null } catch {}
        }
    }
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

function Add-SampleDesignSetupUserForm {
    param($Workbook)

    $script:publishStage = "sample setup form: create component"
    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmSampleDesignSetup"
    $designer = $form.Designer
    $script:publishStage = "sample setup form: set form properties"
    $form.Properties.Item("Caption").Value = "표본설계 설정 마법사"
    $form.Properties.Item("Width").Value = 650
    $form.Properties.Item("Height").Value = 570
    $form.Properties.Item("ScrollBars").Value = 2

    $script:publishStage = "sample setup form: title"
    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "표본설계 시트를 만들기 전에 기본 목표와 사용할 쿼터 변수를 먼저 지정합니다."
    $lblTitle.Left = 12
    $lblTitle.Top = 10
    $lblTitle.Width = 600
    $lblTitle.Height = 28
    $lblTitle.WordWrap = $true

    $script:publishStage = "sample setup form: input mode options"
    $optList = $designer.Controls.Add("Forms.OptionButton.1", "optList", $true)
    $optList.Caption = "리스트형 입력"
    $optList.Left = 18
    $optList.Top = 48
    $optList.Width = 120
    $optList.Height = 20
    $optList.Value = $true

    $optTable = $designer.Controls.Add("Forms.OptionButton.1", "optTable", $true)
    $optTable.Caption = "집계표형 입력"
    $optTable.Left = 148
    $optTable.Top = 48
    $optTable.Width = 120
    $optTable.Height = 20

    $script:publishStage = "sample setup form: basic fields"
    $basic = @(
        @("lblTargetN", "전체 목표 표본수", "txtTargetN", "800", 18, 86),
        @("lblAlpha", "멱배분 alpha", "txtAlpha", "0.5", 18, 146),
        @("lblMinN", "기본 최소표본", "txtMinN", "0", 330, 116),
        @("lblMaxN", "기본 최대표본", "txtMaxN", "0", 330, 146)
    )
    foreach ($item in $basic) {
        $lbl = $designer.Controls.Add("Forms.Label.1", $item[0], $true)
        $lbl.Caption = $item[1]
        $lbl.Left = $item[4]
        $lbl.Top = $item[5]
        $lbl.Width = 110
        $lbl.Height = 18

        $txt = $designer.Controls.Add("Forms.TextBox.1", $item[2], $true)
        $txt.Text = $item[3]
        $txt.Left = $item[4] + 116
        $txt.Top = $item[5] - 2
        $txt.Width = 120
        $txt.Height = 20
    }

    $script:publishStage = "sample setup form: method combo"
    $lblMethod = $designer.Controls.Add("Forms.Label.1", "lblMethod", $true)
    $lblMethod.Caption = "배분 방식"
    $lblMethod.Left = 18
    $lblMethod.Top = 116
    $lblMethod.Width = 110
    $lblMethod.Height = 18

    $cboMethod = $designer.Controls.Add("Forms.ComboBox.1", "cboMethod", $true)
    $cboMethod.Left = 134
    $cboMethod.Top = 114
    $cboMethod.Width = 150
    $cboMethod.Height = 20

    $script:publishStage = "sample setup form: quota labels"
    $lblQuota = $designer.Controls.Add("Forms.Label.1", "lblQuota", $true)
    $lblQuota.Caption = "사용할 쿼터 변수"
    $lblQuota.Left = 18
    $lblQuota.Top = 192
    $lblQuota.Width = 180
    $lblQuota.Height = 18

    $lblQuotaHelp = $designer.Controls.Add("Forms.Label.1", "lblQuotaHelp", $true)
    $lblQuotaHelp.Caption = "체크한 변수만 계산에 사용됩니다. 변수명은 프로젝트 용어에 맞게 직접 수정할 수 있습니다."
    $lblQuotaHelp.Left = 160
    $lblQuotaHelp.Top = 192
    $lblQuotaHelp.Width = 460
    $lblQuotaHelp.Height = 18

    $script:publishStage = "sample setup form: quota controls"
    $quotaNames = @("산업", "대분류", "지역", "조사지역", "대조지역", "존", "성별", "연령")
    for ($i = 1; $i -le 8; $i++) {
        $row = [int][Math]::Floor(($i - 1) / 2)
        $col = [int](($i - 1) % 2)
        $left = [int](18 + ($col * 306))
        $top = [int](222 + ($row * 34))

        $chk = $designer.Controls.Add("Forms.CheckBox.1", "chkQuota$i", $true)
        $chk.Caption = "사용"
        $chk.Left = $left
        $chk.Top = $top
        $chk.Width = 54
        $chk.Height = 18
        $chk.Value = ($i -le 3)

        $txt = $designer.Controls.Add("Forms.TextBox.1", "txtQuota$i", $true)
        $txt.Text = $quotaNames[$i - 1]
        $txt.Left = $left + 58
        $txt.Top = $top - 2
        $txt.Width = 190
        $txt.Height = 20
    }

    $script:publishStage = "sample setup form: guide and buttons"
    $lblGuide = $designer.Controls.Add("Forms.Label.1", "lblGuide", $true)
    $lblGuide.Caption = "그룹별 목표 표본수가 필요한 경우 시트 생성 후 '목표표본 그룹표'에 조건과 목표표본수를 입력하세요. 예: 산업=SW AND 지역=수도권 목표 300"
    $lblGuide.Left = 18
    $lblGuide.Top = 370
    $lblGuide.Width = 590
    $lblGuide.Height = 42
    $lblGuide.WordWrap = $true

    $cmdOK = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOK.Caption = "시트 생성"
    $cmdOK.Left = 422
    $cmdOK.Top = 486
    $cmdOK.Width = 90
    $cmdOK.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 522
    $cmdCancel.Top = 486
    $cmdCancel.Width = 78
    $cmdCancel.Height = 28

    $script:publishStage = "sample setup form: code module"
    $formCode = @'
Option Explicit
Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = True
    cboMethod.Clear
    cboMethod.AddItem "단순 비례"
    cboMethod.AddItem "제곱근 비례"
    cboMethod.AddItem "멱배분"
    cboMethod.AddItem "균등 배분"
    cboMethod.AddItem "사용자 지정 가중"
    cboMethod.ListIndex = 0
End Sub

Private Sub cboMethod_Change()
    txtAlpha.Enabled = (cboMethod.Value = "멱배분")
End Sub

Private Sub cmdOK_Click()
    If CLng(Val(txtTargetN.Text)) <= 0 Then
        MsgBox "전체 목표 표본수는 1 이상의 숫자로 입력하세요.", vbExclamation, "표본설계"
        txtTargetN.SetFocus
        Exit Sub
    End If
    If cboMethod.Value = "멱배분" And CDbl(Val(txtAlpha.Text)) <= 0 Then
        MsgBox "멱배분 alpha는 0보다 큰 숫자로 입력하세요.", vbExclamation, "표본설계"
        txtAlpha.SetFocus
        Exit Sub
    End If

    Dim hasQuota As Boolean
    Dim i As Long
    For i = 1 To 8
        If Controls("chkQuota" & CStr(i)).Value Then
            If Len(Trim$(CStr(Controls("txtQuota" & CStr(i)).Text))) = 0 Then
                MsgBox "사용할 쿼터 변수명은 비워둘 수 없습니다.", vbExclamation, "표본설계"
                Controls("txtQuota" & CStr(i)).SetFocus
                Exit Sub
            End If
            hasQuota = True
        End If
    Next i
    If Not hasQuota Then
        MsgBox "최소 1개 이상의 쿼터 변수를 선택하세요.", vbExclamation, "표본설계"
        Exit Sub
    End If

    Cancelled = False
    Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Hide
End Sub

Public Property Get InputMode() As String
    If optTable.Value Then
        InputMode = "table"
    Else
        InputMode = "list"
    End If
End Property

Public Property Get TargetN() As String
    TargetN = CStr(txtTargetN.Text)
End Property

Public Property Get MethodName() As String
    MethodName = CStr(cboMethod.Value)
End Property

Public Property Get AlphaValue() As String
    AlphaValue = CStr(txtAlpha.Text)
End Property

Public Property Get MinN() As String
    MinN = CStr(txtMinN.Text)
End Property

Public Property Get MaxN() As String
    MaxN = CStr(txtMaxN.Text)
End Property

Public Property Get QuotaHeaders() As String
    Dim i As Long
    For i = 1 To 8
        If i > 1 Then QuotaHeaders = QuotaHeaders & "|"
        QuotaHeaders = QuotaHeaders & Trim$(CStr(Controls("txtQuota" & CStr(i)).Text))
    Next i
End Property

Public Property Get QuotaEnabled() As String
    Dim i As Long
    For i = 1 To 8
        If i > 1 Then QuotaEnabled = QuotaEnabled & "|"
        If Controls("chkQuota" & CStr(i)).Value Then
            QuotaEnabled = QuotaEnabled & "1"
        Else
            QuotaEnabled = QuotaEnabled & "0"
        End If
    Next i
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancelled = True
    End If
End Sub
'@
    $form.CodeModule.AddFromString($formCode)
}

function Add-SampleDesignBandBuilderUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmSampleDesignBandBuilder"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "표본설계 구간 만들기"
    $form.Properties.Item("Width").Value = 720
    $form.Properties.Item("Height").Value = 640
    $form.Properties.Item("ScrollBars").Value = 2

    $lblGuide = $designer.Controls.Add("Forms.Label.1", "lblGuide", $true)
    $lblGuide.Caption = "선택한 숫자 열 오른쪽에 구간 열을 생성합니다. 시작/종료 조건을 확인한 뒤 실행하세요."
    $lblGuide.Left = 12
    $lblGuide.Top = 10
    $lblGuide.Width = 660
    $lblGuide.Height = 30
    $lblGuide.WordWrap = $true

    $lblHeader = $designer.Controls.Add("Forms.Label.1", "lblHeader", $true)
    $lblHeader.Caption = "결과 열 헤더"
    $lblHeader.Left = 18
    $lblHeader.Top = 52
    $lblHeader.Width = 90
    $lblHeader.Height = 18

    $txtOutputHeader = $designer.Controls.Add("Forms.TextBox.1", "txtOutputHeader", $true)
    $txtOutputHeader.Left = 110
    $txtOutputHeader.Top = 48
    $txtOutputHeader.Width = 180
    $txtOutputHeader.Height = 22

    $lblCount = $designer.Controls.Add("Forms.Label.1", "lblCount", $true)
    $lblCount.Caption = "구간 개수"
    $lblCount.Left = 330
    $lblCount.Top = 52
    $lblCount.Width = 70
    $lblCount.Height = 18

    $txtBandCount = $designer.Controls.Add("Forms.TextBox.1", "txtBandCount", $true)
    $txtBandCount.Left = 404
    $txtBandCount.Top = 48
    $txtBandCount.Width = 50
    $txtBandCount.Height = 22

    $lblHint = $designer.Controls.Add("Forms.Label.1", "lblHint", $true)
    $lblHint.Caption = "마지막 구간의 시작 조건은 기본적으로 1구간의 시작 조건과 동일하게 설정됩니다."
    $lblHint.Left = 18
    $lblHint.Top = 82
    $lblHint.Width = 650
    $lblHint.Height = 18

    $headers = @(
        @("lblSeq", "구간", 18, 108, 36),
        @("lblStart", "시작값", 62, 108, 90),
        @("lblStartOp", "시작 조건", 158, 108, 74),
        @("lblEnd", "종료값", 238, 108, 90),
        @("lblEndOp", "종료 조건", 334, 108, 74),
        @("lblBandLabel", "구간명(선택)", 418, 108, 180)
    )
    foreach ($item in $headers) {
        $lbl = $designer.Controls.Add("Forms.Label.1", $item[0], $true)
        $lbl.Caption = $item[1]
        $lbl.Left = $item[2]
        $lbl.Top = $item[3]
        $lbl.Width = $item[4]
        $lbl.Height = 16
    }

    for ($i = 1; $i -le 15; $i++) {
        $top = 130 + (($i - 1) * 28)

        $lbl = $designer.Controls.Add("Forms.Label.1", "lblBand$i", $true)
        $lbl.Caption = "$i"
        $lbl.Left = 26
        $lbl.Top = $top + 4
        $lbl.Width = 24
        $lbl.Height = 16

        $txtStart = $designer.Controls.Add("Forms.TextBox.1", "txtStart$i", $true)
        $txtStart.Left = 62
        $txtStart.Top = $top
        $txtStart.Width = 86
        $txtStart.Height = 22

        $cboStart = $designer.Controls.Add("Forms.ComboBox.1", "cboStartOp$i", $true)
        $cboStart.Left = 158
        $cboStart.Top = $top
        $cboStart.Width = 70
        $cboStart.Height = 22

        $txtEnd = $designer.Controls.Add("Forms.TextBox.1", "txtEnd$i", $true)
        $txtEnd.Left = 238
        $txtEnd.Top = $top
        $txtEnd.Width = 86
        $txtEnd.Height = 22

        $cboEnd = $designer.Controls.Add("Forms.ComboBox.1", "cboEndOp$i", $true)
        $cboEnd.Left = 334
        $cboEnd.Top = $top
        $cboEnd.Width = 70
        $cboEnd.Height = 22

        $txtLabel = $designer.Controls.Add("Forms.TextBox.1", "txtLabel$i", $true)
        $txtLabel.Left = 418
        $txtLabel.Top = $top
        $txtLabel.Width = 210
        $txtLabel.Height = 22
    }

    $cmdOK = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOK.Caption = "구간 열 생성"
    $cmdOK.Left = 486
    $cmdOK.Top = 560
    $cmdOK.Width = 100
    $cmdOK.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 596
    $cmdCancel.Top = 560
    $cmdCancel.Width = 70
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit
Private Const MAX_BANDS As Long = 15
Public Cancelled As Boolean
Public BandSpecText As String

Public Sub Configure(ByVal defaultSpecText As String, ByVal defaultHeader As String)
    Cancelled = True
    BandSpecText = ""
    txtOutputHeader.Text = defaultHeader
    InitCombos
    ClearRows

    If Len(Trim$(defaultSpecText)) > 0 Then
        LoadSpec defaultSpecText
    Else
        txtBandCount.Text = "5"
        SetupDefaultRows 5
    End If

    ApplyBandCount
End Sub

Private Sub InitCombos()
    Dim i As Long
    For i = 1 To MAX_BANDS
        With Controls("cboStartOp" & CStr(i))
            .Clear
            .AddItem "이상"
            .AddItem "초과"
            .Value = "이상"
        End With
        With Controls("cboEndOp" & CStr(i))
            .Clear
            .AddItem "미만"
            .AddItem "이하"
            .Value = "미만"
        End With
    Next i
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To MAX_BANDS
        Controls("txtStart" & CStr(i)).Text = ""
        Controls("txtEnd" & CStr(i)).Text = ""
        Controls("txtLabel" & CStr(i)).Text = ""
        Controls("cboStartOp" & CStr(i)).Value = "이상"
        Controls("cboEndOp" & CStr(i)).Value = "미만"
    Next i
End Sub

Private Sub SetupDefaultRows(ByVal bandCount As Long)
    Dim i As Long
    For i = 1 To bandCount
        Controls("cboStartOp" & CStr(i)).Value = "이상"
        Controls("cboEndOp" & CStr(i)).Value = "미만"
    Next i
End Sub

Private Sub LoadSpec(ByVal specText As String)
    Dim lines As Variant
    lines = Split(Replace(CStr(specText), vbCrLf, vbLf), vbLf)

    Dim visibleCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) > 0 And visibleCount < MAX_BANDS Then
            Dim parts As Variant
            parts = Split(CStr(lines(i)), vbTab)
            If UBound(parts) >= 4 Then
                visibleCount = visibleCount + 1
                Controls("txtStart" & CStr(visibleCount)).Text = CStr(parts(1))
                Controls("cboStartOp" & CStr(visibleCount)).Value = NormalizeStartOp(CStr(parts(2)))
                Controls("txtEnd" & CStr(visibleCount)).Text = CStr(parts(3))
                Controls("cboEndOp" & CStr(visibleCount)).Value = NormalizeEndOp(CStr(parts(4)))
                If UBound(parts) >= 5 Then Controls("txtLabel" & CStr(visibleCount)).Text = CStr(parts(5))
            End If
        End If
    Next i
    If visibleCount = 0 Then visibleCount = 5
    txtBandCount.Text = CStr(visibleCount)
End Sub

Private Function NormalizeStartOp(ByVal opText As String) As String
    If InStr(1, opText, "초과", vbTextCompare) > 0 Then
        NormalizeStartOp = "초과"
    Else
        NormalizeStartOp = "이상"
    End If
End Function

Private Function NormalizeEndOp(ByVal opText As String) As String
    If InStr(1, opText, "이하", vbTextCompare) > 0 Then
        NormalizeEndOp = "이하"
    Else
        NormalizeEndOp = "미만"
    End If
End Function

Private Function BandCount() As Long
    BandCount = CLng(Val(txtBandCount.Text))
    If BandCount < 1 Then BandCount = 1
    If BandCount > MAX_BANDS Then BandCount = MAX_BANDS
End Function

Private Sub txtBandCount_Change()
    ApplyBandCount
End Sub

Private Sub cboStartOp1_Change()
    Dim cnt As Long
    cnt = BandCount()
    If cnt > 1 Then Controls("cboStartOp" & CStr(cnt)).Value = cboStartOp1.Value
End Sub

Private Sub ApplyBandCount()
    Dim cnt As Long
    cnt = BandCount()

    Dim i As Long
    For i = 1 To MAX_BANDS
        Dim isVisible As Boolean
        isVisible = (i <= cnt)
        Controls("lblBand" & CStr(i)).Visible = isVisible
        Controls("txtStart" & CStr(i)).Visible = isVisible
        Controls("cboStartOp" & CStr(i)).Visible = isVisible
        Controls("txtEnd" & CStr(i)).Visible = isVisible
        Controls("cboEndOp" & CStr(i)).Visible = isVisible
        Controls("txtLabel" & CStr(i)).Visible = isVisible
    Next i

    If cnt > 1 Then Controls("cboStartOp" & CStr(cnt)).Value = cboStartOp1.Value
    cmdOK.Top = 142 + cnt * 28
    cmdCancel.Top = cmdOK.Top
    ScrollHeight = cmdOK.Top + 60
    Height = Application.WorksheetFunction.Min(640, cmdOK.Top + 78)
End Sub

Private Sub cmdOK_Click()
    Dim cnt As Long
    cnt = BandCount()
    txtBandCount.Text = CStr(cnt)

    If Len(Trim$(txtOutputHeader.Text)) = 0 Then
        MsgBox "결과 열 헤더를 입력하세요.", vbExclamation, "표본설계 구간 만들기"
        txtOutputHeader.SetFocus
        Exit Sub
    End If

    Dim i As Long
    BandSpecText = ""
    For i = 1 To cnt
        Dim startText As String, endText As String, labelText As String
        startText = Trim$(Controls("txtStart" & CStr(i)).Text)
        endText = Trim$(Controls("txtEnd" & CStr(i)).Text)
        labelText = Trim$(Controls("txtLabel" & CStr(i)).Text)

        If Len(startText) = 0 And Len(endText) = 0 Then
            MsgBox CStr(i) & "구간은 시작값 또는 종료값 중 하나 이상이 필요합니다.", vbExclamation, "표본설계 구간 만들기"
            Controls("txtStart" & CStr(i)).SetFocus
            Exit Sub
        End If
        If Len(BandSpecText) > 0 Then BandSpecText = BandSpecText & vbLf
        BandSpecText = BandSpecText & CStr(i) & vbTab & startText & vbTab & _
                       Controls("cboStartOp" & CStr(i)).Value & vbTab & _
                       endText & vbTab & Controls("cboEndOp" & CStr(i)).Value & vbTab & labelText
    Next i

    Cancelled = False
    Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    BandSpecText = ""
    Hide
End Sub

Public Property Get OutputHeader() As String
    OutputHeader = CStr(txtOutputHeader.Text)
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancelled = True
        BandSpecText = ""
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

function Add-DatabaseMergeSetupUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmDatabaseMergeSetup"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "시트 데이터 통합 설정"
    $form.Properties.Item("Width").Value = 760
    $form.Properties.Item("Height").Value = 660

    $lblGuide = $designer.Controls.Add("Forms.Label.1", "lblGuide", $true)
    $lblGuide.Caption = "통합할 시트, 헤더 행/구조, KEY 열 지정 방식, ID 처리 방식, 가져올 헤더 또는 표시행 기준, 결과 시트 위치를 지정하세요. 다음 단계에서 시트별 헤더 매칭을 확인합니다."
    $lblGuide.Left = 12
    $lblGuide.Top = 10
    $lblGuide.Width = 720
    $lblGuide.Height = 30
    $lblGuide.WordWrap = $true

    $lblSheets = $designer.Controls.Add("Forms.Label.1", "lblSheets", $true)
    $lblSheets.Caption = "통합할 시트"
    $lblSheets.Left = 12
    $lblSheets.Top = 52
    $lblSheets.Width = 180
    $lblSheets.Height = 18

    $lstSheets = $designer.Controls.Add("Forms.ListBox.1", "lstSheets", $true)
    $lstSheets.Left = 12
    $lstSheets.Top = 74
    $lstSheets.Width = 250
    $lstSheets.Height = 410
    $lstSheets.MultiSelect = 2

    $cmdAll = $designer.Controls.Add("Forms.CommandButton.1", "cmdAll", $true)
    $cmdAll.Caption = "전체 선택"
    $cmdAll.Left = 12
    $cmdAll.Top = 500
    $cmdAll.Width = 92
    $cmdAll.Height = 26

    $cmdNone = $designer.Controls.Add("Forms.CommandButton.1", "cmdNone", $true)
    $cmdNone.Caption = "전체 해제"
    $cmdNone.Left = 114
    $cmdNone.Top = 500
    $cmdNone.Width = 92
    $cmdNone.Height = 26

    $lblKey = $designer.Controls.Add("Forms.Label.1", "lblKey", $true)
    $lblKey.Caption = "결과 KEY 헤더"
    $lblKey.Left = 286
    $lblKey.Top = 54
    $lblKey.Width = 120
    $lblKey.Height = 18

    $txtKeyHeader = $designer.Controls.Add("Forms.TextBox.1", "txtKeyHeader", $true)
    $txtKeyHeader.Left = 410
    $txtKeyHeader.Top = 50
    $txtKeyHeader.Width = 176
    $txtKeyHeader.Height = 22

    $chkManualKeyColumn = $designer.Controls.Add("Forms.CheckBox.1", "chkManualKeyColumn", $true)
    $chkManualKeyColumn.Caption = "KEY 열 직접 지정"
    $chkManualKeyColumn.Left = 594
    $chkManualKeyColumn.Top = 50
    $chkManualKeyColumn.Width = 138
    $chkManualKeyColumn.Height = 22

    $lblOutputName = $designer.Controls.Add("Forms.Label.1", "lblOutputName", $true)
    $lblOutputName.Caption = "결과 시트 이름"
    $lblOutputName.Left = 286
    $lblOutputName.Top = 88
    $lblOutputName.Width = 120
    $lblOutputName.Height = 18

    $txtOutputSheetName = $designer.Controls.Add("Forms.TextBox.1", "txtOutputSheetName", $true)
    $txtOutputSheetName.Left = 410
    $txtOutputSheetName.Top = 84
    $txtOutputSheetName.Width = 180
    $txtOutputSheetName.Height = 22

    $lblPosition = $designer.Controls.Add("Forms.Label.1", "lblPosition", $true)
    $lblPosition.Caption = "생성 위치"
    $lblPosition.Left = 600
    $lblPosition.Top = 88
    $lblPosition.Width = 64
    $lblPosition.Height = 18

    $cmbPosition = $designer.Controls.Add("Forms.ComboBox.1", "cmbPosition", $true)
    $cmbPosition.Left = 660
    $cmbPosition.Top = 84
    $cmbPosition.Width = 72
    $cmbPosition.Height = 22
    $cmbPosition.Style = 2

    $lblHeaderRow = $designer.Controls.Add("Forms.Label.1", "lblHeaderRow", $true)
    $lblHeaderRow.Caption = "헤더 행"
    $lblHeaderRow.Left = 286
    $lblHeaderRow.Top = 122
    $lblHeaderRow.Width = 120
    $lblHeaderRow.Height = 18

    $txtHeaderRow = $designer.Controls.Add("Forms.TextBox.1", "txtHeaderRow", $true)
    $txtHeaderRow.Left = 410
    $txtHeaderRow.Top = 118
    $txtHeaderRow.Width = 72
    $txtHeaderRow.Height = 22
    $txtHeaderRow.Text = "자동"

    $lblHeaderMode = $designer.Controls.Add("Forms.Label.1", "lblHeaderMode", $true)
    $lblHeaderMode.Caption = "헤더 구조"
    $lblHeaderMode.Left = 494
    $lblHeaderMode.Top = 122
    $lblHeaderMode.Width = 76
    $lblHeaderMode.Height = 18

    $cmbHeaderMode = $designer.Controls.Add("Forms.ComboBox.1", "cmbHeaderMode", $true)
    $cmbHeaderMode.Left = 576
    $cmbHeaderMode.Top = 118
    $cmbHeaderMode.Width = 156
    $cmbHeaderMode.Height = 22
    $cmbHeaderMode.Style = 2

    $lblIdMode = $designer.Controls.Add("Forms.Label.1", "lblIdMode", $true)
    $lblIdMode.Caption = "ID 처리 방식"
    $lblIdMode.Left = 286
    $lblIdMode.Top = 154
    $lblIdMode.Width = 120
    $lblIdMode.Height = 18

    $cmbIdMode = $designer.Controls.Add("Forms.ComboBox.1", "cmbIdMode", $true)
    $cmbIdMode.Left = 410
    $cmbIdMode.Top = 150
    $cmbIdMode.Width = 322
    $cmbIdMode.Height = 22
    $cmbIdMode.Style = 2

    $chkUseMarkerRow = $designer.Controls.Add("Forms.CheckBox.1", "chkUseMarkerRow", $true)
    $chkUseMarkerRow.Caption = "표시행 기준으로 통합할 열 선택"
    $chkUseMarkerRow.Left = 286
    $chkUseMarkerRow.Top = 182
    $chkUseMarkerRow.Width = 360
    $chkUseMarkerRow.Height = 18

    $lblHeaders = $designer.Controls.Add("Forms.Label.1", "lblHeaders", $true)
    $lblHeaders.Caption = "통합할 헤더 목록"
    $lblHeaders.Left = 286
    $lblHeaders.Top = 214
    $lblHeaders.Width = 180
    $lblHeaders.Height = 18

    $lblHeadersHelp = $designer.Controls.Add("Forms.Label.1", "lblHeadersHelp", $true)
    $lblHeadersHelp.Caption = "쉼표 또는 줄바꿈으로 입력하세요. 표시행 기준을 켜면 표시된 열에서 자동 생성되며, 이 목록은 선택사항입니다."
    $lblHeadersHelp.Left = 286
    $lblHeadersHelp.Top = 234
    $lblHeadersHelp.Width = 430
    $lblHeadersHelp.Height = 30
    $lblHeadersHelp.WordWrap = $true

    $txtOutputHeaders = $designer.Controls.Add("Forms.TextBox.1", "txtOutputHeaders", $true)
    $txtOutputHeaders.Left = 286
    $txtOutputHeaders.Top = 270
    $txtOutputHeaders.Width = 446
    $txtOutputHeaders.Height = 228
    $txtOutputHeaders.MultiLine = $true
    $txtOutputHeaders.ScrollBars = 2
    $txtOutputHeaders.EnterKeyBehavior = $true
    $txtOutputHeaders.WordWrap = $false

    $cmdOK = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOK.Caption = "매칭 확인"
    $cmdOK.Left = 544
    $cmdOK.Top = 548
    $cmdOK.Width = 90
    $cmdOK.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 646
    $cmdCancel.Top = 548
    $cmdCancel.Width = 86
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean

Public Sub Configure(ByVal sheetNamesText As String, ByVal defaultKeyHeader As String, _
                     ByVal defaultOutputHeaders As String, ByVal defaultOutputSheetName As String, _
                     ByVal defaultOutputPosition As String, ByVal defaultHeaderRowSetting As String, _
                     ByVal defaultHeaderMode As String, ByVal defaultIdMode As String, _
                     ByVal defaultKeyMode As String, ByVal defaultUseMarkerRow As Boolean)
    Cancelled = True
    lstSheets.Clear

    Dim lines As Variant
    lines = Split(Replace(sheetNamesText, vbCr, vbLf), vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim sheetName As String
        sheetName = Trim$(CStr(lines(i)))
        If Len(sheetName) > 0 Then
            lstSheets.AddItem sheetName
            lstSheets.Selected(lstSheets.ListCount - 1) = True
        End If
    Next i

    txtKeyHeader.Text = defaultKeyHeader
    chkManualKeyColumn.Value = (LCase$(Trim$(defaultKeyMode)) = "manual")
    txtOutputHeaders.Text = HeadersForDisplay(defaultOutputHeaders)
    txtOutputSheetName.Text = defaultOutputSheetName
    txtHeaderRow.Text = IIf(Len(Trim$(defaultHeaderRowSetting)) > 0, defaultHeaderRowSetting, "자동")

    cmbPosition.Clear
    cmbPosition.AddItem "맨 뒤"
    cmbPosition.AddItem "맨 앞"
    cmbPosition.AddItem "현재 시트 뒤"
    If Len(defaultOutputPosition) > 0 Then
        cmbPosition.Value = defaultOutputPosition
    Else
        cmbPosition.Value = "맨 뒤"
    End If

    cmbHeaderMode.Clear
    cmbHeaderMode.AddItem "1줄 헤더"
    cmbHeaderMode.AddItem "위=표시명, 아래=변수"
    cmbHeaderMode.AddItem "이전=표시명, 선택=변수"

    Select Case defaultHeaderMode
        Case "display_next_key"
            cmbHeaderMode.Value = "위=표시명, 아래=변수"
        Case "key_prev_display"
            cmbHeaderMode.Value = "이전=표시명, 선택=변수"
        Case Else
            cmbHeaderMode.Value = "1줄 헤더"
    End Select

    cmbIdMode.Clear
    cmbIdMode.AddItem "ID 없으면 결과 시트에만 자동 생성"
    cmbIdMode.AddItem "기존 KEY 헤더 사용"
    cmbIdMode.AddItem "ID 없으면 원본 시트에 ID열 생성"
    Select Case defaultIdMode
        Case "existing"
            cmbIdMode.Value = "기존 KEY 헤더 사용"
        Case "source_auto"
            cmbIdMode.Value = "ID 없으면 원본 시트에 ID열 생성"
        Case Else
            cmbIdMode.Value = "ID 없으면 결과 시트에만 자동 생성"
    End Select
    chkUseMarkerRow.Value = defaultUseMarkerRow
End Sub

Public Property Get SelectedSheets() As String
    Dim result As String
    Dim i As Long
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then
            If Len(result) > 0 Then result = result & vbLf
            result = result & CStr(lstSheets.List(i))
        End If
    Next i
    SelectedSheets = result
End Property

Public Property Get KeyHeader() As String
    KeyHeader = Trim$(txtKeyHeader.Text)
End Property

Public Property Get OutputHeaders() As String
    OutputHeaders = Trim$(txtOutputHeaders.Text)
End Property

Public Property Get OutputSheetName() As String
    OutputSheetName = Trim$(txtOutputSheetName.Text)
End Property

Public Property Get OutputPosition() As String
    OutputPosition = CStr(cmbPosition.Value)
End Property

Public Property Get HeaderRowSetting() As String
    HeaderRowSetting = Trim$(txtHeaderRow.Text)
    If Len(HeaderRowSetting) = 0 Then HeaderRowSetting = "자동"
End Property

Public Property Get HeaderMode() As String
    Select Case CStr(cmbHeaderMode.Value)
        Case "위=표시명, 아래=변수"
            HeaderMode = "display_next_key"
        Case "이전=표시명, 선택=변수"
            HeaderMode = "key_prev_display"
        Case Else
            HeaderMode = "single"
    End Select
End Property

Public Property Get IdMode() As String
    Select Case CStr(cmbIdMode.Value)
        Case "기존 KEY 헤더 사용"
            IdMode = "existing"
        Case "ID 없으면 원본 시트에 ID열 생성"
            IdMode = "source_auto"
        Case Else
            IdMode = "output_auto"
    End Select
End Property

Public Property Get KeyMode() As String
    If CBool(chkManualKeyColumn.Value) Then
        KeyMode = "manual"
    Else
        KeyMode = "header"
    End If
End Property

Public Property Get UseMarkerRow() As Boolean
    UseMarkerRow = CBool(chkUseMarkerRow.Value)
End Property

Private Function HeadersForDisplay(ByVal headerText As String) As String
    Dim t As String
    t = Replace(headerText, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, ",", vbLf)
    t = Replace(t, ";", vbLf)
    HeadersForDisplay = Replace(t, vbLf, vbCrLf)
End Function

Private Function SelectedSheetCount() As Long
    Dim i As Long
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then SelectedSheetCount = SelectedSheetCount + 1
    Next i
End Function

Private Sub cmdAll_Click()
    Dim i As Long
    For i = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(i) = True
    Next i
End Sub

Private Sub cmdNone_Click()
    Dim i As Long
    For i = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(i) = False
    Next i
End Sub

Private Sub cmdOK_Click()
    If SelectedSheetCount() = 0 Then
        MsgBox "통합할 시트를 하나 이상 선택하세요.", vbExclamation, "시트 데이터 통합"
        lstSheets.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtKeyHeader.Text)) = 0 And IdMode = "existing" Then
        MsgBox "기존 KEY 헤더 사용 방식에서는 메인 KEY 헤더를 입력하세요.", vbExclamation, "시트 데이터 통합"
        txtKeyHeader.SetFocus
        Exit Sub
    End If
    If Not HeaderRowSettingIsValid() Then
        MsgBox "헤더 행은 '자동', '0', 또는 1 이상의 행 번호로 입력하세요.", vbExclamation, "시트 데이터 통합"
        txtHeaderRow.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtOutputHeaders.Text)) = 0 And Not CBool(chkUseMarkerRow.Value) Then
        MsgBox "표시행 기준을 사용하지 않을 때는 통합할 헤더를 하나 이상 입력하세요.", vbExclamation, "시트 데이터 통합"
        txtOutputHeaders.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtOutputSheetName.Text)) = 0 Then
        MsgBox "결과 시트 이름을 입력하세요.", vbExclamation, "시트 데이터 통합"
        txtOutputSheetName.SetFocus
        Exit Sub
    End If

    Cancelled = False
    Hide
End Sub

Private Function HeaderRowSettingIsValid() As Boolean
    Dim t As String
    t = LCase$(Trim$(txtHeaderRow.Text))
    If Len(t) = 0 Or t = "자동" Or t = "auto" Or t = "0" Then
        HeaderRowSettingIsValid = True
        Exit Function
    End If

    If Not IsNumeric(t) Then Exit Function
    HeaderRowSettingIsValid = (CLng(t) > 0)
End Function

Private Sub cmdCancel_Click()
    Cancelled = True
    Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancelled = True
End Sub
'@
    $form.CodeModule.AddFromString($formCode)
}

function Add-DatabaseMergeReviewUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmDatabaseMergeReview"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "시트 데이터 통합 매칭 확인"
    $form.Properties.Item("Width").Value = 760
    $form.Properties.Item("Height").Value = 560

    $lblGuide = $designer.Controls.Add("Forms.Label.1", "lblGuide", $true)
    $lblGuide.Caption = "설정된 헤더 행/구조, 마지막 행, 시트별 컬럼 매칭을 확인하세요. '없음'으로 표시된 헤더는 해당 시트에서 빈값으로 생성됩니다."
    $lblGuide.Left = 12
    $lblGuide.Top = 10
    $lblGuide.Width = 720
    $lblGuide.Height = 30
    $lblGuide.WordWrap = $true

    $lblSummary = $designer.Controls.Add("Forms.Label.1", "lblSummary", $true)
    $lblSummary.Caption = "통합 계획"
    $lblSummary.Left = 12
    $lblSummary.Top = 52
    $lblSummary.Width = 180
    $lblSummary.Height = 18

    $txtSummary = $designer.Controls.Add("Forms.TextBox.1", "txtSummary", $true)
    $txtSummary.Left = 12
    $txtSummary.Top = 74
    $txtSummary.Width = 720
    $txtSummary.Height = 150
    $txtSummary.MultiLine = $true
    $txtSummary.ScrollBars = 2
    $txtSummary.Locked = $true

    $lblMapping = $designer.Controls.Add("Forms.Label.1", "lblMapping", $true)
    $lblMapping.Caption = "헤더 매칭 검토"
    $lblMapping.Left = 12
    $lblMapping.Top = 238
    $lblMapping.Width = 180
    $lblMapping.Height = 18

    $txtMapping = $designer.Controls.Add("Forms.TextBox.1", "txtMapping", $true)
    $txtMapping.Left = 12
    $txtMapping.Top = 260
    $txtMapping.Width = 720
    $txtMapping.Height = 206
    $txtMapping.MultiLine = $true
    $txtMapping.ScrollBars = 3
    $txtMapping.Locked = $true
    $txtMapping.WordWrap = $false

    $cmdRun = $designer.Controls.Add("Forms.CommandButton.1", "cmdRun", $true)
    $cmdRun.Caption = "통합 실행"
    $cmdRun.Left = 544
    $cmdRun.Top = 486
    $cmdRun.Width = 90
    $cmdRun.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 646
    $cmdCancel.Top = 486
    $cmdCancel.Width = 86
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean

Public Sub Configure(ByVal summaryText As String, ByVal mappingText As String)
    Cancelled = True
    txtSummary.Text = summaryText
    txtMapping.Text = mappingText
End Sub

Private Sub cmdRun_Click()
    Cancelled = False
    Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancelled = True
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
    $form.Properties.Item("Height").Value = 540

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
    $form.Properties.Item("Width").Value = 580
    $form.Properties.Item("Height").Value = 540

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "적용할 정리 규칙을 선택하세요."
    $lblTitle.Left = 12
    $lblTitle.Top = 10
    $lblTitle.Width = 540
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
        @("chkCutBeforePossessive", "'의 ' 앞부분만 남기기", 250, 120, $false),
        @("chkPronounceEnglishWords", "영어 단어를 한글 발음으로 변환", 250, 146, $false),
        @("chkSpellAlphabetKorean", "알파벳을 한글 읽기로 변환", 250, 172, $true),
        @("chkRemoveSpecialChars", "괄호·특수문자 제거", 250, 198, $false),
        @("chkRemoveBranchMarkers", "본점/지점/점포명 제거", 250, 224, $false),
        @("chkRemoveLegalEntityMarkers", "Co.,Ltd./유한회사 등 법인표기 삭제", 250, 250, $false),
        @("chkRemoveParentheticalText", "괄호 안 문구까지 삭제", 250, 276, $false),
        @("chkRemoveBilingualDuplicates", "한글/영문 중복 병기 제거", 250, 302, $true)
    )

    foreach ($item in $labels) {
        $chk = $designer.Controls.Add("Forms.CheckBox.1", $item[0], $true)
        $chk.Caption = $item[1]
        $chk.Left = [int]$item[2]
        $chk.Top = [int]$item[3]
        $chk.Width = 300
        $chk.Height = 20
        $chk.Value = [bool]$item[4]
    }

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "기본값은 제공된 기업이름정리.xlsm의 체크 상태를 따릅니다. 결과는 원본 오른쪽 복사본 열에 생성됩니다."
    $lblTip.Left = 12
    $lblTip.Top = 318
    $lblTip.Width = 540
    $lblTip.Height = 42

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 374
    $cmdOk.Top = 448
    $cmdOk.Width = 86
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 474
    $cmdCancel.Top = 448
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

Public Property Get RemoveBranchMarkers() As Boolean
    RemoveBranchMarkers = CBool(chkRemoveBranchMarkers.Value)
End Property

Public Property Get RemoveLegalEntityMarkers() As Boolean
    RemoveLegalEntityMarkers = CBool(chkRemoveLegalEntityMarkers.Value)
End Property

Public Property Get RemoveSpecialChars() As Boolean
    RemoveSpecialChars = CBool(chkRemoveSpecialChars.Value)
End Property

Public Property Get RemoveParentheticalText() As Boolean
    RemoveParentheticalText = CBool(chkRemoveParentheticalText.Value)
End Property

Public Property Get PronounceEnglishWords() As Boolean
    PronounceEnglishWords = CBool(chkPronounceEnglishWords.Value)
End Property

Public Property Get SpellAlphabetKorean() As Boolean
    SpellAlphabetKorean = CBool(chkSpellAlphabetKorean.Value)
End Property

Public Property Get RemoveBilingualDuplicates() As Boolean
    RemoveBilingualDuplicates = CBool(chkRemoveBilingualDuplicates.Value)
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
    $form.Properties.Item("Width").Value = 560
    $form.Properties.Item("Height").Value = 320

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "선택한 행정동/법정동마다 업종을 검색해 신규 업체 리스트를 만듭니다."
    $lblTitle.Left = 12
    $lblTitle.Top = 12
    $lblTitle.Width = 520
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
    $txtIndustry.Width = 390
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
    $lblDelay.Left = 240
    $lblDelay.Top = 84
    $lblDelay.Width = 90
    $lblDelay.Height = 18

    $txtDelay = $designer.Controls.Add("Forms.TextBox.1", "txtDelay", $true)
    $txtDelay.Left = 338
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
    $cmbSort.Style = 2

    $lblDongBasis = $designer.Controls.Add("Forms.Label.1", "lblDongBasis", $true)
    $lblDongBasis.Caption = "동 기준"
    $lblDongBasis.Left = 300
    $lblDongBasis.Top = 120
    $lblDongBasis.Width = 58
    $lblDongBasis.Height = 18

    $cmbDongBasis = $designer.Controls.Add("Forms.ComboBox.1", "cmbDongBasis", $true)
    $cmbDongBasis.Left = 370
    $cmbDongBasis.Top = 116
    $cmbDongBasis.Width = 100
    $cmbDongBasis.Height = 22
    $cmbDongBasis.Style = 2

    $lblSido = $designer.Controls.Add("Forms.Label.1", "lblSido", $true)
    $lblSido.Caption = "광역시도"
    $lblSido.Left = 12
    $lblSido.Top = 154
    $lblSido.Width = 110
    $lblSido.Height = 18

    $cmbSido = $designer.Controls.Add("Forms.ComboBox.1", "cmbSido", $true)
    $cmbSido.Left = 128
    $cmbSido.Top = 150
    $cmbSido.Width = 160
    $cmbSido.Height = 22
    $cmbSido.Style = 2

    $lblSigungu = $designer.Controls.Add("Forms.Label.1", "lblSigungu", $true)
    $lblSigungu.Caption = "시군구"
    $lblSigungu.Left = 304
    $lblSigungu.Top = 154
    $lblSigungu.Width = 58
    $lblSigungu.Height = 18

    $cmbSigungu = $designer.Controls.Add("Forms.ComboBox.1", "cmbSigungu", $true)
    $cmbSigungu.Left = 370
    $cmbSigungu.Top = 150
    $cmbSigungu.Width = 148
    $cmbSigungu.Height = 22
    $cmbSigungu.Style = 2

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "내장 행정동/법정동 마스터 또는 현재 통합문서의 사용자 마스터 시트를 사용합니다. 광역시도/시군구를 전체로 두면 선택 기준 전체를 대상으로 합니다."
    $lblTip.Left = 12
    $lblTip.Top = 188
    $lblTip.Width = 520
    $lblTip.Height = 44

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 358
    $cmdOk.Top = 232
    $cmdOk.Width = 80
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 450
    $cmdCancel.Top = 232
    $cmdCancel.Width = 80
    $cmdCancel.Height = 28

$formCode = @'
Option Explicit

Private Const ALL_TEXT As String = "(전체)"
Public Cancelled As Boolean
Private regionOptionsText As String
Private loading As Boolean

Private Sub UserForm_Initialize()
    Cancelled = True
    cmbSort.AddItem "comment"
    cmbSort.AddItem "random"
    cmbSort.Value = "comment"
    cmbDongBasis.AddItem "행정동"
    cmbDongBasis.AddItem "법정동"
    cmbDongBasis.Value = "행정동"
    cmbSido.AddItem ALL_TEXT
    cmbSido.Value = ALL_TEXT
    cmbSigungu.AddItem ALL_TEXT
    cmbSigungu.Value = ALL_TEXT
End Sub

Public Sub Configure(ByVal defaultIndustry As String, ByVal defaultDisplayCount As Long, _
                     ByVal defaultDelaySeconds As Double, ByVal defaultSortMode As String, _
                     ByVal defaultDongBasis As String, ByVal defaultSido As String, _
                     ByVal defaultSigungu As String, ByVal regionOptions As String)
    loading = True
    regionOptionsText = regionOptions
    txtIndustry.Value = defaultIndustry
    txtDisplay.Value = CStr(defaultDisplayCount)
    txtDelay.Value = CStr(defaultDelaySeconds)

    If LCase$(Trim$(defaultSortMode)) = "random" Then
        cmbSort.Value = "random"
    Else
        cmbSort.Value = "comment"
    End If

    If LCase$(Trim$(defaultDongBasis)) = "legal" Then
        cmbDongBasis.Value = "법정동"
    Else
        cmbDongBasis.Value = "행정동"
    End If

    LoadSidoList defaultSido
    LoadSigunguList defaultSigungu
    loading = False
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

Public Property Get DongBasis() As String
    DongBasis = BasisValue()
End Property

Public Property Get SidoFilter() As String
    SidoFilter = CurrentSidoFilter()
End Property

Public Property Get SigunguFilter() As String
    If CStr(cmbSigungu.Value) = ALL_TEXT Then
        SigunguFilter = ""
    Else
        SigunguFilter = Trim$(CStr(cmbSigungu.Value))
    End If
End Property

Private Function BasisValue() As String
    If CStr(cmbDongBasis.Value) = "법정동" Then
        BasisValue = "legal"
    Else
        BasisValue = "admin"
    End If
End Function

Private Function CurrentSidoFilter() As String
    If CStr(cmbSido.Value) = ALL_TEXT Then
        CurrentSidoFilter = ""
    Else
        CurrentSidoFilter = Trim$(CStr(cmbSido.Value))
    End If
End Function

Private Sub cmbDongBasis_Change()
    If loading Then Exit Sub
    loading = True
    LoadSidoList ""
    LoadSigunguList ""
    loading = False
End Sub

Private Sub cmbSido_Change()
    If loading Then Exit Sub
    loading = True
    LoadSigunguList ""
    loading = False
End Sub

Private Sub LoadSidoList(ByVal defaultValue As String)
    cmbSido.Clear
    cmbSido.AddItem ALL_TEXT

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim lines As Variant
    lines = Split(regionOptionsText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim parts As Variant
        parts = Split(CStr(lines(i)), vbTab)
        If UBound(parts) >= 1 Then
            If CStr(parts(0)) = BasisValue() Then
                Dim sido As String
                sido = Trim$(CStr(parts(1)))
                If Len(sido) > 0 And Not seen.Exists(sido) Then
                    seen.Add sido, True
                    cmbSido.AddItem sido
                End If
            End If
        End If
    Next i

    If Len(Trim$(defaultValue)) > 0 Then AddComboItemIfMissing cmbSido, Trim$(defaultValue)
    If Len(Trim$(defaultValue)) > 0 And ComboHasValue(cmbSido, Trim$(defaultValue)) Then
        cmbSido.Value = Trim$(defaultValue)
    Else
        cmbSido.Value = ALL_TEXT
    End If
End Sub

Private Sub LoadSigunguList(ByVal defaultValue As String)
    cmbSigungu.Clear
    cmbSigungu.AddItem ALL_TEXT

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim selectedSido As String
    selectedSido = CurrentSidoFilter()

    Dim lines As Variant
    lines = Split(regionOptionsText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim parts As Variant
        parts = Split(CStr(lines(i)), vbTab)
        If UBound(parts) >= 2 Then
            If CStr(parts(0)) = BasisValue() Then
                If Len(selectedSido) = 0 Or CStr(parts(1)) = selectedSido Then
                    Dim sigungu As String
                    sigungu = Trim$(CStr(parts(2)))
                    If Len(sigungu) > 0 And Not seen.Exists(sigungu) Then
                        seen.Add sigungu, True
                        cmbSigungu.AddItem sigungu
                    End If
                End If
            End If
        End If
    Next i

    If Len(Trim$(defaultValue)) > 0 Then AddComboItemIfMissing cmbSigungu, Trim$(defaultValue)
    If Len(Trim$(defaultValue)) > 0 And ComboHasValue(cmbSigungu, Trim$(defaultValue)) Then
        cmbSigungu.Value = Trim$(defaultValue)
    Else
        cmbSigungu.Value = ALL_TEXT
    End If
End Sub

Private Function ComboHasValue(ByVal combo As Object, ByVal valueText As String) As Boolean
    Dim i As Long
    For i = 0 To combo.ListCount - 1
        If CStr(combo.List(i)) = valueText Then
            ComboHasValue = True
            Exit Function
        End If
    Next i
End Function

Private Sub AddComboItemIfMissing(ByVal combo As Object, ByVal valueText As String)
    If Not ComboHasValue(combo, valueText) Then combo.AddItem valueText
End Sub

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
    $lblTitle.Caption = "분석에 사용할 항목을 선택하세요. 범위는 실행 후 열리는 Excel 범위 선택창에서 지정합니다."
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
    $cmdPickName.Caption = "선택 안내"
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
    $cmdPickAddress.Caption = "선택 안내"
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
    $lblTip.Caption = "업체명만 또는 주소만 선택할 수 있습니다. 둘 다 선택하면 두 범위는 같은 시트와 같은 행 수여야 합니다. 실행 후 열리는 범위 선택창에서는 Shift+방향키, Ctrl+Shift+방향키, 직접 주소 입력을 사용할 수 있습니다."
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
    chkUseName.Value = True
    MsgBox "실행 버튼을 누른 뒤 업체명 범위 선택창이 열립니다." & vbCrLf & _
           "그 창에서 Shift+방향키, Ctrl+Shift+방향키, 직접 주소 입력을 사용할 수 있습니다." & vbCrLf & _
           "범위 주소를 알고 있으면 현재 입력칸에 직접 입력해도 됩니다.", vbInformation, Me.Caption
    txtNameRange.SetFocus
End Sub

Private Sub cmdPickAddress_Click()
    chkUseAddress.Value = True
    MsgBox "실행 버튼을 누른 뒤 주소 범위 선택창이 열립니다." & vbCrLf & _
           "그 창에서 Shift+방향키, Ctrl+Shift+방향키, 직접 주소 입력을 사용할 수 있습니다." & vbCrLf & _
           "범위 주소를 알고 있으면 현재 입력칸에 직접 입력해도 됩니다.", vbInformation, Me.Caption
    txtAddressRange.SetFocus
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

Public Property Get NameRangeAddress() As String
    NameRangeAddress = Trim$(txtNameRange.Value)
End Property

Public Property Get AddressRangeAddress() As String
    AddressRangeAddress = Trim$(txtAddressRange.Value)
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

function Add-LegalDongMasterWorksheet {
    param(
        $Workbook,
        [string]$CsvPath
    )

    if (-not (Test-Path -LiteralPath $CsvPath)) {
        throw "Legal dong master CSV not found: $CsvPath"
    }

    foreach ($sheet in @($Workbook.Worksheets)) {
        if ($sheet.Name -eq "_NaverLocalLegalDongMaster") {
            $sheet.Delete()
            break
        }
    }

    $ws = $Workbook.Worksheets.Add()
    $ws.Name = "_NaverLocalLegalDongMaster"

    $query = $ws.QueryTables.Add("TEXT;$CsvPath", $ws.Range("A1"))
    $query.TextFileParseType = 1
    $query.TextFileCommaDelimiter = $true
    $query.TextFilePlatform = 65001
    $query.Refresh($false) | Out-Null
    $query.Delete()

    $lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    $lastCol = $ws.Cells.Item(1, $ws.Columns.Count).End(-4159).Column
    $useCol = $lastCol + 1
    $ws.Cells.Item(1, $useCol).Value2 = "사용여부"
    if ($lastRow -ge 2) {
        $ws.Range($ws.Cells.Item(2, $useCol), $ws.Cells.Item($lastRow, $useCol)).Value2 = "Y"
    }

    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $useCol)).Font.Bold = $true
    $ws.Columns.AutoFit() | Out-Null
    $ws.Visible = 2
}

function Add-ProgressStatusUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmProgressStatus"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "작업 진행 현황"
    $form.Properties.Item("Width").Value = 420
    $form.Properties.Item("Height").Value = 210

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

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "작업 취소"
    $cmdCancel.Left = 294
    $cmdCancel.Top = 142
    $cmdCancel.Width = 86
    $cmdCancel.Height = 26

    $formCode = @'
Option Explicit

Private mTotal As Long
Private Const BAR_MAX_WIDTH As Single = 360
Public CancelRequested As Boolean

Public Sub Configure(ByVal title As String, ByVal totalCount As Long, ByVal message As String)
    Me.Caption = title
    lblTitle.Caption = title
    mTotal = totalCount
    CancelRequested = False
    cmdCancel.Enabled = True
    cmdCancel.Caption = "작업 취소"
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

Public Property Get IsCancelled() As Boolean
    IsCancelled = CancelRequested
End Property

Private Sub cmdCancel_Click()
    CancelRequested = True
    cmdCancel.Enabled = False
    cmdCancel.Caption = "취소 요청됨"
    lblWait.Caption = "현재 처리 단위가 끝나면 중단합니다."
    Repaint
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        CancelRequested = True
        Cancel = True
        cmdCancel.Enabled = False
        cmdCancel.Caption = "취소 요청됨"
        lblWait.Caption = "현재 처리 단위가 끝나면 중단합니다."
        Repaint
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
          <button id="btnLaunchAddressSearch" label="주소/검색" size="large" imageMso="AddressBook" onAction="Ribbon_LaunchAddressSearch" screentip="주소/검색 통합 실행" supertip="주소 유형 판별, 주소 체계 분리/통합, 행정구역 분해, 네이버 주소 검증과 신규 리스트 생성을 하나의 선택 창에서 실행합니다." />
          <button id="btnLaunchRosterQuality" label="명부 품질" size="large" imageMso="ReviewAcceptChange" onAction="Ribbon_LaunchRosterQuality" screentip="명부 품질 통합 실행" supertip="품질 진단, 중복 탐지, 업체명 매칭, 전화번호 검사, 사업자등록 상태조회를 하나의 선택 창에서 실행합니다." />
          <button id="btnLaunchSpssTools" label="SPSS 표 도구" size="large" imageMso="CreateReport" onAction="Ribbon_LaunchSpssTools" screentip="SPSS 표 도구 통합 실행" supertip="목차 생성, 표/배너 재구성, 세로 배너 삭제, 표 메타 추출을 하나의 선택 창에서 실행합니다." />
          <button id="btnLaunchSampleDesign" label="표본설계" size="large" imageMso="PivotTableInsert" onAction="Ribbon_LaunchSampleDesign" screentip="표본설계 통합 실행" supertip="표본설계 시트 생성, 구성표 리스트화, 리스트→집계표, 복합 목표그룹 할당 계산, 진행률 관리표 생성을 하나의 선택 창에서 실행합니다." />
        </group>
        <group id="grpNumberTools" label="값 정리">
          <button id="btnConvertRange" label="숫자 변환" size="large" imageMso="NumberFormatGeneral" onAction="Ribbon_ConvertTextNumbersInColumns" screentip="선택 범위 숫자 변환" supertip="선택한 범위에서 텍스트 형식 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnConvertSheet" label="시트 숫자 변환" imageMso="TableSelect" onAction="Ribbon_ConvertTextNumbersInUsedRange" screentip="시트 전체 숫자 변환" supertip="현재 시트에서 실제 데이터가 있는 마지막 행과 열까지만 잡아 텍스트 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnPhoneRange" label="전화번호 정리" size="large" imageMso="ContactPictureMenu" onAction="Ribbon_NormalizePhonesInColumns" screentip="전화번호 하이픈 정리" supertip="선택 범위 오른쪽에 복사본 열을 만들고 한국 전화번호를 하이픈 포함 표준 형식으로 정리합니다." />
          <button id="btnPhoneDashless" label="전화번호 숫자만" imageMso="ClearFormats" onAction="Ribbon_RemovePhoneDashes" screentip="전화번호 숫자만 남기기" supertip="선택한 전화번호 범위 오른쪽에 복사본 열을 만들고 대시와 구분기호를 제거해 숫자만 남깁니다." />
          <button id="btnCompanyNameClean" label="기업명 정리" size="large" imageMso="TextBoxInsert" onAction="Ribbon_CompanyNameClean" screentip="기업체명/사업체명 정리" supertip="선택 범위 오른쪽에 복사본 열을 만들고 공백, 회사형태 표기, 괄호·특수문자, 영어 단어 한글 발음, 알파벳 한글 읽기 등 선택한 규칙으로 기업체명/사업체명을 정리합니다." />
          <button id="btnAddressConvert" label="주소 유형 판별" size="large" imageMso="AddressBook" onAction="Ribbon_AddressConvertOldToRoad" screentip="도로명/동 기준 주소 판별" supertip="선택한 주소 범위 오른쪽에 도로명주소인지 법정/행정동 기준 주소인지 판별 결과를 생성합니다." />
          <button id="btnAddressSplit" label="주소 체계 분리" imageMso="AddressBook" onAction="Ribbon_AddressSplitDongAndRoad" screentip="동 기준/도로명 주소 분리" supertip="주소 범위를 선택하면 원본 오른쪽에 동 기준 주소와 도로명주소 2개 열을 추가합니다. 변환표가 있으면 지번/도로명 양방향 매칭을 사용합니다." />
          <button id="btnAddressCombine" label="주소 통합" imageMso="CellsMerge" onAction="Ribbon_AddressCombineDongAndRoad" screentip="분할 주소 통합" supertip="동 기준 주소 범위와 도로명주소 범위를 선택하고 우선순위를 지정해 오른쪽에 통합주소 열을 생성합니다." />
        </group>
        <group id="grpBusinessNoTools" label="사업자등록">
          <button id="btnBizApiKey" label="인증키 설정" imageMso="ServerConnection" onAction="Ribbon_BizNoSetApiKey" screentip="사업자등록 API 인증키 설정" supertip="공공데이터포털 국세청 사업자등록정보 API 인증키를 저장하거나 변경합니다." />
          <button id="btnBizDashless" label="번호 숫자만" imageMso="ClearFormats" onAction="Ribbon_RemoveBizNoDashes" screentip="사업자등록번호 숫자만 남기기" supertip="선택한 사업자등록번호 범위 오른쪽에 복사본 열을 만들고 대시와 구분기호를 제거해 숫자 10자리 형태로 정리합니다." />
          <button id="btnBizMerge" label="3열 병합" imageMso="CellsMerge" onAction="Ribbon_BizNoMergeThreeColumns" screentip="사업자등록번호 3열 병합" supertip="3개 열로 나뉜 사업자등록번호를 000-00-00000 형식의 새 열로 병합합니다." />
          <button id="btnBizStatus" label="상태조회" size="large" imageMso="DataRefreshAll" onAction="Ribbon_BizNoQueryStatus" screentip="사업자등록 상태조회" supertip="선택한 사업자등록번호 범위를 국세청 API로 조회하고 오른쪽에 결과 열을 추가합니다." />
        </group>
        <group id="grpNaverLocalTools" label="네이버 지역검색">
          <button id="btnNaverApiKey" label="인증키 설정" imageMso="ServerConnection" onAction="Ribbon_NaverLocalSetApiCredentials" screentip="네이버 지역검색 API 인증키 설정" supertip="네이버 Developers의 Client ID와 Client Secret을 저장하거나 변경합니다." />
          <button id="btnNaverBuildList" label="신규 리스트" size="large" imageMso="AddressBook" onAction="Ribbon_NaverLocalBuildListByDong" screentip="행정동/법정동 기준 업체 리스트 생성" supertip="검색 기준을 행정동 또는 법정동으로 선택하고 업종을 입력해 네이버 지역검색 결과를 신규 시트에 생성합니다." />
          <button id="btnNaverVerifyAddress" label="주소 검증" size="large" imageMso="ReviewCompareTwoVersions" onAction="Ribbon_NaverLocalVerifyExistingData" screentip="업체명/주소 검증" supertip="업체명 범위와 주소 범위를 선택하면 네이버 지역검색 결과, 유사도, 검증등급 열을 추가합니다." />
          <button id="btnNaverOfficialMall" label="공식몰 검색" size="large" imageMso="HyperlinkInsert" onAction="Ribbon_NaverOfficialMallFindLinks" screentip="자체 공식몰 링크 검색" supertip="업체명 범위를 선택하면 네이버 웹문서 API로 '업체명 공식몰'을 검색하고, 네이버 브랜드스토어/스마트스토어/오픈마켓을 제외한 자체 쇼핑몰 후보 URL을 추가합니다." />
          <button id="btnNaverDongMaster" label="행정동 양식" imageMso="TableInsert" onAction="Ribbon_NaverLocalCreateDongMasterTemplate" screentip="행정동 마스터 양식 생성" supertip="현재 통합문서에 내장 행정동 마스터와 같은 양식의 행정동_마스터 시트를 생성합니다. 이 시트를 수정하면 신규 리스트 생성에 우선 사용됩니다." />
          <button id="btnNaverLegalDongMaster" label="법정동 양식" imageMso="TableInsert" onAction="Ribbon_NaverLocalCreateLegalDongMasterTemplate" screentip="법정동 마스터 양식 생성" supertip="현재 통합문서에 내장 법정동 마스터와 같은 양식의 법정동_마스터 시트를 생성합니다. 이 시트를 수정하면 법정동 기준 신규 리스트 생성에 우선 사용됩니다." />
        </group>
        <group id="grpDatabaseTools" label="데이터 통합">
          <button id="btnDatabaseMergeSheets" label="시트 통합" size="large" imageMso="TableInsert" onAction="Ribbon_DatabaseMergeSheets" screentip="여러 시트 데이터 통합" supertip="KEY와 통합 헤더를 지정하고 시트별 헤더 매핑을 검증한 뒤 여러 시트의 데이터를 하나의 통합 시트로 쌓습니다." />
          <button id="btnDatabaseCreateMarkerRows" label="통합 표시행" imageMso="InsertRows" onAction="Ribbon_DatabaseCreateMergeMarkerRows" screentip="통합할 열 표시행 만들기" supertip="선택한 시트의 헤더 위에 표시행을 추가합니다. 통합할 열 위 셀에 x, 1 등 아무 값을 입력한 뒤 시트 통합에서 표시행 기준을 선택하세요." />
          <button id="btnDatabaseSplitMergedSheet" label="통합 시트 분리" imageMso="TableSelect" onAction="Ribbon_DatabaseSplitMergedSheet" screentip="통합 데이터 원본별 분리" supertip="통합 시트 1행의 관리코드를 기준으로 원본시트별 분리 시트를 다시 생성합니다." />
        </group>
        <group id="grpSpssTableTools" label="SPSS 표 정리">
          <button id="btnSpssCreateIndex" label="목차 생성" imageMso="TableOfContentsGallery" onAction="Ribbon_SPSSCreateIndexForActiveSheet" screentip="SPSS 표 목차 생성" supertip="목차가 없는 집계표 시트에서 [로 시작하는 표 제목 행을 찾아 새 목차 시트와 링크를 생성합니다." />
          <button id="btnSpssRebuildFromIndex" label="표/배너 재구성" size="large" imageMso="CreateReport" onAction="Ribbon_SPSSRebuildTablesFromIndex" screentip="SPSS 표/배너 재구성" supertip="표 목록을 드래그하거나 버튼으로 이동하고, 삭제할 표/배너를 선택해 새 집계표와 목차 링크를 생성합니다." />
          <button id="btnSpssDeleteVertical" label="세로 배너 삭제" imageMso="DeleteRows" onAction="Ribbon_SPSSDeleteVerticalBanners" screentip="세로 배너 삭제" supertip="원본 집계표 시트를 복사한 뒤 ▣ 세로 배너와 BOT/TOP/MID 계열 세로형 배너 행을 삭제하고 새 목차를 생성합니다." />
          <button id="btnSpssMetadata" label="표 메타 추출" imageMso="TablePropertiesDialog" onAction="Ribbon_SPSSExtractTableMetadata" screentip="SPSS 표 메타 추출" supertip="집계표의 표 제목, 행 위치, 배너 목록, 이동 링크를 새 시트로 추출합니다." />
        </group>
        <group id="grpRosterQualityTools" label="명부 품질">
          <button id="btnRosterValidationSheet" label="검증 시트" size="large" imageMso="TableInsert" onAction="Ribbon_RosterValidationSheet" screentip="명부 검증용 표준 시트 생성" supertip="ID, 사업체명/기업체명/이름, 연락처1~3, 주소1~3, 이메일, 직무/직책/소속부서/기타정보 헤더가 있는 표준 입력 시트를 만들거나 기존 명부를 자동 인식해 구성합니다." />
          <button id="btnRosterQuality" label="품질 진단" size="large" imageMso="ReviewAcceptChange" onAction="Ribbon_RosterQualityReport" screentip="명부 품질 진단 리포트" supertip="선택한 명부에서 빈값, 전화번호, 주소 유형, 사업자번호 상태를 요약한 리포트 시트를 만듭니다." />
          <button id="btnRosterDuplicate" label="중복 탐지" size="large" imageMso="ReviewCompareTwoVersions" onAction="Ribbon_RosterDetectDuplicates" screentip="중복 업체 탐지" supertip="업체명, 주소, 전화번호, 사업자번호를 조합해 중복 의심 쌍과 점수를 새 시트에 생성합니다." />
          <button id="btnRosterCompanyMatch" label="업체명 매칭" imageMso="TextBoxInsert" onAction="Ribbon_RosterMatchCompanyNames" screentip="업체명 유사그룹 생성" supertip="업체명을 표준화하고 같은 표준화명을 유사그룹으로 묶는 열을 생성합니다." />
          <button id="btnAddressParts" label="행정구역 분해" imageMso="AddressBook" onAction="Ribbon_AddressSplitAdministrativeParts" screentip="주소 행정구역 분해" supertip="주소를 광역시도, 시군구, 읍면동, 상세주소 열로 분해합니다." />
          <button id="btnPhoneValidate" label="전화번호 검사" imageMso="ContactPictureMenu" onAction="Ribbon_PhoneValidateSelection" screentip="전화번호 유효성 검사" supertip="전화번호 범위 오른쪽에 번호 유형과 정상/확인 필요 결과를 생성합니다." />
        </group>
        <group id="grpOutlierTools" label="이상치 탐지">
          <button id="btnOutlierRun" label="분석 열 선택" size="large" imageMso="ConditionalFormattingHighlightCellsRules" onAction="Ribbon_OutlierDetectRun" screentip="이상치 확인 범위 선택" supertip="실행 전 2차 자료/전년도 자료 대체 여부를 확인하고, 필요하면 이상치 탐지 작업 시트를 만든 뒤 범위 검증, 기초통계량표, Q1/Q3 기준값 열, 원본 검토복사 열, 조건부 서식을 생성합니다." />
        </group>
        <group id="grpSampleDesignTools" label="표본설계">
          <button id="btnSampleDesignListify" label="구성표 리스트화" size="large" imageMso="TableInsert" onAction="Ribbon_SampleDesignListifyPopulationTable" screentip="모집단 구성표 리스트화" supertip="자동 감지 또는 범위 직접 지정을 선택하고, 미리보기 창에서 행축/가로축/값 구조를 확인한 뒤 리스트형 데이터로 변환합니다." />
          <button id="btnSampleDesignPivotList" label="리스트→집계표" imageMso="PivotTableInsert" onAction="Ribbon_SampleDesignPivotPopulationList" screentip="리스트형 표본자료를 집계표로 변환" supertip="구성표 리스트화 결과 또는 할당결과 범위를 선택하면 행축, 열축, 값 열을 자동 제안하고 표제목별 집계표 블록을 생성합니다." />
          <button id="btnSampleDesignTemplate" label="설계 시트 생성" size="large" imageMso="TableInsert" onAction="Ribbon_SampleDesignCreateTemplate" screentip="표본설계 입력 시트 생성" supertip="리스트형 또는 집계표형 입력을 선택하고 동적 쿼터, 목표표본 그룹표, 배분 방식을 입력할 표본설계 시트를 생성합니다." />
          <button id="btnSampleDesignGuide" label="표본설계 가이드" imageMso="Info" onAction="Ribbon_SampleDesignShowGuide" screentip="표본설계 사용 가이드" supertip="표본설계 입력 흐름, 필수 입력값, 목표표본 그룹표 사용법, 자주 나는 오류와 조치를 안내하는 가이드 시트를 엽니다." />
          <button id="btnSampleDesignBandColumn" label="구간 만들기" imageMso="GroupInsert" onAction="Ribbon_SampleDesignCreateBandColumn" screentip="숫자값을 표본설계용 구간 변수로 변환" supertip="매출액, 종사자수, 인구수 같은 숫자 열을 선택하고 UserForm에서 구간 시작/종료값과 포함 조건을 지정하면 오른쪽에 구간 열을 생성하고 설정을 숨김 시트에 저장합니다." />
          <button id="btnSampleDesignAllocate" label="할당 계산" size="large" imageMso="PivotTableInsert" onAction="Ribbon_SampleDesignRunAllocation" screentip="표본할당 계산" supertip="표본설계_입력 시트를 읽어 단순 비례, 제곱근 비례, 멱배분 등 선택한 방식으로 복합 목표그룹별 최종 표본할당표와 검토 시트를 생성합니다." />
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
$LegalDongMasterCsv = Join-Path $ProjectRoot "resources\legal_dong_master.csv"
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
$tempLatestAddin = $null
$tempReleaseAddin = $null

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
$publishStage = "start"
try {
    $publishStage = "sanitize versioned workbook API credentials"
    Clear-WorkbookFileStoredApiSecrets -Excel $excel -Path $ReleaseWorkbook

    $workbook = $excel.Workbooks.Open($DevWorkbook)
    try {
        $publishStage = "remove old VBA components"
        foreach ($component in @($workbook.VBProject.VBComponents)) {
            if (($component.Type -eq 1 -and ($component.Name -eq "MacroToolkit" -or $component.Name -like "Module*")) -or
                ($component.Type -eq 3 -and ($component.Name -eq "frmSpssRebuildOptions" -or $component.Name -eq "frmRangeSelectOptions" -or $component.Name -eq "frmCompanyNameOptions" -or $component.Name -eq "frmBizStatusFieldsOptions" -or $component.Name -eq "frmNaverLocalBuildOptions" -or $component.Name -eq "frmNaverLocalVerifyOptions" -or $component.Name -eq "frmProgressStatus" -or $component.Name -eq "frmToolkitButtonMenu" -or $component.Name -eq "frmToolkitCheckListMenu" -or $component.Name -eq "frmSampleDesignSetup" -or $component.Name -eq "frmSampleDesignBandBuilder" -or $component.Name -eq "frmPopulationListifyPreview" -or $component.Name -eq "frmDatabaseMergeSetup" -or $component.Name -eq "frmDatabaseMergeReview"))) {
                $workbook.VBProject.VBComponents.Remove($component)
            }
        }
        $publishStage = "import MacroToolkit module"
        $imported = $workbook.VBProject.VBComponents.Import($ImportModulePath)
        $imported.Name = "MacroToolkit"
        $publishStage = "add frmToolkitButtonMenu"
        Add-ToolkitButtonMenuUserForm -Workbook $workbook
        $publishStage = "add frmToolkitCheckListMenu"
        Add-ToolkitCheckListMenuUserForm -Workbook $workbook
        $publishStage = "add frmSampleDesignSetup"
        Add-SampleDesignSetupUserForm -Workbook $workbook
        $publishStage = "add frmSampleDesignBandBuilder"
        Add-SampleDesignBandBuilderUserForm -Workbook $workbook
        $publishStage = "add frmPopulationListifyPreview"
        Add-PopulationListifyPreviewUserForm -Workbook $workbook
        $publishStage = "add frmDatabaseMergeSetup"
        Add-DatabaseMergeSetupUserForm -Workbook $workbook
        $publishStage = "add frmDatabaseMergeReview"
        Add-DatabaseMergeReviewUserForm -Workbook $workbook
        $publishStage = "add frmSpssRebuildOptions"
        Add-SpssRebuildUserForm -Workbook $workbook
        $publishStage = "add frmCompanyNameOptions"
        Add-CompanyNameOptionsUserForm -Workbook $workbook
        $publishStage = "add frmBizStatusFieldsOptions"
        Add-BizStatusFieldsOptionsUserForm -Workbook $workbook
        $publishStage = "add frmNaverLocalBuildOptions"
        Add-NaverLocalBuildOptionsUserForm -Workbook $workbook
        $publishStage = "add frmNaverLocalVerifyOptions"
        Add-NaverLocalVerifyOptionsUserForm -Workbook $workbook
        $publishStage = "add frmProgressStatus"
        Add-ProgressStatusUserForm -Workbook $workbook
        $publishStage = "add admin dong master"
        Add-AdminDongMasterWorksheet -Workbook $workbook -CsvPath $AdminDongMasterCsv
        $publishStage = "add legal dong master"
        Add-LegalDongMasterWorksheet -Workbook $workbook -CsvPath $LegalDongMasterCsv
    }
    catch {
        throw "Could not import VBA module into .xlam. Stage: $publishStage. Enable Excel Trust Center > Macro Settings > Trust access to the VBA project object model. Details: $($_.Exception.Message)"
    }

    $workbook.IsAddin = $true
    $tempLatestAddin = Join-Path $LatestDir ("_tmp_" + $LatestAddinName)
    $tempReleaseAddin = Join-Path $VersionDir ("_tmp_" + $ReleaseAddinName)
    if (Test-Path -LiteralPath $tempLatestAddin) { Remove-Item -LiteralPath $tempLatestAddin -Force }
    if (Test-Path -LiteralPath $tempReleaseAddin) { Remove-Item -LiteralPath $tempReleaseAddin -Force }

    $publishStage = "save latest add-in with local credentials"
    $workbook.SaveAs($tempLatestAddin, 55)
    Add-RibbonCustomUi -PackagePath $tempLatestAddin
    Copy-Item -Path $tempLatestAddin -Destination $LatestAddin -Force

    $publishStage = "sanitize versioned add-in API credentials"
    Clear-WorkbookStoredApiSecrets -Workbook $workbook

    $publishStage = "save sanitized versioned add-in"
    $workbook.SaveAs($tempReleaseAddin, 55)
    $workbook.Close($false)
    $workbook = $null

    Add-RibbonCustomUi -PackagePath $tempReleaseAddin

    Copy-Item -Path $tempReleaseAddin -Destination $ReleaseAddin -Force
    Remove-Item -LiteralPath $tempLatestAddin -Force
    Remove-Item -LiteralPath $tempReleaseAddin -Force
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) | Out-Null } catch {} }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    if (Test-Path -LiteralPath $ImportModulePath) { Remove-Item -LiteralPath $ImportModulePath -Force }
    if ($tempLatestAddin -and (Test-Path -LiteralPath $tempLatestAddin)) { Remove-Item -LiteralPath $tempLatestAddin -Force }
    if ($tempReleaseAddin -and (Test-Path -LiteralPath $tempReleaseAddin)) { Remove-Item -LiteralPath $tempReleaseAddin -Force }
}

if ($wasDirtyBeforePublish) { $commit = "$commit-dirty" }

function Copy-OrderedDictionary {
    param($Source)

    $copy = [ordered]@{}
    foreach ($key in $Source.Keys) {
        $copy[$key] = $Source[$key]
    }
    return $copy
}

$versionedArtifactsSanitized = $true
$latestMayContainLocalCredentials = $true

$baseManifest = [ordered]@{
    name = "Data Editing Macro Toolkit"
    repository = "data-editing-macro-toolkit"
    license = "data-editing-macro-toolkit Non-Commercial Software Use License"
    licenseFile = "LICENSE.md"
    copyright = "Copyright (c) 2026 Jung Suhyun (정수현). All rights reserved."
    contact = "elbr1ll1antant@gmail.com"
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
    apiKeyStorage = "API credentials are read from the current Windows user profile via VBA SaveSetting/GetSetting."
    apiKeyReleasePolicy = "Only release/latest artifacts may keep workbook-scoped local API credentials. Versioned release artifacts are sanitized before save."
    notes = "Install the .xlam file as an Excel add-in. It adds a 데이터 툴킷 ribbon tab next to View."
}

$versionManifest = Copy-OrderedDictionary $baseManifest
$versionManifest["credentialMode"] = "sanitized-versioned"
$versionManifest["versionedArtifactsSanitized"] = $versionedArtifactsSanitized
$versionManifest["latestMayContainLocalCredentials"] = $false

$latestManifest = Copy-OrderedDictionary $baseManifest
$latestManifest["addin"] = $LatestAddinName
$latestManifest["workbook"] = $LatestWorkbookName
$latestManifest["installFile"] = $LatestAddinName
$latestManifest["credentialMode"] = "latest-local"
$latestManifest["versionedArtifactsSanitized"] = $false
$latestManifest["latestMayContainLocalCredentials"] = $latestMayContainLocalCredentials

$versionManifest | ConvertTo-Json | Set-Content -Encoding UTF8 -Path (Join-Path $VersionDir "manifest.json")
$latestManifest | ConvertTo-Json | Set-Content -Encoding UTF8 -Path (Join-Path $LatestDir "manifest.json")

Write-Output "Published add-in: $ReleaseAddin"
Write-Output "Published workbook: $ReleaseWorkbook"
Write-Output "Updated latest add-in: $LatestAddin"

