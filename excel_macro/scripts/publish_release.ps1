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

function Add-SpssRebuildUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmSpssRebuildOptions"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "SPSS 표/배너 재구성"
    $form.Properties.Item("Width").Value = 700
    $form.Properties.Item("Height").Value = 470

    $lblHelp = $designer.Controls.Add("Forms.Label.1", "lblHelp", $true)
    $lblHelp.Caption = "삭제할 표와 배너를 선택하세요. Shift+클릭은 연속 선택, Ctrl+클릭은 개별 추가 선택입니다."
    $lblHelp.Left = 12
    $lblHelp.Top = 10
    $lblHelp.Width = 660
    $lblHelp.Height = 28

    $lblTables = $designer.Controls.Add("Forms.Label.1", "lblTables", $true)
    $lblTables.Caption = "삭제할 표"
    $lblTables.Left = 12
    $lblTables.Top = 44
    $lblTables.Width = 310
    $lblTables.Height = 16

    $lstTables = $designer.Controls.Add("Forms.ListBox.1", "lstTables", $true)
    $lstTables.Left = 12
    $lstTables.Top = 64
    $lstTables.Width = 320
    $lstTables.Height = 306
    $lstTables.ColumnCount = 2
    $lstTables.ColumnWidths = "290 pt;0 pt"
    $lstTables.MultiSelect = 2

    $lblBanners = $designer.Controls.Add("Forms.Label.1", "lblBanners", $true)
    $lblBanners.Caption = "삭제할 배너"
    $lblBanners.Left = 352
    $lblBanners.Top = 44
    $lblBanners.Width = 310
    $lblBanners.Height = 16

    $lstBanners = $designer.Controls.Add("Forms.ListBox.1", "lstBanners", $true)
    $lstBanners.Left = 352
    $lstBanners.Top = 64
    $lstBanners.Width = 320
    $lstBanners.Height = 306
    $lstBanners.ColumnCount = 2
    $lstBanners.ColumnWidths = "290 pt;0 pt"
    $lstBanners.MultiSelect = 2

    $cmdUp = $designer.Controls.Add("Forms.CommandButton.1", "cmdBannerUp", $true)
    $cmdUp.Caption = "위로"
    $cmdUp.Visible = $false
    $cmdUp.Left = 542
    $cmdUp.Top = 84
    $cmdUp.Width = 54
    $cmdUp.Height = 24

    $cmdDown = $designer.Controls.Add("Forms.CommandButton.1", "cmdBannerDown", $true)
    $cmdDown.Caption = "아래로"
    $cmdDown.Visible = $false
    $cmdDown.Left = 542
    $cmdDown.Top = 114
    $cmdDown.Width = 54
    $cmdDown.Height = 24

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "선택된 항목은 삭제됩니다. 원본 시트를 복사한 뒤 결과를 만들며, BOT/TOP/MID 계열은 세로 배너 삭제 버튼에서 별도로 처리하세요."
    $lblTip.Left = 12
    $lblTip.Top = 382
    $lblTip.Width = 660
    $lblTip.Height = 34

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 486
    $cmdOk.Top = 424
    $cmdOk.Width = 86
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 586
    $cmdCancel.Top = 424
    $cmdCancel.Width = 86
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean
Public DeleteTablesText As String
Public DeleteBannersText As String
Public BannerOrderText As String

Private Sub UserForm_Initialize()
    Cancelled = True
End Sub

Public Sub InitializeOptions(ByVal tables As Object, ByVal banners As Object)
    Dim i As Long
    Dim item As Variant

    lstTables.Clear
    For i = 1 To tables.Count
        item = tables(i)
        lstTables.AddItem CStr(i) & ". " & CStr(item(0))
        lstTables.List(lstTables.ListCount - 1, 1) = CStr(i)
    Next i

    lstBanners.Clear
    For i = 1 To banners.Count
        lstBanners.AddItem CStr(i) & ". " & CStr(banners(i))
        lstBanners.List(lstBanners.ListCount - 1, 1) = CStr(i)
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
    DeleteTablesText = SelectedIds(lstTables)
    DeleteBannersText = SelectedIds(lstBanners)
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
'@

    $form.CodeModule.AddFromString($formCode)
}

function Add-RangeSelectUserForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "frmRangeSelectOptions"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "범위 선택"
    $form.Properties.Item("Width").Value = 520
    $form.Properties.Item("Height").Value = 240

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Left = 12
    $lblTitle.Top = 10
    $lblTitle.Width = 480
    $lblTitle.Height = 18

    $lblGuide = $designer.Controls.Add("Forms.Label.1", "lblGuide", $true)
    $lblGuide.Left = 12
    $lblGuide.Top = 34
    $lblGuide.Width = 480
    $lblGuide.Height = 66

    $txtRange = $designer.Controls.Add("Forms.TextBox.1", "txtRange", $true)
    $txtRange.Left = 12
    $txtRange.Top = 112
    $txtRange.Width = 350
    $txtRange.Height = 22

    $cmdPick = $designer.Controls.Add("Forms.CommandButton.1", "cmdPickRange", $true)
    $cmdPick.Caption = "현재 선택 사용"
    $cmdPick.Left = 374
    $cmdPick.Top = 110
    $cmdPick.Width = 118
    $cmdPick.Height = 26

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 306
    $cmdOk.Top = 166
    $cmdOk.Width = 86
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 406
    $cmdCancel.Top = 166
    $cmdCancel.Width = 86
    $cmdCancel.Height = 28

    $formCode = @'
Option Explicit

Public Cancelled As Boolean
Public Done As Boolean
Public SelectedRange As Object

Private Sub UserForm_Initialize()
    Cancelled = True
    Done = False
End Sub

Public Sub InitializeOptions(ByVal titleText As String, ByVal guideText As String, ByVal defaultAddress As String)
    Me.Caption = titleText
    lblTitle.Caption = titleText
    lblGuide.Caption = guideText & vbCrLf & "범위를 선택한 뒤 이 창의 [실행]을 누르거나 [현재 선택 사용]을 누르세요."
    txtRange.Value = defaultAddress
    txtRange.Tag = defaultAddress
End Sub

Private Sub cmdPickRange_Click()
    On Error Resume Next
    If TypeName(Application.Selection) = "Range" Then
        Set SelectedRange = Application.Selection
        txtRange.Value = SelectedRange.Address(External:=True)
        txtRange.Tag = txtRange.Value
    End If
    On Error GoTo 0
End Sub

Private Sub cmdOK_Click()
    Dim typedAddress As String
    typedAddress = Trim$(txtRange.Value)

    On Error Resume Next
    If Len(typedAddress) > 0 And typedAddress <> CStr(txtRange.Tag) Then
        Set SelectedRange = Application.Range(typedAddress)
    ElseIf TypeName(Application.Selection) = "Range" Then
        Set SelectedRange = Application.Selection
    ElseIf Len(typedAddress) > 0 Then
        Set SelectedRange = Application.Range(typedAddress)
    End If
    On Error GoTo 0

    If SelectedRange Is Nothing Then
        MsgBox "엑셀 시트에서 처리할 범위를 먼저 선택하거나 주소를 입력하세요.", vbExclamation, Me.Caption
        Exit Sub
    End If

    Cancelled = False
    Done = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Done = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancelled = True
    Done = True
End Sub
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
        @("chkSpellAlphabetKorean", "알파벳을 한글 읽기로 변환", 250, 146, $true)
    )

    foreach ($item in $labels) {
        $chk = $designer.Controls.Add("Forms.CheckBox.1", $item[0], $true)
        $chk.Caption = $item[1]
        $chk.Left = [int]$item[2]
        $chk.Top = [int]$item[3]
        $chk.Width = 220
        $chk.Height = 20
        $chk.Value = [bool]$item[4]
    }

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "기본값은 제공된 기업이름정리.xlsm의 체크 상태를 따릅니다. 결과는 원본 오른쪽 복사본 열에 생성됩니다."
    $lblTip.Left = 12
    $lblTip.Top = 214
    $lblTip.Width = 460
    $lblTip.Height = 42

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 286
    $cmdOk.Top = 330
    $cmdOk.Width = 86
    $cmdOk.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 386
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

function Add-RibbonCustomUi {
    param([string]$PackagePath)

    $customUiXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabDataEditingToolkit" label="데이터 툴킷" insertAfterMso="TabView">
        <group id="grpNumberTools" label="값 정리">
          <button id="btnConvertRange" label="숫자 변환" size="large" imageMso="NumberFormatGeneral" onAction="Ribbon_ConvertTextNumbersInColumns" screentip="선택 범위 숫자 변환" supertip="선택한 범위에서 텍스트 형식 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnConvertSheet" label="시트 숫자 변환" imageMso="TableSelect" onAction="Ribbon_ConvertTextNumbersInUsedRange" screentip="시트 전체 숫자 변환" supertip="현재 시트에서 실제 데이터가 있는 마지막 행과 열까지만 잡아 텍스트 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnPhoneRange" label="전화번호 정리" size="large" imageMso="ContactPictureMenu" onAction="Ribbon_NormalizePhonesInColumns" screentip="전화번호 하이픈 정리" supertip="선택 범위 오른쪽에 복사본 열을 만들고 한국 전화번호를 하이픈 포함 표준 형식으로 정리합니다." />
          <button id="btnPhoneDashless" label="전화번호 숫자만" imageMso="ClearFormats" onAction="Ribbon_RemovePhoneDashes" screentip="전화번호 숫자만 남기기" supertip="선택한 전화번호 범위 오른쪽에 복사본 열을 만들고 대시와 구분기호를 제거해 숫자만 남깁니다." />
          <button id="btnCompanyNameClean" label="기업명 정리" size="large" imageMso="TextBoxInsert" onAction="Ribbon_CompanyNameClean" screentip="기업체명/사업체명 정리" supertip="선택 범위 오른쪽에 복사본 열을 만들고 공백, 회사형태 표기, 알파벳 한글 읽기 등 선택한 규칙으로 기업체명/사업체명을 정리합니다." />
        </group>
        <group id="grpBusinessNoTools" label="사업자등록">
          <button id="btnBizApiKey" label="인증키 설정" imageMso="ServerConnection" onAction="Ribbon_BizNoSetApiKey" screentip="사업자등록 API 인증키 설정" supertip="공공데이터포털 국세청 사업자등록정보 API 인증키를 저장하거나 변경합니다." />
          <button id="btnBizDashless" label="번호 숫자만" imageMso="ClearFormats" onAction="Ribbon_RemoveBizNoDashes" screentip="사업자등록번호 숫자만 남기기" supertip="선택한 사업자등록번호 범위 오른쪽에 복사본 열을 만들고 대시와 구분기호를 제거해 숫자 10자리 형태로 정리합니다." />
          <button id="btnBizMerge" label="3열 병합" imageMso="CellsMerge" onAction="Ribbon_BizNoMergeThreeColumns" screentip="사업자등록번호 3열 병합" supertip="3개 열로 나뉜 사업자등록번호를 000-00-00000 형식의 새 열로 병합합니다." />
          <button id="btnBizStatus" label="상태조회" size="large" imageMso="DataRefreshAll" onAction="Ribbon_BizNoQueryStatus" screentip="사업자등록 상태조회" supertip="선택한 사업자등록번호 범위를 국세청 API로 조회하고 오른쪽에 결과 열을 추가합니다." />
        </group>
        <group id="grpSpssTableTools" label="SPSS 표 정리">
          <button id="btnSpssCreateIndex" label="목차 생성" imageMso="TableOfContentsGallery" onAction="Ribbon_SPSSCreateIndexForActiveSheet" screentip="SPSS 표 목차 생성" supertip="목차가 없는 집계표 시트에서 [로 시작하는 표 제목 행을 찾아 새 목차 시트와 링크를 생성합니다." />
          <button id="btnSpssRebuildFromIndex" label="표/배너 재구성" size="large" imageMso="CreateReport" onAction="Ribbon_SPSSRebuildTablesFromIndex" screentip="SPSS 표/배너 재구성" supertip="표와 배너 목록을 선택해 원본 집계표 시트의 복사본에 삭제 내용을 반영하고 새 목차 링크를 생성합니다." />
          <button id="btnSpssDeleteVertical" label="세로 배너 삭제" imageMso="DeleteRows" onAction="Ribbon_SPSSDeleteVerticalBanners" screentip="세로 배너 삭제" supertip="원본 집계표 시트를 복사한 뒤 ▣ 세로 배너와 BOT/TOP/MID 계열 세로형 배너 행을 삭제하고 새 목차를 생성합니다." />
        </group>
        <group id="grpOutlierTools" label="이상치 탐지">
          <button id="btnOutlierRun" label="분석 열 선택" size="large" imageMso="ConditionalFormattingHighlightCellsRules" onAction="Ribbon_OutlierDetectRun" screentip="이상치 확인 범위 선택" supertip="범위를 직접 지정하면 범위 검증 후 기초통계량표를 먼저 만들고, 오른쪽에 Q1/Q3 기준값 열과 원본 검토복사 열을 추가해 조건부 서식으로 표시합니다." />
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
                ($component.Type -eq 3 -and ($component.Name -eq "frmSpssRebuildOptions" -or $component.Name -eq "frmRangeSelectOptions" -or $component.Name -eq "frmCompanyNameOptions"))) {
                $workbook.VBProject.VBComponents.Remove($component)
            }
        }
        $imported = $workbook.VBProject.VBComponents.Import($ImportModulePath)
        $imported.Name = "MacroToolkit"
        Add-SpssRebuildUserForm -Workbook $workbook
        Add-RangeSelectUserForm -Workbook $workbook
        Add-CompanyNameOptionsUserForm -Workbook $workbook
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

