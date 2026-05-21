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
    $form.Properties.Item("Width").Value = 620
    $form.Properties.Item("Height").Value = 430

    $lblHelp = $designer.Controls.Add("Forms.Label.1", "lblHelp", $true)
    $lblHelp.Caption = "삭제할 표와 배너를 선택하고, 배너는 위/아래 버튼으로 순서를 조정하세요."
    $lblHelp.Left = 12
    $lblHelp.Top = 10
    $lblHelp.Width = 580
    $lblHelp.Height = 18

    $lblTables = $designer.Controls.Add("Forms.Label.1", "lblTables", $true)
    $lblTables.Caption = "삭제할 표"
    $lblTables.Left = 12
    $lblTables.Top = 34
    $lblTables.Width = 260
    $lblTables.Height = 16

    $lstTables = $designer.Controls.Add("Forms.ListBox.1", "lstTables", $true)
    $lstTables.Left = 12
    $lstTables.Top = 52
    $lstTables.Width = 280
    $lstTables.Height = 292
    $lstTables.ColumnCount = 2
    $lstTables.ColumnWidths = "240 pt;0 pt"
    $lstTables.MultiSelect = 1

    $lblBanners = $designer.Controls.Add("Forms.Label.1", "lblBanners", $true)
    $lblBanners.Caption = "삭제/이동할 배너"
    $lblBanners.Left = 312
    $lblBanners.Top = 34
    $lblBanners.Width = 220
    $lblBanners.Height = 16

    $lstBanners = $designer.Controls.Add("Forms.ListBox.1", "lstBanners", $true)
    $lstBanners.Left = 312
    $lstBanners.Top = 52
    $lstBanners.Width = 220
    $lstBanners.Height = 292
    $lstBanners.ColumnCount = 2
    $lstBanners.ColumnWidths = "180 pt;0 pt"
    $lstBanners.MultiSelect = 1

    $cmdUp = $designer.Controls.Add("Forms.CommandButton.1", "cmdBannerUp", $true)
    $cmdUp.Caption = "위로"
    $cmdUp.Left = 542
    $cmdUp.Top = 84
    $cmdUp.Width = 54
    $cmdUp.Height = 24

    $cmdDown = $designer.Controls.Add("Forms.CommandButton.1", "cmdBannerDown", $true)
    $cmdDown.Caption = "아래로"
    $cmdDown.Left = 542
    $cmdDown.Top = 114
    $cmdDown.Width = 54
    $cmdDown.Height = 24

    $lblTip = $designer.Controls.Add("Forms.Label.1", "lblTip", $true)
    $lblTip.Caption = "체크된 항목은 삭제됩니다. 순서만 바꿀 배너는 선택하지 말고 위/아래만 사용하세요."
    $lblTip.Left = 12
    $lblTip.Top = 354
    $lblTip.Width = 580
    $lblTip.Height = 28

    $cmdOk = $designer.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    $cmdOk.Caption = "실행"
    $cmdOk.Left = 414
    $cmdOk.Top = 386
    $cmdOk.Width = 82
    $cmdOk.Height = 26

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 510
    $cmdCancel.Top = 386
    $cmdCancel.Width = 82
    $cmdCancel.Height = 26

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

function Add-RibbonCustomUi {
    param([string]$PackagePath)

    $customUiXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabDataEditingToolkit" label="데이터 툴킷" insertAfterMso="TabView">
        <group id="grpNumberTools" label="숫자 변환">
          <button id="btnConvertSelection" label="선택 영역" size="large" imageMso="NumberFormatGeneral" onAction="Ribbon_ConvertTextNumbersInSelection" screentip="선택 영역 숫자 변환" supertip="선택한 셀의 텍스트 형식 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnConvertUsedRange" label="현재 시트" imageMso="TableSelect" onAction="Ribbon_ConvertTextNumbersInUsedRange" screentip="현재 시트 숫자 변환" supertip="활성 시트의 사용 범위 전체에서 텍스트 형식 숫자를 실제 숫자값으로 변환합니다." />
          <button id="btnConvertRange" label="범위 선택" imageMso="SelectCell" onAction="Ribbon_ConvertTextNumbersInColumns" screentip="범위 지정 숫자 변환" supertip="마우스로 선택하거나 직접 입력한 범위의 텍스트 형식 숫자를 실제 숫자값으로 변환합니다." />
        </group>
        <group id="grpPhoneTools" label="전화번호">
          <button id="btnPhoneSelection" label="선택 영역" size="large" imageMso="ContactPictureMenu" onAction="Ribbon_NormalizePhonesInSelection" screentip="선택 영역 전화번호 정리" supertip="선택한 셀의 한국 전화번호 형식을 하이픈 포함 표준 형식으로 정규화합니다." />
          <button id="btnPhoneRange" label="범위 선택" imageMso="AddressBook" onAction="Ribbon_NormalizePhonesInColumns" screentip="범위 지정 전화번호 정리" supertip="마우스로 선택하거나 직접 입력한 범위의 전화번호를 정규화합니다." />
        </group>
        <group id="grpBusinessNoTools" label="사업자등록">
          <button id="btnBizApiKey" label="인증키 설정" imageMso="ServerConnection" onAction="Ribbon_BizNoSetApiKey" screentip="사업자등록 API 인증키 설정" supertip="공공데이터포털 국세청 사업자등록정보 API 인증키를 저장하거나 변경합니다." />
          <button id="btnBizMerge" label="3열 병합" imageMso="CellsMerge" onAction="Ribbon_BizNoMergeThreeColumns" screentip="사업자등록번호 병합" supertip="3개 열로 나뉜 사업자등록번호를 000-00-00000 형식의 새 열로 병합합니다." />
          <button id="btnBizStatus" label="상태조회" size="large" imageMso="DataRefreshAll" onAction="Ribbon_BizNoQueryStatus" screentip="사업자등록 상태조회" supertip="선택한 사업자등록번호 범위를 국세청 상태조회 API로 조회하고 오른쪽에 결과 열을 추가합니다." />
        </group>
        <group id="grpSpssTableTools" label="SPSS 표 정리">
          <button id="btnSpssDeleteBanner" label="선택 배너 삭제" size="large" imageMso="DeleteTableRows" onAction="Ribbon_SPSSDeleteSelectedBanner" screentip="SPSS 배너 삭제" supertip="선택한 ◐ 가로 배너의 열 블록 또는 ▣ 세로 배너의 행 블록을 삭제하고 남은 셀을 당겨 붙입니다." />
          <button id="btnSpssRebuildFromIndex" label="목차 기반 재구성" imageMso="CreateReport" onAction="Ribbon_SPSSRebuildTablesFromIndex" screentip="목차 기반 표/배너 재구성" supertip="첫 시트의 목차 링크와 다음 시트의 집계표를 탐지해 삭제할 표, 삭제할 배너, 배너 순서를 지정한 뒤 신규 목차와 표 시트를 생성합니다." />
        </group>
        <group id="grpOutlierTools" label="이상치 탐지">
          <button id="btnOutlierRun" label="분석 열 선택" size="large" imageMso="ConditionalFormattingHighlightCellsRules" onAction="Ribbon_OutlierDetectRun" screentip="이상치 확인 범위 선택" supertip="범위를 직접 지정하면 범위 검증 후 기초통계량표를 먼저 만들고, 오른쪽에 Q1/Q3 기준값 열과 원본 검토복사 열을 추가해 조건부 서식으로 표시합니다." />
          <button id="btnOutlierSetup" label="고급 설정" imageMso="ControlsGallery" onAction="Ribbon_OutlierDetectSetup" screentip="이상치 탐지 고급 설정" supertip="설정 시트를 사용하는 기존 방식의 이상치 탐지 및 결과 시트 생성을 준비합니다." />
        </group>
        <group id="grpToolkitInfo" label="정보">
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
                ($component.Type -eq 3 -and $component.Name -eq "frmSpssRebuildOptions")) {
                $workbook.VBProject.VBComponents.Remove($component)
            }
        }
        $imported = $workbook.VBProject.VBComponents.Import($ImportModulePath)
        $imported.Name = "MacroToolkit"
        Add-SpssRebuildUserForm -Workbook $workbook
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

