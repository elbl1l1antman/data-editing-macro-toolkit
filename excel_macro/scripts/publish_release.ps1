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
        </group>
        <group id="grpOutlierTools" label="이상치 탐지">
          <button id="btnOutlierSetup" label="설정" size="large" imageMso="ControlsGallery" onAction="Ribbon_OutlierDetectSetup" screentip="이상치 탐지 설정" supertip="이상치 탐지 및 대체 기준을 지정하는 설정 시트를 만듭니다." />
          <button id="btnOutlierRun" label="실행" imageMso="CalculateNow" onAction="Ribbon_OutlierDetectRun" screentip="이상치 탐지 실행" supertip="설정값에 따라 이상치를 탐지하고 대체 결과 시트를 생성합니다." />
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
            if ($component.Type -eq 1 -and ($component.Name -eq "MacroToolkit" -or $component.Name -like "Module*")) {
                $workbook.VBProject.VBComponents.Remove($component)
            }
        }
        $imported = $workbook.VBProject.VBComponents.Import($ImportModulePath)
        $imported.Name = "MacroToolkit"
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

