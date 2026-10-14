param(
    [string]$OutputPath,
    [string]$ModulePath
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

if (-not $OutputPath) {
    $OutputPath = Join-Path $ProjectRoot "dev\DataEditingMacroToolkit_dev.xlsm"
}

function Get-ToolkitModulePaths {
    param([string]$Root)

    return @(Get-ChildItem -LiteralPath (Join-Path $Root "src") -Filter "*.bas" -File | Sort-Object @{ Expression = { if ($_.Name -eq "MacroToolkit.bas") { 0 } else { 1 } } }, Name | Select-Object -ExpandProperty FullName)
}

$ModulePaths = if ($ModulePath) { @($ModulePath) } else { @(Get-ToolkitModulePaths -Root $ProjectRoot) }
$ModulePath = $ModulePaths | Where-Object { [System.IO.Path]::GetFileName($_) -eq "MacroToolkit.bas" } | Select-Object -First 1
if (-not $ModulePath) { throw "MacroToolkit.bas not found under src." }
foreach ($sourceModulePath in $ModulePaths) {
    if (-not (Test-Path -LiteralPath $sourceModulePath)) { throw "VBA module not found: $sourceModulePath" }
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null

$ToolkitVersion = "unknown"
$ModuleText = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
if ($ModuleText -match 'TOOLKIT_VERSION\s+As\s+String\s*=\s*"([^"]+)"') {
    $ToolkitVersion = $Matches[1]
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "MacroToolkit"
    $sheet.Range("A1").Value = "데이터 편집 매크로 툴킷"
    $sheet.Range("A2").Value = "버전"
    $sheet.Range("B2").Value = $ToolkitVersion
    $sheet.Range("A4").Value = "설치 상태"
    $sheet.Range("B4").Value = "VBA 모듈 자동 삽입 대기"
    $sheet.Range("A6").Value = "설치 마무리 방법"
    $sheet.Range("A7").Value = "1. Excel에서 'VBA 프로젝트 개체 모델 접근 신뢰' 옵션을 켭니다."
    $sheet.Range("A8").Value = "2. 이 스크립트를 다시 실행하거나, VBA 편집기에서 src 폴더의 BAS 모듈을 가져옵니다."
    $sheet.Range("A10").Value = "포함 매크로"
    $sheet.Range("A11").Value = "ConvertTextNumbersInSelection"
    $sheet.Range("A12").Value = "ConvertTextNumbersInUsedRange"
    $sheet.Range("A13").Value = "ConvertTextNumbersInColumns"
    $sheet.Range("A14").Value = "NormalizePhonesInSelection"
    $sheet.Range("A15").Value = "NormalizePhonesInColumns"
    $sheet.Range("A16").Value = "OutlierDetect_Setup"
    $sheet.Range("A17").Value = "OutlierDetect_Run"
    $sheet.Range("A18").Value = "SPSS_DeleteSelectedBanner"
    $sheet.Columns("A:B").AutoFit()

    try {
        if ($null -ne $workbook.VBProject) {
            foreach ($sourceModulePath in $ModulePaths) {
                $workbook.VBProject.VBComponents.Import($sourceModulePath) | Out-Null
            }
            $excel.Run("'" + $workbook.Name + "'!Toolkit_SetupWorkbook")
            $workbook.Worksheets.Item("_MacroToolkit").Visible = 2
            $sheet.Range("B4").Value = "VBA 모듈 삽입 완료"
        }
    }
    catch {
        $sheet.Range("B4").Value = "Excel 보안 설정으로 VBA 모듈 자동 삽입 차단"
    }

    $OutputDir = Split-Path -Parent $OutputPath
    $TempOutput = Join-Path $OutputDir ("_tmp_" + [System.IO.Path]::GetFileName($OutputPath))
    if (Test-Path -LiteralPath $TempOutput) { Remove-Item -LiteralPath $TempOutput -Force }

    $workbook.SaveAs($TempOutput, 52)
    $workbook.Close($true)
    $workbook = $null
    Copy-Item -Path $TempOutput -Destination $OutputPath -Force
    Remove-Item -LiteralPath $TempOutput -Force
    Write-Output "Created: $OutputPath"

    $DocScript = Join-Path $ScriptDir "update_workbook_docs.ps1"
    if (Test-Path -LiteralPath $DocScript) {
        & $DocScript -WorkbookPath $OutputPath -ModulePath $ModulePath
    }
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) | Out-Null } catch {}
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}


